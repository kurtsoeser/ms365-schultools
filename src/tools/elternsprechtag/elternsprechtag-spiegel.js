/**
 * IT schreibt Bookings-Termine in die SharePoint-Liste.
 * Lehrkräfte lesen danach nur die Zeile mit der eigenen Mail.
 */

import { fetchAllPages, getGraphToken, graphJson, sleep } from '../../shared/graph-client.js';
import { toGraphColumnBody } from '../../shared/spo-graph-column.js';
import { createListEnsureHelpers } from '../sharepoint/sharepoint-liste-ensure-core.js';
import {
    ELTERNSPRECHTAG_LIST_COLUMNS,
    ELTERNSPRECHTAG_LIST_TITLE,
    dayWasPublished,
    fieldsForRow,
    itemIdsForDate,
    rowsFromAppointments,
    datesFromItems,
    latestUpdate,
    slotsFromSpiegelItems,
    windowFromItems,
    withServiceWindow
} from './elternsprechtag-spiegel-logic.js';

const READ_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.Read.All'
];

const WRITE_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function spo() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api || typeof api.resolveSiteFromWebUrl !== 'function' || typeof api.graphPathSite !== 'function') {
        throw new Error('SharePoint-Hilfen nicht geladen.');
    }
    return api;
}

/**
 * @param {string} token
 * @param {string} siteId
 */
async function findList(token, siteId) {
    const path =
        spo().graphPathSite(siteId) +
        "/lists?$filter=" +
        encodeURIComponent("displayName eq '" + ELTERNSPRECHTAG_LIST_TITLE.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await graphJson('GET', path, token);
    return ((data && data.value) || [])[0] || null;
}

/**
 * @param {string} token
 * @param {string} siteId
 * @param {string} listId
 */
async function listItems(token, siteId, listId) {
    const path =
        spo().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=200';
    const { items } = await fetchAllPages(token, path, { maxItems: 4000, maxPages: 25 });
    return items || [];
}

/**
 * @param {string} token
 * @param {string} businessId
 * @param {string} eventDate
 * @param {(msg: string) => void} log
 */
async function loadAppointments(token, businessId, eventDate, log) {
    const start = encodeURIComponent(eventDate + 'T00:00:00.0000000');
    const end = encodeURIComponent(eventDate + 'T23:59:59.0000000');
    const base = '/solutions/bookingBusinesses/' + encodeURIComponent(businessId);
    log('Lade Buchungen ' + eventDate + ' …');
    const { items } = await fetchAllPages(
        token,
        base + '/calendarView?start=' + start + '&end=' + end,
        { maxItems: 2000, maxPages: 20 }
    );
    const enriched = [];
    for (let i = 0; i < items.length; i++) {
        const ap = items[i];
        const hasCustomer =
            (Array.isArray(ap && ap.customers) && ap.customers.length) ||
            (ap && (ap.customerEmailAddress || ap.customerName));
        if (hasCustomer || !ap || !ap.id) {
            enriched.push(ap);
            continue;
        }
        try {
            const detail = await graphJson(
                'GET',
                base + '/appointments/' + encodeURIComponent(ap.id),
                token
            );
            enriched.push(detail || ap);
        } catch {
            enriched.push(ap);
        }
        await sleep(60);
    }
    return enriched;
}

/**
 * @param {string} token
 * @param {string} businessId
 */
async function loadStaff(token, businessId) {
    const base = '/solutions/bookingBusinesses/' + encodeURIComponent(businessId);
    const { items } = await fetchAllPages(token, base + '/staffMembers', { maxItems: 500, maxPages: 10 });
    const full = [];
    for (let i = 0; i < items.length; i++) {
        const member = items[i];
        if (member && member.emailAddress) {
            full.push(member);
            continue;
        }
        if (member && member.id) {
            try {
                const one = await graphJson(
                    'GET',
                    base + '/staffMembers/' + encodeURIComponent(member.id),
                    token
                );
                full.push(one && one.id ? one : member);
            } catch {
                full.push(member);
            }
            await sleep(60);
        }
    }
    return full;
}

/**
 * @param {{
 *   businessId: string,
 *   eventDate: string,
 *   siteUrl: string,
 *   serviceName?: string,
 *   bookingsToken: string,
 *   log?: (msg: string) => void
 * }} opts
 */
export async function publishSpiegelDay(opts) {
    const o = opts || {};
    const log = typeof o.log === 'function' ? o.log : function () {};
    const businessId = String(o.businessId || '').trim();
    const eventDate = String(o.eventDate || '').trim();
    const siteUrl = String(o.siteUrl || '').trim();
    if (!businessId) throw new Error('Bookings-Business-ID fehlt.');
    if (!/^\d{4}-\d{2}-\d{2}$/.test(eventDate)) throw new Error('Datum des Sprechtags fehlt.');
    if (!/^https:\/\//i.test(siteUrl)) throw new Error('Keine Schul-Website hinterlegt. In den Stammdaten die Planer-Website setzen.');
    const bookingsToken = String(o.bookingsToken || '').trim();
    if (!bookingsToken) throw new Error('Anmeldung für Bookings fehlt.');

    const appointments = await loadAppointments(bookingsToken, businessId, eventDate, log);
    const staff = await loadStaff(bookingsToken, businessId);
    const built = rowsFromAppointments(appointments, staff, {
        eventDate: eventDate,
        serviceName: o.serviceName || ''
    });
    const rows = withServiceWindow(built.rows, o.window, eventDate);
    log(
        'Buchungen: ' +
            appointments.length +
            ', Zeilen: ' +
            built.rows.length +
            (built.skippedNoTeacher ? ', ohne Lehrkraft-Mail: ' + built.skippedNoTeacher : '')
    );

    const spoToken = await getGraphToken(WRITE_SCOPES);
    log('Website ' + siteUrl);
    const site = await spo().resolveSiteFromWebUrl(spoToken, siteUrl);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('SharePoint-Website nicht gefunden.');
    const { ensureList } = createListEnsureHelpers(toGraphColumnBody);
    const list = await ensureList(spoToken, siteId, ELTERNSPRECHTAG_LIST_TITLE, ELTERNSPRECHTAG_LIST_COLUMNS, log);
    const listId = list && list.id ? String(list.id) : '';
    if (!listId) throw new Error('Liste fehlt.');

    const existing = await listItems(spoToken, siteId, listId);
    const removeIds = itemIdsForDate(existing, eventDate);
    log('Ersetze ' + removeIds.length + ' vorhandene Zeile(n) dieses Tages.');
    for (let i = 0; i < removeIds.length; i++) {
        await graphJson(
            'DELETE',
            spo().graphPathSite(siteId) +
                '/lists/' +
                encodeURIComponent(listId) +
                '/items/' +
                encodeURIComponent(removeIds[i]),
            spoToken
        );
        await sleep(40);
    }
    for (let i = 0; i < rows.length; i++) {
        await graphJson(
            'POST',
            spo().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/items',
            spoToken,
            { fields: fieldsForRow(rows[i]) }
        );
        if (i % 10 === 9) log('Geschrieben: ' + (i + 1) + ' / ' + rows.length);
        await sleep(40);
    }
    log('Fertig. ' + built.rows.length + ' Zeile(n) in „' + ELTERNSPRECHTAG_LIST_TITLE + '“.');
    return {
        written: built.rows.length,
        removed: removeIds.length,
        skippedNoTeacher: built.skippedNoTeacher,
        listUrl: (list && list.webUrl) || ''
    };
}

/**
 * @param {string} siteUrl
 * @param {string} eventDate
 * @param {string} selfEmail
 * @returns {Promise<{ published: boolean, slots: object[] }>}
 */
export async function loadSpiegelSlots(siteUrl, eventDate, selfEmail) {
    const url = String(siteUrl || '').trim();
    const day = String(eventDate || '').trim();
    if (!/^https:\/\//i.test(url) || !/^\d{4}-\d{2}-\d{2}$/.test(day)) {
        return { published: false, slots: [] };
    }
    const token = await getGraphToken(READ_SCOPES);
    const site = await spo().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) return { published: false, slots: [] };
    const list = await findList(token, siteId);
    if (!list || !list.id) return { published: false, slots: [] };
    const items = await listItems(token, siteId, String(list.id));
    return {
        published: dayWasPublished(items, day),
        slots: slotsFromSpiegelItems(items, selfEmail, day),
        window: windowFromItems(items, day),
        dates: datesFromItems(items),
        updatedAt: latestUpdate(items, day)
    };
}
