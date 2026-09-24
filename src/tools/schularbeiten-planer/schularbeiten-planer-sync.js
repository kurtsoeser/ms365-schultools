/**
 * Sync fixierter Schularbeiten → SharePoint-Liste „Schultermine“.
 */
import { toIsoDateOnly } from './schularbeiten-planer-logic.js';

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All'
];

function G() {
    const api = typeof window !== 'undefined' ? window.ms365SpoGraph : null;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen.');
    return api;
}

async function token() {
    return await G().getGraphToken(SCOPES);
}

async function findListByTitle(tok, siteId, listTitle) {
    const title = String(listTitle || '').trim() || 'Schultermine';
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
    return ((data && data.value) || [])[0] || null;
}

/**
 * Marker in Info-Feld für Idempotenz.
 * @param {string} schularbeitId
 */
export function saMarker(schularbeitId) {
    return '[SA:' + String(schularbeitId || '').trim() + ']';
}

/**
 * @param {object} sa
 * @param {{ fach?: string, klasse?: string }} [labels]
 */
export function buildSchulterminFields(sa, labels) {
    const datum = toIsoDateOnly(sa.datum);
    const fach = (labels && labels.fach) || sa.fachCode || 'Fach';
    const klasse = (labels && labels.klasse) || sa.klasseCode || '';
    const title =
        fach + (klasse ? ' · ' + klasse : '') + (sa.thema ? ' – ' + String(sa.thema).slice(0, 80) : '');
    const marker = saMarker(sa.schularbeitId);
    const info = [marker, 'Schularbeit', sa.lehrerCode ? 'Lehrer: ' + sa.lehrerCode : '', sa.dauerMinuten ? sa.dauerMinuten + ' Min.' : '']
        .filter(Boolean)
        .join(' · ');
    return {
        Title: title.slice(0, 250),
        Beginn: datum,
        Ende: datum,
        Kategorie: 'Prüfung',
        Info: info,
        ZeitraumText: datum || '',
        AllDay: true,
        SyncStatus: 'pending'
    };
}

/**
 * Findet vorhandenes Schultermine-Item anhand des SA-Markers.
 * @param {string} tok
 * @param {string} siteId
 * @param {string} listId
 * @param {string} schularbeitId
 */
async function findItemByMarker(tok, siteId, listId, schularbeitId) {
    const marker = saMarker(schularbeitId);
    let path =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(listId) +
        '/items?$expand=fields&$top=100';
    while (path) {
        const data = await G().graphJson('GET', path, tok, undefined, 'v1.0');
        const rows = (data && data.value) || [];
        for (let i = 0; i < rows.length; i++) {
            const info = String((rows[i].fields && rows[i].fields.Info) || '');
            if (info.indexOf(marker) !== -1) return rows[i];
        }
        path = data && data['@odata.nextLink'] ? data['@odata.nextLink'] : '';
    }
    return null;
}

/**
 * @param {{ siteId: string, webUrl: string }} ctx Planer-Context (gleiche Site)
 * @param {object} sa
 * @param {{ listTitle?: string, fachLabel?: string, klasseLabel?: string }} [opts]
 * @returns {Promise<{ itemId: string, created: boolean }>}
 */
export async function upsertSchulterminFromSchularbeit(ctx, sa, opts) {
    if (!sa || !sa.schularbeitId) throw new Error('SchularbeitId fehlt für Schultermine-Sync.');
    const listTitle = (opts && opts.listTitle) || 'Schultermine';
    const tok = await token();
    const list = await findListByTitle(tok, ctx.siteId, listTitle);
    if (!list || !list.id) {
        throw new Error('Liste „' + listTitle + '“ nicht gefunden. Bitte zuerst anlegen.');
    }

    const fields = buildSchulterminFields(sa, {
        fach: opts && opts.fachLabel,
        klasse: opts && opts.klasseLabel
    });

    const existing = await findItemByMarker(tok, ctx.siteId, list.id, sa.schularbeitId);
    if (existing && existing.id) {
        await G().graphJson(
            'PATCH',
            G().graphPathSite(ctx.siteId) +
                '/lists/' +
                encodeURIComponent(list.id) +
                '/items/' +
                encodeURIComponent(existing.id) +
                '/fields',
            tok,
            fields,
            'v1.0'
        );
        return { itemId: String(existing.id), created: false };
    }

    const created = await G().graphJson(
        'POST',
        G().graphPathSite(ctx.siteId) + '/lists/' + encodeURIComponent(list.id) + '/items',
        tok,
        { fields },
        'v1.0'
    );
    return { itemId: created && created.id != null ? String(created.id) : '', created: true };
}
