/**
 * SharePoint: sRDP-Anmeldeliste anlegen (Spalten, Ansichten, Rechte).
 */
import {
    STORAGE_KEY,
    PROFILE_IDS,
    getProfile,
    listTitleForYear,
    listTitleCandidates,
    buildAnmeldungColumns,
    toGraphColumnBody,
    defaultViewsForProfile,
    requiredColumnNamesForProfile
} from '../srdp-anmeldung/srdp-anmeldung-schema.js';
import {
    filterAbschlussklassen,
    suggestLfsChoices,
    teacherChoiceLabels,
    parseChoiceLines,
    buildItemTitle,
    buildPruefplanKurz
} from '../srdp-anmeldung/srdp-anmeldung-logic.js';
import {
    buildFieldFormatSpecs,
    buildClientFormCustomFormatterString
} from '../srdp-anmeldung/srdp-anmeldung-formatting.js';
import {
    entraGroupLogonName,
    isBroadSiteAudience,
    SPO_ROLE
} from '../../shared/stammdaten-sharepoint-sync-logic.js';

const SCOPES_GRAPH = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/Sites.ReadWrite.All',
    'https://graph.microsoft.com/Group.Read.All'
];

function G() {
    const api = window.ms365SpoGraph;
    if (!api) throw new Error('SharePoint-Graph-Hilfen nicht geladen (spo-graph-shared).');
    return api;
}

function $(id) {
    return document.getElementById(id);
}

function log(msg) {
    const el = $('srsLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + msg;
    el.scrollTop = el.scrollHeight;
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function loadWizardState() {
    try {
        const raw = localStorage.getItem(STORAGE_KEY);
        if (!raw) return {};
        const o = JSON.parse(raw);
        return o && typeof o === 'object' ? o : {};
    } catch {
        return {};
    }
}

function saveWizardState(partial) {
    const next = Object.assign({}, loadWizardState(), partial || {});
    try {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(next));
    } catch {
        /* ignore */
    }
    return next;
}

function loadStammdaten() {
    if (typeof window.ms365TenantSettingsLoad === 'function') {
        try {
            return window.ms365TenantSettingsLoad() || {};
        } catch {
            /* ignore */
        }
    }
    return {};
}

async function ensureToken(scopes) {
    return await G().getGraphToken(scopes || SCOPES_GRAPH);
}

async function findListByTitle(token, siteId, listTitle) {
    const title = String(listTitle || '').trim();
    const path =
        G().graphPathSite(siteId) +
        '/lists?$filter=' +
        encodeURIComponent("displayName eq '" + title.replace(/'/g, "''") + "'") +
        '&$select=id,displayName,webUrl';
    const data = await G().graphJson('GET', path, token, undefined, 'v1.0');
    const list = (data && data.value) || [];
    return list[0] || null;
}

async function addMissingColumns(siteId, listId, token, defs, write) {
    const colsPath = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns?$top=200';
    const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
    const existing = new Set(
        ((colsData && colsData.value) || [])
            .map((c) => String((c && c.name) || '').trim())
            .filter(Boolean)
    );
    const base = G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(listId) + '/columns';
    let added = 0;
    for (let i = 0; i < defs.length; i++) {
        const def = defs[i];
        if (existing.has(def.name)) continue;
        await G().graphJson('POST', base, token, toGraphColumnBody(def), 'v1.0');
        added++;
        write('  + Spalte ' + def.name);
        await G().sleep(120);
    }
    return added;
}

async function resolveGroupId(token, mailOrId) {
    const raw = String(mailOrId || '').trim();
    if (!raw) return '';
    if (/^[0-9a-f-]{36}$/i.test(raw)) return raw;
    const esc = raw.replace(/'/g, "''");
    const filter = encodeURIComponent(
        "mail eq '" + esc + "' or mailNickname eq '" + esc + "' or displayName eq '" + esc + "'"
    );
    const data = await G().graphJson(
        'GET',
        '/groups?$filter=' + filter + '&$select=id,displayName,mail,mailNickname&$top=5',
        token,
        undefined,
        'v1.0'
    );
    const g = ((data && data.value) || [])[0];
    return g && g.id ? String(g.id) : '';
}

/**
 * @param {string} webUrl
 * @param {{
 *   profileId?: string,
 *   terminJahr: string|number,
 *   klassen: string[],
 *   lehrer: string[],
 *   lfs: string[],
 *   wahlfaecher: string[],
 *   seminare: string[],
 *   groupAdmin?: string,
 *   groupLehrer?: string,
 *   groupKandidaten?: string,
 *   applyPermissions?: boolean
 * }} opts
 * @param {(msg: string) => void} [logFn]
 */
async function createSrdpList(webUrl, opts, logFn) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    if (!url) throw new Error('Bitte die Adresse der SharePoint-Website eintragen.');
    const o = opts || {};
    const profile = getProfile(o.profileId);
    const terminJahr = String(o.terminJahr || '').trim();
    const title = listTitleForYear(terminJahr, profile.id);

    const token = await ensureToken();
    write('Löse Website auf …');
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt in der Graph-Antwort.');
    write('Site: ' + (site.displayName || siteId));
    write('Schulform: ' + profile.label + ' · Liste: „' + title + '"');

    const columnDefs = buildAnmeldungColumns({
        profileId: profile.id,
        klassen: o.klassen,
        lehrer: o.lehrer,
        lfs: o.lfs,
        wahlfaecher: o.wahlfaecher,
        seminare: o.seminare,
        terminJahr
    });

    let list = await findListByTitleCandidates(token, siteId, terminJahr, profile.id);
    if (!list || !list.id) {
        write('Erstelle Liste …');
        const created = await G().graphJson(
            'POST',
            G().graphPathSite(siteId) + '/lists',
            token,
            {
                displayName: title,
                description:
                    'Anmeldung ' +
                    profile.examShort +
                    ' ' +
                    profile.label +
                    ' Haupttermin ' +
                    terminJahr +
                    '. Formular mit Varianten-Prüfplan; Anzeige-Name aus Nachname/Vorname.',
                list: { template: 'genericList' }
            },
            'v1.0'
        );
        const listId = created && created.id ? String(created.id) : '';
        if (!listId) throw new Error('Listen-ID fehlt.');
        list = { id: listId, displayName: title, webUrl: created.webUrl || '' };
        write('Liste angelegt: ' + (list.webUrl || listId));
    } else {
        write('Liste bereits vorhanden („' + (list.displayName || title) + '") – ergänze fehlende Spalten.');
    }

    write('Prüfe Spalten …');
    const added = await addMissingColumns(siteId, list.id, token, columnDefs, write);
    write(added ? added + ' Spalte(n) ergänzt.' : 'Spalten vollständig.');

    let host = '';
    try {
        host = new URL(url).hostname;
    } catch {
        host = '';
    }
    if (!host) throw new Error('SharePoint-Host aus URL nicht lesbar.');

    write('SharePoint-Token (Ansichten / Rechte) …');
    const spoScope = 'https://' + host + '/Sites.FullControl.All';
    let spoToken;
    try {
        spoToken = await G().getGraphToken([spoScope]);
    } catch (e) {
        throw new Error(
            'SharePoint-Token fehlgeschlagen (Sites.FullControl.All?). ' +
                (e && e.message ? e.message : e)
        );
    }
    const digest = await G().getSpoRequestDigest(url, spoToken);
    const listDisplayName = list.displayName || title;

    try {
        await G().spoPatchTitleField(url, spoToken, digest, listDisplayName, {
            displayName: 'Name (aus Nachname, Vorname)',
            description:
                'Wird in Ansichten aus Nachname und Vorname angezeigt; im Formular ausgeblendet. Optional „Titel nachziehen“ im Wizard.',
            required: false
        });
        write('Title-Feld: optional, im Formular ausgeblendet.');
    } catch (e) {
        write('Hinweis Title-Feld: ' + (e && e.message ? e.message : e));
    }

    write('Formular & Spaltenformatierung …');
    try {
        await applyFormAndColumnFormats(url, spoToken, digest, listDisplayName, terminJahr, profile.id, write);
    } catch (e) {
        write('Hinweis Formatierung: ' + (e && e.message ? e.message : e));
    }

    write('Ansichten …');
    const views = defaultViewsForProfile(profile.id);
    for (let i = 0; i < views.length; i++) {
        const v = views[i];
        try {
            const r = await G().spoEnsureListView(url, spoToken, digest, listDisplayName, v);
            write((r.created ? '  + Ansicht ' : '  · vorhanden ') + v.title);
        } catch (e) {
            write('  ! Ansicht „' + v.title + '": ' + (e && e.message ? e.message : e));
        }
        await G().sleep(150);
    }

    if (o.applyPermissions !== false) {
        write('Berechtigungen …');
        try {
            await applyListPermissions(url, spoToken, digest, listDisplayName, token, o, write);
        } catch (e) {
            write('Hinweis Rechte: ' + (e && e.message ? e.message : e));
        }
    }

    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
        window.ms365ActionLog.append({
            tool: 'sharepoint',
            action: 'create-srdp-list',
            target: url,
            summary: title + ' (' + profile.label + ')'
        });
    }

    write('Fertig. Prüfplan-Beispiel Variante 1: ' + buildPruefplanKurz(1, profile.id));
    return { siteId, list, title, webUrl: list.webUrl || '', profileId: profile.id };
}

async function findListByTitleCandidates(token, siteId, terminJahr, profileId) {
    const titles = listTitleCandidates(terminJahr, profileId);
    for (let i = 0; i < titles.length; i++) {
        const hit = await findListByTitle(token, siteId, titles[i]);
        if (hit && hit.id) return hit;
    }
    return null;
}

/**
 * Conditional Formulas, Column Formatting, ClientFormCustomFormatter.
 */
async function applyFormAndColumnFormats(siteWebUrl, spoToken, digest, listTitle, terminJahr, profileId, write) {
    const specs = buildFieldFormatSpecs(terminJahr, profileId);
    for (let i = 0; i < specs.length; i++) {
        const s = specs[i];
        /** @type {Record<string, unknown>} */
        const body = {};
        if (s.conditionalShowFormula != null) body.ConditionalShowFormula = s.conditionalShowFormula;
        if (s.customFormatter) body.CustomFormatter = JSON.stringify(s.customFormatter);
        if (s.required === false) body.Required = false;
        if (s.required === true) body.Required = true;
        if (s.displayName) body.Title = s.displayName;
        if (s.defaultValue != null && s.defaultValue !== '') body.DefaultValue = String(s.defaultValue);
        if (!Object.keys(body).length) continue;
        try {
            await G().spoPatchListField(siteWebUrl, spoToken, digest, listTitle, s.internalName, body);
            write('  · Feld ' + s.internalName);
        } catch (e) {
            write('  ! Feld ' + s.internalName + ': ' + (e && e.message ? e.message : e));
        }
        await G().sleep(100);
    }
    try {
        const fmt = buildClientFormCustomFormatterString(profileId);
        await G().spoSetListClientFormCustomFormatter(siteWebUrl, spoToken, digest, listTitle, fmt);
        write('  · Formular-Header/Sektionen gesetzt');
    } catch (e) {
        write('  ! Formular-Layout: ' + (e && e.message ? e.message : e));
    }
}

/**
 * Title aus Nachname/Vorname nachziehen (bestehende Elemente).
 * @param {string} webUrl
 * @param {string|number} terminJahr
 * @param {(msg: string) => void} [logFn]
 * @param {string} [profileId]
 */
async function backfillItemTitles(webUrl, terminJahr, logFn, profileId) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    const profile = getProfile(profileId);
    const title = listTitleForYear(terminJahr, profile.id);
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');
    const list = await findListByTitleCandidates(token, siteId, terminJahr, profile.id);
    if (!list || !list.id) throw new Error('Liste „' + title + '" nicht gefunden.');

    let patched = 0;
    let skipped = 0;
    let next =
        G().graphPathSite(siteId) +
        '/lists/' +
        encodeURIComponent(list.id) +
        '/items?$expand=fields&$top=50';
    while (next) {
        const data = await G().graphJson('GET', next, token, undefined, 'v1.0');
        const items = (data && data.value) || [];
        for (let i = 0; i < items.length; i++) {
            const item = items[i];
            const fields = (item && item.fields) || {};
            const want = buildItemTitle(fields.Nachname, fields.Vorname);
            if (!want) {
                skipped++;
                continue;
            }
            const cur = String(fields.Title || '').trim();
            const plan = buildPruefplanKurz(fields.Variante, profile.id);
            if (cur === want && String(fields.PruefplanKurz || '') === plan) {
                skipped++;
                continue;
            }
            await G().graphJson(
                'PATCH',
                G().graphPathSite(siteId) +
                    '/lists/' +
                    encodeURIComponent(list.id) +
                    '/items/' +
                    encodeURIComponent(item.id) +
                    '/fields',
                token,
                { Title: want, PruefplanKurz: plan },
                'v1.0'
            );
            patched++;
            await G().sleep(80);
        }
        next = (data && data['@odata.nextLink']) || '';
    }
    write('Titel nachgezogen: ' + patched + ' geändert, ' + skipped + ' übersprungen.');
    return { patched, skipped };
}

async function applyListPermissions(siteWebUrl, spoToken, digest, listTitle, graphToken, opts, write) {
    await G().spoBreakListInheritance(siteWebUrl, spoToken, digest, listTitle, true);
    write('  Vererbung gebrochen (bestehende Rollen kopiert).');

    const assignments = await G().spoListRoleAssignments(siteWebUrl, spoToken, digest, listTitle);
    let removed = 0;
    for (let i = 0; i < assignments.length; i++) {
        const a = assignments[i];
        const member = a.Member || a.member || {};
        if (!isBroadSiteAudience(member)) continue;
        const pid = member.Id != null ? member.Id : a.PrincipalId;
        try {
            await G().spoRemoveRoleAssignment(siteWebUrl, spoToken, digest, listTitle, pid);
            removed++;
            write('  − ' + (member.Title || pid));
        } catch (e) {
            write('  Hinweis Entfernen: ' + (e && e.message ? e.message : e));
        }
    }
    if (removed) write('  Breite Rollen entfernt: ' + removed);

    async function grant(raw, roleId, label) {
        const id = await resolveGroupId(graphToken, raw);
        if (!id) {
            write('  ! Gruppe nicht gefunden (' + label + '): ' + raw);
            return;
        }
        const principal = await G().spoEnsureUser(siteWebUrl, spoToken, digest, entraGroupLogonName(id));
        await G().spoAddRoleAssignment(siteWebUrl, spoToken, digest, listTitle, principal.id, roleId);
        write('  + ' + label + ': ' + (principal.title || raw));
    }

    if (opts.groupAdmin) await grant(opts.groupAdmin, SPO_ROLE.fullControl, 'Admin');
    if (opts.groupLehrer) await grant(opts.groupLehrer, SPO_ROLE.read, 'Lehrer');
    if (opts.groupKandidaten) await grant(opts.groupKandidaten, SPO_ROLE.contribute, 'Kandidaten');
}

/**
 * @param {string} webUrl
 * @param {string|number} terminJahr
 * @param {(msg: string) => void} [logFn]
 * @param {string} [profileId]
 */
async function probeSrdpHealth(webUrl, terminJahr, logFn, profileId) {
    const write = typeof logFn === 'function' ? logFn : log;
    const url = String(webUrl || '').trim();
    const profile = getProfile(profileId);
    const title = listTitleForYear(terminJahr, profile.id);
    const required = requiredColumnNamesForProfile(profile.id);
    const token = await ensureToken();
    const site = await G().resolveSiteFromWebUrl(token, url);
    const siteId = site && site.id ? String(site.id) : '';
    if (!siteId) throw new Error('Site-ID fehlt.');

    const list = await findListByTitleCandidates(token, siteId, terminJahr, profile.id);
    if (!list || !list.id) {
        write('Fehlt: Liste „' + title + '"');
        return { ok: false, title, missingList: true, missingColumns: required.slice() };
    }

    const colsPath =
        G().graphPathSite(siteId) + '/lists/' + encodeURIComponent(list.id) + '/columns?$top=200';
    const colsData = await G().graphJson('GET', colsPath, token, undefined, 'v1.0');
    const colNames = ((colsData && colsData.value) || [])
        .map((c) => String((c && c.name) || '').trim())
        .filter(Boolean);
    const missing = required.filter((n) => colNames.indexOf(n) === -1);
    if (missing.length) write('Fehlende Spalten: ' + missing.join(', '));
    else write('Spalten OK · ' + (list.webUrl || list.displayName || title));

    const panel = $('srsHealth');
    if (panel) {
        panel.hidden = false;
        panel.textContent = missing.length
            ? 'Status: Spalten fehlen – Details im Protokoll.'
            : 'Status: OK – Liste „' + (list.displayName || title) + '" bereit.';
        panel.classList.toggle('ok', missing.length === 0);
        panel.classList.toggle('warn', missing.length > 0);
    }

    return {
        ok: missing.length === 0,
        title: list.displayName || title,
        webUrl: list.webUrl || '',
        missingColumns: missing,
        profileId: profile.id
    };
}

function currentProfileId() {
    return String(($('srsProfile') && $('srsProfile').value) || 'hak').trim().toLowerCase() || 'hak';
}

function readFormOpts() {
    const terminJahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    const profileId = currentProfileId();
    const klassen = Array.from(document.querySelectorAll('input[name="srsKlasse"]:checked')).map(
        (el) => el.value
    );
    const lfs = parseChoiceLines(($('srsLfs') && $('srsLfs').value) || '');
    const wahlfaecher = parseChoiceLines(($('srsWahlfach') && $('srsWahlfach').value) || '');
    const seminare = parseChoiceLines(($('srsSeminar') && $('srsSeminar').value) || '');
    const stammdaten = loadStammdaten();
    const lehrer = teacherChoiceLabels(stammdaten.teachers || []);
    return {
        profileId,
        terminJahr,
        klassen,
        lehrer,
        lfs,
        wahlfaecher,
        seminare,
        groupAdmin: String(($('srsGroupAdmin') && $('srsGroupAdmin').value) || '').trim(),
        groupLehrer: String(($('srsGroupLehrer') && $('srsGroupLehrer').value) || '').trim(),
        groupKandidaten: String(($('srsGroupKandidaten') && $('srsGroupKandidaten').value) || '').trim(),
        applyPermissions: !($('srsSkipPerms') && $('srsSkipPerms').checked)
    };
}

function persistForm() {
    saveWizardState({
        siteUrl: String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim(),
        profileId: currentProfileId(),
        terminJahr: String(($('srsJahr') && $('srsJahr').value) || '').trim(),
        lfs: String(($('srsLfs') && $('srsLfs').value) || ''),
        wahlfach: String(($('srsWahlfach') && $('srsWahlfach').value) || ''),
        seminar: String(($('srsSeminar') && $('srsSeminar').value) || ''),
        groupAdmin: String(($('srsGroupAdmin') && $('srsGroupAdmin').value) || ''),
        groupLehrer: String(($('srsGroupLehrer') && $('srsGroupLehrer').value) || ''),
        groupKandidaten: String(($('srsGroupKandidaten') && $('srsGroupKandidaten').value) || ''),
        skipPerms: !!($('srsSkipPerms') && $('srsSkipPerms').checked)
    });
}

function fillKlassenCheckboxes() {
    const host = $('srsKlassenHost');
    if (!host) return;
    const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
    const stammdaten = loadStammdaten();
    const codes = filterAbschlussklassen(stammdaten.classes || [], jahr);
    host.replaceChildren();
    if (!/^\d{4}$/.test(jahr)) {
        host.innerHTML = '<p class="muted" style="margin:0;">Zuerst ein gültiges Terminjahr (YYYY) wählen.</p>';
        return;
    }
    if (!codes.length) {
        host.innerHTML =
            '<p class="muted" style="margin:0;">Keine Klassen mit Abschlussjahr ' +
            jahr +
            ' in den Stammdaten. Bitte Klassen pflegen oder Jahr prüfen.</p>';
        return;
    }
    codes.forEach((code) => {
        const id = 'srsKlasse_' + code.replace(/[^a-zA-Z0-9]/g, '_');
        const label = document.createElement('label');
        label.className = 'checkbox-label';
        label.style.display = 'inline-flex';
        label.style.margin = '0 12px 8px 0';
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.name = 'srsKlasse';
        input.value = code;
        input.id = id;
        input.checked = true;
        label.appendChild(input);
        label.appendChild(document.createTextNode(' ' + code));
        host.appendChild(label);
    });
}

function applyProfileDefaults(forceChoices) {
    const profile = getProfile(currentProfileId());
    const stammdaten = loadStammdaten();
    const note = $('srsProfileNote');
    if (note) note.textContent = profile.notes || '';
    const lfsLabel = $('srsLfsLabel');
    if (lfsLabel) lfsLabel.textContent = (profile.lfsFieldLabel || 'LFS') + '-Optionen';
    if (forceChoices || ($('srsLfs') && !$('srsLfs').value.trim())) {
        if ($('srsLfs')) {
            $('srsLfs').value = suggestLfsChoices(stammdaten.subjects || [], profile.lfsHints).join('\n');
        }
    }
    if (forceChoices || ($('srsWahlfach') && !$('srsWahlfach').value.trim())) {
        if ($('srsWahlfach')) {
            $('srsWahlfach').value = (profile.defaultWahlfaecher || []).join('\n');
        }
    }
    updateListPreview();
}

function prefillFromStammdatenAndState() {
    const state = loadWizardState();
    const stammdaten = loadStammdaten();

    try {
        const setup =
            window.ms365AppDataV2 && window.ms365AppDataV2.getSetup
                ? window.ms365AppDataV2.getSetup()
                : null;
        const saved = setup && setup.intranetSiteUrl ? String(setup.intranetSiteUrl).trim() : '';
        if ($('srsSiteUrl') && !$('srsSiteUrl').value) {
            $('srsSiteUrl').value = state.siteUrl || saved || '';
        }
    } catch {
        if ($('srsSiteUrl') && state.siteUrl) $('srsSiteUrl').value = state.siteUrl;
    }

    if ($('srsProfile')) {
        const pid = state.profileId && PROFILE_IDS.indexOf(state.profileId) !== -1 ? state.profileId : 'hak';
        $('srsProfile').value = pid;
    }

    const yearNow = String(new Date().getFullYear());
    if ($('srsJahr')) $('srsJahr').value = state.terminJahr || yearNow;

    if ($('srsLfs') && state.lfs) $('srsLfs').value = state.lfs;
    if ($('srsWahlfach') && state.wahlfach) $('srsWahlfach').value = state.wahlfach;
    if ($('srsSeminar')) $('srsSeminar').value = state.seminar || '';
    if ($('srsGroupAdmin')) $('srsGroupAdmin').value = state.groupAdmin || '';
    if ($('srsGroupLehrer')) $('srsGroupLehrer').value = state.groupLehrer || '';
    if ($('srsGroupKandidaten')) $('srsGroupKandidaten').value = state.groupKandidaten || '';
    if ($('srsSkipPerms')) $('srsSkipPerms').checked = !!state.skipPerms;

    const teacherCount = teacherChoiceLabels(stammdaten.teachers || []).length;
    const el = $('srsTeacherHint');
    if (el) {
        el.textContent =
            teacherCount +
            ' Lehrkraft/Lehrkräfte aus Stammdaten → Choice-Felder Betreuung / Fach / Wahlfach.';
    }

    applyProfileDefaults(!state.lfs && !state.wahlfach);
    fillKlassenCheckboxes();
    updateListPreview();
}

function updateListPreview() {
    const el = $('srsListPreview');
    if (!el) return;
    try {
        const y = String(($('srsJahr') && $('srsJahr').value) || '').trim();
        const pid = currentProfileId();
        el.textContent = /^\d{4}$/.test(y)
            ? listTitleForYear(y, pid)
            : listTitleForYear(new Date().getFullYear(), pid).replace(/\d{4}$/, 'YYYY');
    } catch {
        el.textContent = 'sRDP-Anmeldungen … YYYY';
    }
}

async function runCreate() {
    const logEl = $('srsLog');
    if (logEl) logEl.textContent = '';
    persistForm();
    const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
    const opts = readFormOpts();
    if (!/^\d{4}$/.test(opts.terminJahr)) throw new Error('Terminjahr (YYYY) fehlt.');
    if (!opts.klassen.length) throw new Error('Mindestens eine Abschlussklasse auswählen.');
    if (!opts.lehrer.length) throw new Error('Keine Lehrer in den Stammdaten – bitte zuerst pflegen.');
    if (!opts.lfs.length) throw new Error('Mindestens eine LFS-/Sprach-Option eintragen.');
    if (!opts.wahlfaecher.length) throw new Error('Mindestens ein Wahlfach eintragen.');
    if (opts.applyPermissions) {
        if (!opts.groupAdmin || !opts.groupLehrer || !opts.groupKandidaten) {
            throw new Error(
                'Für Rechte: Admin-, Lehrer- und Kandidaten-Gruppe angeben – oder „Rechte überspringen“.'
            );
        }
    }
    const result = await createSrdpList(webUrl, opts);
    toast('Liste angelegt bzw. ergänzt (' + getProfile(opts.profileId).label + ').');
    const link = $('srsListLink');
    if (link && result.webUrl) {
        link.hidden = false;
        link.href = result.webUrl;
        link.textContent = 'Liste öffnen: ' + result.title;
    }
    return result;
}

window.ms365SpoSrdp = {
    createList: createSrdpList,
    probeHealth: probeSrdpHealth,
    backfillItemTitles: backfillItemTitles,
    listTitleForYear,
    getProfile,
    PROFILE_IDS,
    STORAGE_KEY
};

function wireUi() {
    if ($('srsProfile')) {
        PROFILE_IDS.forEach((id) => {
            const p = getProfile(id);
            const opt = document.createElement('option');
            opt.value = p.id;
            opt.textContent = p.label + ' (' + p.examShort + ')';
            $('srsProfile').appendChild(opt);
        });
    }

    prefillFromStammdatenAndState();

    if ($('srsProfile')) {
        $('srsProfile').addEventListener('change', () => {
            applyProfileDefaults(true);
            persistForm();
        });
    }

    if ($('srsJahr')) {
        $('srsJahr').addEventListener('input', () => {
            fillKlassenCheckboxes();
            updateListPreview();
        });
        $('srsJahr').addEventListener('change', () => {
            fillKlassenCheckboxes();
            updateListPreview();
        });
    }

    const runBtn = $('srsBtnRun');
    if (runBtn) {
        runBtn.addEventListener('click', () => {
            const y = String(($('srsJahr') && $('srsJahr').value) || '').trim();
            const pid = currentProfileId();
            let title = 'Anmeldung';
            try {
                title = listTitleForYear(y, pid);
            } catch {
                /* ignore */
            }
            if (
                !window.confirm(
                    'Liste „' +
                        title +
                        '" anlegen bzw. fehlende Spalten/Ansichten ergänzen und Rechte setzen?'
                )
            ) {
                return;
            }
            runCreate().catch((e) => {
                log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                toast('Fehler: ' + (e && e.message ? e.message : e));
            });
        });
    }

    const probeBtn = $('srsBtnProbe');
    if (probeBtn) {
        probeBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
            if (!webUrl) {
                toast('Website-URL fehlt.');
                return;
            }
            ensureToken()
                .then((token) => G().resolveSiteFromWebUrl(token, webUrl))
                .then((site) => {
                    log('Site gefunden: ' + (site.displayName || '') + '\nid: ' + (site.id || ''));
                    if (site.webUrl) log('webUrl: ' + site.webUrl);
                    toast('Website erkannt.');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const healthBtn = $('srsBtnHealth');
    if (healthBtn) {
        healthBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
            const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
            probeSrdpHealth(webUrl, jahr, log, currentProfileId())
                .then((summary) => {
                    toast(summary && summary.ok ? 'Listen-Check OK' : 'Listen-Check: bitte Protokoll prüfen');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const titleBtn = $('srsBtnTitles');
    if (titleBtn) {
        titleBtn.addEventListener('click', () => {
            if ($('srsLog')) $('srsLog').textContent = '';
            const webUrl = String(($('srsSiteUrl') && $('srsSiteUrl').value) || '').trim();
            const jahr = String(($('srsJahr') && $('srsJahr').value) || '').trim();
            backfillItemTitles(webUrl, jahr, log, currentProfileId())
                .then((r) => {
                    toast('Titel: ' + (r.patched || 0) + ' aktualisiert');
                })
                .catch((e) => {
                    log('FEHLER: ' + (e && e.message ? e.message : String(e)));
                    toast('Fehler: ' + (e && e.message ? e.message : e));
                });
        });
    }

    const saveBtn = $('srsBtnSave');
    if (saveBtn) {
        saveBtn.addEventListener('click', () => {
            persistForm();
            toast('Einstellungen gespeichert.');
        });
    }

    const refreshBtn = $('srsBtnRefreshKlassen');
    if (refreshBtn) {
        refreshBtn.addEventListener('click', () => fillKlassenCheckboxes());
    }
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', wireUi);
} else {
    wireUi();
}
