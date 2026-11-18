/**
 * Reine Logik: Fachschafts-/ARGE-Kanalvorlagen, Team-Filter, Idempotenz-Plan.
 * Kein DOM, kein Graph.
 * @file
 */

/** Standard-Vorlage für Fachschafts-/ARGE-Teams (Screenshot-Stil mit Bindestrich). */
var FACHSCHAFT_PRESET_TITLES = ['Fachkonferenz', 'Curriculum', 'Material', 'Verschiedenes'];

/**
 * @param {unknown} v
 * @returns {string}
 */
function normalizeSearchTerm(v) {
    return String(v == null ? '' : v).trim();
}

function normLower(v) {
    return String(v ?? '').trim().toLowerCase();
}

/**
 * Kanalanzeigename aus Nummer + Titel.
 * @param {number|string} number
 * @param {string} title
 * @param {'hyphen'|'space'} [style]
 * @returns {string} max. 50 Zeichen (Teams-Limit)
 */
function formatChannelDisplayName(number, title, style) {
    var n = String(number == null ? '' : number).trim();
    var padded = /^\d{1,2}$/.test(n) ? n.padStart(2, '0') : n;
    var t = String(title == null ? '' : title).trim().replace(/\s+/g, ' ');
    if (!padded || !t) return '';
    var sep = style === 'space' ? ' ' : '-';
    return (padded + sep + t).slice(0, 50);
}

/**
 * @param {'hyphen'|'space'} [style]
 * @returns {Array<{ displayName: string, title: string, number: string }>}
 */
function buildFachschaftPresetChannels(style) {
    return FACHSCHAFT_PRESET_TITLES.map(function (title, i) {
        const number = String(i + 1).padStart(2, '0');
        return {
            number: number,
            title: title,
            displayName: formatChannelDisplayName(number, title, style || 'hyphen')
        };
    });
}

/**
 * UI-/JSON-Zeilen → bereinigte Kanalliste (ohne Duplikate, ohne Leerzeilen).
 * @param {Array<{ displayName?: string, description?: string }|string>} rows
 * @returns {Array<{ displayName: string, description: string }>}
 */
function normalizeChannelSpecs(rows) {
    const list = Array.isArray(rows) ? rows : [];
    /** @type {Array<{ displayName: string, description: string }>} */
    const out = [];
    const seen = Object.create(null);
    for (const row of list) {
        const displayName =
            typeof row === 'string'
                ? String(row || '').trim().slice(0, 50)
                : String((row && row.displayName) || '').trim().slice(0, 50);
        if (!displayName) continue;
        const key = normLower(displayName);
        if (seen[key]) continue;
        seen[key] = true;
        const description =
            typeof row === 'string'
                ? ''
                : String((row && row.description) || '').trim().slice(0, 1024);
        out.push({ displayName, description });
    }
    return out;
}

/**
 * Eine Gruppe passt, wenn der Suchbegriff im Anzeigenamen oder mailNickname vorkommt.
 * @param {{ displayName?: string, mailNickname?: string }} group
 * @param {string} term
 * @returns {boolean}
 */
function groupMatchesTerm(group, term) {
    const t = normLower(term);
    if (!t) return false;
    const dn = normLower(group && group.displayName);
    const mn = normLower(group && group.mailNickname);
    return (!!dn && dn.indexOf(t) !== -1) || (!!mn && mn.indexOf(t) !== -1);
}

/**
 * @param {Array<{ id?: string, displayName?: string, mailNickname?: string, mail?: string }>} groups
 * @param {string} term
 * @returns {Array<{ id: string, displayName: string, mailNickname: string, mail: string }>}
 */
function filterAndSortGroupsByTerm(groups, term) {
    const t = normalizeSearchTerm(term);
    if (!t) return [];
    return (Array.isArray(groups) ? groups : [])
        .filter((g) => g && g.id && groupMatchesTerm(g, t))
        .map((g) => ({
            id: String(g.id),
            displayName: String(g.displayName || ''),
            mailNickname: String(g.mailNickname || ''),
            mail: String(g.mail || '')
        }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName, 'de'));
}

/**
 * Pro gewünschtem Kanal: create | exists | skip.
 * @param {Array<{ displayName: string, description?: string }>} wanted
 * @param {Array<{ id?: string, displayName?: string }>} existingChannels
 * @returns {Array<{ displayName: string, description: string, action: 'create'|'exists', existingId?: string }>}
 */
function planChannelsForTeam(wanted, existingChannels) {
    const specs = normalizeChannelSpecs(wanted);
    const byName = Object.create(null);
    (Array.isArray(existingChannels) ? existingChannels : []).forEach(function (c) {
        if (!c || !c.displayName) return;
        byName[normLower(c.displayName)] = c;
    });
    return specs.map(function (spec) {
        const hit = byName[normLower(spec.displayName)];
        if (hit && hit.id) {
            return {
                displayName: spec.displayName,
                description: spec.description || '',
                action: 'exists',
                existingId: String(hit.id)
            };
        }
        return {
            displayName: spec.displayName,
            description: spec.description || '',
            action: 'create'
        };
    });
}

/**
 * @param {Array<{ ok?: boolean, action?: string, skipped?: boolean }>} results
 * @returns {{ total: number, created: number, existed: number, fail: number }}
 */
function buildBulkChannelSummary(results) {
    const list = Array.isArray(results) ? results : [];
    let created = 0;
    let existed = 0;
    let fail = 0;
    list.forEach(function (r) {
        if (!r || r.ok === false) {
            fail += 1;
            return;
        }
        if (r.action === 'exists' || r.skipped) existed += 1;
        else created += 1;
    });
    return { total: list.length, created, existed, fail };
}

window.ms365TeamsKanaeleBulkLogic = {
    FACHSCHAFT_PRESET_TITLES: FACHSCHAFT_PRESET_TITLES,
    normalizeSearchTerm,
    formatChannelDisplayName,
    buildFachschaftPresetChannels,
    normalizeChannelSpecs,
    groupMatchesTerm,
    filterAndSortGroupsByTerm,
    planChannelsForTeam,
    buildBulkChannelSummary
};
