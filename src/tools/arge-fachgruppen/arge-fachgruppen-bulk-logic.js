/**
 * Reine Hilfen für Sammelaktionen ARGE & Fachgruppen (ohne DOM/Graph).
 */

/** Wie Graph Unified-Group mailNickname: @ ( ) \ [ ] " ; : < > , Leerzeichen verboten; _ . - ok. */
const GRAPH_MAILNICKNAME_INVALID = /[@()[\]\\";:<>,\s]/;

/**
 * @param {string} raw
 * @returns {string}
 */
export function sanitizeAfgMailNickname(raw) {
    const s = String(raw ?? '')
        .trim()
        .toLowerCase();
    let out = '';
    for (let i = 0; i < s.length; i++) {
        const c = s.charCodeAt(i);
        if (c < 32 || c === 127 || c > 127) continue;
        const ch = s.charAt(i);
        if (GRAPH_MAILNICKNAME_INVALID.test(ch)) continue;
        out += ch;
    }
    if (!out) out = 'group';
    return out.slice(0, 60);
}

/**
 * Alias = optionales Präfix + Kürzel. Leeres Präfix → nur Kürzel (Unterstriche bleiben).
 * @param {string} prefix
 * @param {string} code
 * @returns {string}
 */
export function buildAfgMailNickname(prefix, code) {
    const pre = String(prefix ?? '');
    const tail = sanitizeAfgMailNickname(String(code || 'x'));
    return sanitizeAfgMailNickname(pre + tail);
}

/**
 * @param {{ code?: string, nick?: string }[]} rows
 * @param {number} [max=12]
 * @returns {string}
 */
export function formatAfgBulkCreatePreviewLines(rows, max) {
    const list = Array.isArray(rows) ? rows : [];
    const lim = typeof max === 'number' && max > 0 ? max : 12;
    const lines = list.slice(0, lim).map(function (r) {
        const code = String((r && r.code) || '').trim() || '–';
        const nick = String((r && r.nick) || '').trim() || '–';
        return code + ' → ' + nick;
    });
    if (list.length > lim) lines.push('…');
    return lines.join('\n');
}

/**
 * @param {{ headEmail?: string }|null|undefined} row
 * @param {string[]} [direktionEmails]
 * @returns {string[]}
 */
export function collectAfgCreateOwners(row, direktionEmails) {
    const out = [];
    const seen = new Set();
    function add(raw) {
        const em = String(raw || '')
            .trim()
            .toLowerCase();
        if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
        seen.add(em);
        out.push(em);
    }
    if (row && row.headEmail) add(row.headEmail);
    (Array.isArray(direktionEmails) ? direktionEmails : []).forEach(add);
    return out;
}

/**
 * @param {{
 *   kindLabel: string,
 *   count: number,
 *   previewLines: string,
 *   createTeam?: boolean
 * }} opts
 * @returns {string}
 */
export function buildAfgBulkCreateConfirmMessage(opts) {
    const o = opts || {};
    const n = Number(o.count) || 0;
    const label = String(o.kindLabel || 'Einträge').trim() || 'Einträge';
    const preview = String(o.previewLines || '').trim();
    let msg =
        String(n) +
        ' ungematchte ' +
        label +
        ' als Microsoft‑365‑Gruppe anlegen und matchen?\n\n' +
        'Kürzel → geplanter Alias:\n' +
        (preview || '–');
    if (o.createTeam) {
        msg += '\n\nPro Gruppe wird auch ein Microsoft Team bereitgestellt.';
    }
    return msg;
}

/**
 * Anzeigename ohne „Arbeitsgruppe “ / „Fach “ – Katalogname bzw. Kürzel.
 * @param {{ code?: string, name?: string }|null|undefined} row
 * @returns {string}
 */
export function displayNameForAfgCreate(row) {
    const code = String((row && row.code) || '').trim();
    const name = String((row && row.name) || '').trim();
    return name || code || '';
}

/**
 * Entfernt führendes „Arbeitsgruppe “ oder „Fach “ (einmalig).
 * @param {string} displayName
 * @returns {string}
 */
export function stripAfgDisplayNamePrefix(displayName) {
    let s = String(displayName ?? '').trim();
    const prefixes = ['Arbeitsgruppe ', 'Fach '];
    for (let i = 0; i < prefixes.length; i++) {
        const p = prefixes[i];
        if (s.length >= p.length && s.slice(0, p.length).toLowerCase() === p.toLowerCase()) {
            s = s.slice(p.length).trim();
            break;
        }
    }
    return s;
}

/**
 * @param {string} displayName
 * @returns {boolean}
 */
export function afgNeedsDisplayNameCleanup(displayName) {
    const raw = String(displayName ?? '').trim();
    if (!raw) return false;
    return stripAfgDisplayNamePrefix(raw) !== raw;
}

/**
 * Exakter Treffer über mailNickname oder SMTP-Lokalteil (ohne Domain).
 * @param {Array<{ id?: string, mailNickname?: string, mail?: string, displayName?: string }>} candidates
 * @param {string} nick
 * @returns {{ id?: string, mailNickname?: string, mail?: string, displayName?: string }|null}
 */
export function pickAfgGroupByNickname(candidates, nick) {
    const want = String(nick || '')
        .trim()
        .toLowerCase();
    if (!want) return null;
    const list = Array.isArray(candidates) ? candidates : [];
    for (let i = 0; i < list.length; i++) {
        const g = list[i];
        if (!g || !g.id) continue;
        const mn = String(g.mailNickname || '')
            .trim()
            .toLowerCase();
        if (mn && mn === want) return g;
        const mail = String(g.mail || '')
            .trim()
            .toLowerCase();
        const at = mail.indexOf('@');
        const local = at > 0 ? mail.slice(0, at) : mail;
        if (local && local === want) return g;
    }
    return null;
}

/**
 * @param {{ kindLabel: string, count: number, previewLines: string }} opts
 * @returns {string}
 */
export function buildAfgBulkMatchExistingConfirmMessage(opts) {
    const o = opts || {};
    const n = Number(o.count) || 0;
    const label = String(o.kindLabel || 'Einträge').trim() || 'Einträge';
    const preview = String(o.previewLines || '').trim();
    return (
        String(n) +
        ' ungematchte ' +
        label +
        ' mit bestehenden Microsoft‑365‑Gruppen per Alias abgleichen?\n\n' +
        'Kürzel → gesuchter Alias:\n' +
        (preview || '–') +
        '\n\nEs werden keine neuen Gruppen angelegt – nur Treffer mit exakt gleichem Alias/SMTP werden gematcht.'
    );
}

/**
 * @param {{ count: number, previewLines: string }} opts
 * @returns {string}
 */
export function buildAfgBulkRenameConfirmMessage(opts) {
    const o = opts || {};
    const n = Number(o.count) || 0;
    const preview = String(o.previewLines || '').trim();
    return (
        'Bei ' +
        String(n) +
        ' Gruppe(n) das Präfix „Arbeitsgruppe “ / „Fach “ aus dem Anzeigenamen entfernen?\n\n' +
        (preview || '–') +
        '\n\nNur der Anzeigename (und ggf. die Beschreibung) wird geändert – Alias/E‑Mail bleiben gleich.'
    );
}
