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
