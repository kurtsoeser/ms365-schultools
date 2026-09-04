
/**
 * Reine Logik für Teamnamen-Muster (Drag-and-drop-Builder / Generierung).
 * Kein DOM – nur Datenstrukturen und Zusammensetzung.
 */
function defaultTeamNamePattern() {
    return [
        { type: 'yearPrefix' },
        { type: 'text', value: ' | ' },
        { type: 'klasse' },
        { type: 'text', value: ' | ' },
        { type: 'fach' }
    ];
}

const FIELD_TOKEN_TYPES = new Set(['yearPrefix', 'klasse', 'fach', 'gruppe', 'lehrer']);

function normalizePattern(pattern) {
    const arr = Array.isArray(pattern) ? pattern : [];
    const out = [];
    arr.forEach((p) => {
        if (!p || typeof p !== 'object') return;
        const type = String(p.type || '').trim();
        if (!type) return;
        if (type === 'text') {
            out.push({ type: 'text', value: String(p.value ?? '') });
        } else if (FIELD_TOKEN_TYPES.has(type)) {
            out.push({ type });
        }
    });
    return out.length ? out : defaultTeamNamePattern();
}

function buildTeamNameFromPattern(pattern, ctx) {
    const parts = [];
    normalizePattern(pattern).forEach((p) => {
        if (p.type === 'text') parts.push(String(p.value ?? ''));
        else if (p.type === 'yearPrefix') parts.push(String(ctx.yearPrefix ?? ''));
        else if (p.type === 'klasse') parts.push(String(ctx.klasse ?? ''));
        else if (p.type === 'fach') parts.push(String(ctx.fach ?? ''));
        else if (p.type === 'gruppe') parts.push(String(ctx.gruppe ?? ''));
        else if (p.type === 'lehrer') parts.push(String(ctx.lehrer ?? ''));
    });
    return parts.join('');
}

function sanitizeGruppenmailPart(value) {
    return String(value ?? '').trim().replace(/\s+/g, '-');
}

/**
 * Gleiche Bausteine wie Team-Name – Trenner (text-Tokens) werden zu „-“ zwischen Segmenten.
 * @param {object} [helpers] {{ formatKlasse?: fn, sanitizeGruppe?: fn }}
 */
function buildGruppenmailFromPattern(pattern, ctx, helpers) {
    helpers = helpers || {};
    const formatKlasse = helpers.formatKlasse || sanitizeGruppenmailPart;
    const sanitizeGruppe =
        helpers.sanitizeGruppe ||
        function (g) {
            if (!g || !String(g).trim()) return '';
            return sanitizeGruppenmailPart(g);
        };
    const segments = [];
    normalizePattern(pattern).forEach((p) => {
        if (p.type === 'text') return;
        let v = '';
        if (p.type === 'yearPrefix') v = sanitizeGruppenmailPart(ctx.yearPrefix);
        else if (p.type === 'klasse') v = formatKlasse(ctx.klasse);
        else if (p.type === 'fach') v = sanitizeGruppenmailPart(ctx.fach);
        else if (p.type === 'gruppe') v = sanitizeGruppe(ctx.gruppe);
        else if (p.type === 'lehrer') v = sanitizeGruppenmailPart(ctx.lehrer);
        if (v) segments.push(v);
    });
    return segments.join('-');
}

function tokenLabel(t) {
    if (t.type === 'yearPrefix') return 'Schuljahr';
    if (t.type === 'klasse') return 'Klasse';
    if (t.type === 'fach') return 'Fach';
    if (t.type === 'gruppe') return 'Gruppe';
    if (t.type === 'lehrer') return 'Lehrer';
    if (t.type === 'text') return `Text`;
    return t.type;
}

window.ms365KursteamTeamNames = {
    defaultTeamNamePattern,
    normalizePattern,
    buildTeamNameFromPattern,
    buildGruppenmailFromPattern,
    tokenLabel
};
