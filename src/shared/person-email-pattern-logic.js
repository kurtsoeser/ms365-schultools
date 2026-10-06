/**
 * E-Mail-Local-Part-Muster: Presets + Baustein-Builder (Kürzel, Vorname, Nachname, …).
 */

export const EMAIL_PATTERN_PRESETS = [
    { id: 'vorname.nachname', label: 'vorname.nachname' },
    { id: 'nachname.vorname', label: 'nachname.vorname' },
    { id: 'v.nachname', label: 'v.nachname' },
    { id: 'vorname.n', label: 'vorname.n' },
    { id: 'vorname_nachname', label: 'vorname_nachname' },
    { id: 'nachname_vorname', label: 'nachname_vorname' },
    { id: 'kuerzel', label: 'kuerzel (nur Kürzel)' },
    { id: 'kuerzel.nachname', label: 'kuerzel.nachname' },
    { id: 'vorname.kuerzel', label: 'vorname.kuerzel' }
];

const FIELD_TYPES = new Set(['vorname', 'nachname', 'kuerzel', 'v0', 'n0']);

export function defaultEmailBuilderPattern() {
    return [
        { type: 'vorname' },
        { type: 'sep', value: '.' },
        { type: 'nachname' }
    ];
}

export function normalizeEmailBuilderPattern(pattern) {
    const out = [];
    (pattern || []).forEach(function (p) {
        if (!p || typeof p !== 'object') return;
        const type = String(p.type || '').trim();
        if (type === 'sep') {
            out.push({ type: 'sep', value: String(p.value ?? '.') });
        } else if (FIELD_TYPES.has(type)) {
            out.push({ type });
        }
    });
    return out.length ? out : defaultEmailBuilderPattern();
}

export function isBuilderPatternId(id) {
    return String(id || '').startsWith('build:');
}

export function serializeEmailBuilderPattern(tokens) {
    return 'build:' + JSON.stringify(normalizeEmailBuilderPattern(tokens));
}

export function parseEmailPatternId(id) {
    const raw = String(id || 'vorname.nachname');
    if (!isBuilderPatternId(raw)) {
        return { kind: 'preset', id: raw };
    }
    try {
        const tokens = JSON.parse(raw.slice(6));
        return { kind: 'build', tokens: normalizeEmailBuilderPattern(tokens) };
    } catch {
        return { kind: 'preset', id: 'vorname.nachname' };
    }
}

export function emailPatternTokenLabel(token) {
    if (!token) return '';
    if (token.type === 'sep') {
        const v = String(token.value ?? '');
        if (v === '') return 'ohne Trenner';
        if (v === '.') return 'Punkt (.)';
        if (v === '_') return 'Unterstrich (_)';
        if (v === '-') return 'Bindestrich (-)';
        return '„' + v + '“';
    }
    if (token.type === 'vorname') return 'Vorname';
    if (token.type === 'nachname') return 'Nachname';
    if (token.type === 'kuerzel') return 'Kürzel';
    if (token.type === 'v0') return 'V-Anfangsbuchstabe';
    if (token.type === 'n0') return 'N-Anfangsbuchstabe';
    return token.type;
}

function normStr(v) {
    return String(v ?? '').trim();
}

function stripDiacritics(s) {
    return String(s || '')
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/Ä/g, 'Ae')
        .replace(/Ö/g, 'Oe')
        .replace(/Ü/g, 'Ue')
        .replace(/ß/g, 'ss')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '');
}

function sanitizeToken(s) {
    return stripDiacritics(normStr(s))
        .toLowerCase()
        .replace(/[^a-z0-9]+/g, '')
        .replace(/^-+|-+$/g, '');
}

function nameParts(s) {
    return normStr(s)
        .split(/[\s\-]+/)
        .map(sanitizeToken)
        .filter(Boolean);
}

function resolveFirstNameMode(mode) {
    const m = String(mode ?? 'first').trim().toLowerCase();
    if (m === 'all' || m === 'alle') return 'all';
    return 'first';
}

function givenPartsForPattern(vorname, firstNameMode) {
    const parts = nameParts(vorname);
    if (!parts.length) return [];
    if (resolveFirstNameMode(firstNameMode) === 'all') return parts;
    return [parts[0]];
}

function joinParts(parts, sep) {
    return (parts || []).filter(Boolean).join(sep || '');
}

function presetLocalPart(id, vorname, nachname, firstNameMode) {
    const patternId = String(id || 'vorname.nachname').toLowerCase();
    const vParts = givenPartsForPattern(vorname, firstNameMode);
    const nParts = nameParts(nachname);
    const v0 = vParts[0] || '';
    const n0 = nParts[0] || '';
    const vAll = joinParts(vParts, '');
    const nAll = joinParts(nParts, '');
    const vDot = joinParts(vParts, '.') || vAll;
    const nDot = joinParts(nParts, '.') || nAll;
    const k = sanitizeToken(arguments[4]);

    if (patternId === 'kuerzel') return k || vAll || nAll;
    if (patternId === 'kuerzel.nachname') return k && nDot ? k + '.' + nDot : k || nAll;
    if (patternId === 'vorname.kuerzel') return vDot && k ? vDot + '.' + k : vAll || k;
    if (patternId === 'nachname.vorname') return nDot && vDot ? nDot + '.' + vDot : nAll + vAll;
    if (patternId === 'v.nachname') return v0 && nDot ? v0.charAt(0) + '.' + nDot : (v0.charAt(0) || '') + nAll;
    if (patternId === 'vorname.n') return vDot && n0 ? vDot + '.' + n0.charAt(0) : vAll + (n0.charAt(0) || '');
    if (patternId === 'vorname_nachname') return vAll && nAll ? vAll + '_' + nAll : vAll || nAll;
    if (patternId === 'nachname_vorname') return nAll && vAll ? nAll + '_' + vAll : nAll || vAll;
    return vDot && nDot ? vDot + '.' + nDot : vAll || nAll;
}

function tokenValue(type, ctx) {
    const vorname = normStr(ctx.vorname);
    const nachname = normStr(ctx.nachname);
    const kuerzel = sanitizeToken(ctx.kuerzel);
    const mode = ctx.firstNameMode;
    const vParts = givenPartsForPattern(vorname, mode);
    const nParts = nameParts(nachname);
    const v0 = vParts[0] || '';
    const n0 = nParts[0] || '';
    if (type === 'vorname') return joinParts(vParts, '') || sanitizeToken(vorname);
    if (type === 'nachname') return joinParts(nParts, '') || sanitizeToken(nachname);
    if (type === 'kuerzel') return kuerzel;
    if (type === 'v0') return v0 ? v0.charAt(0) : '';
    if (type === 'n0') return n0 ? n0.charAt(0) : '';
    return '';
}

function cleanLocal(s) {
    return stripDiacritics(normStr(s))
        .toLowerCase()
        .replace(/[^a-z0-9._-]+/g, '')
        .replace(/^[._-]+|[._-]+$/g, '');
}

/**
 * @param {string} patternId preset oder build:…
 * @param {{ vorname?: string, nachname?: string, kuerzel?: string, firstNameMode?: string }} ctx
 */
export function localPartFromEmailPatternId(patternId, ctx) {
    const parsed = parseEmailPatternId(patternId);
    const c = ctx || {};
    if (parsed.kind === 'preset') {
        return cleanLocal(
            presetLocalPart(parsed.id, c.vorname, c.nachname, c.firstNameMode, c.kuerzel)
        );
    }
    let raw = '';
    (parsed.tokens || []).forEach(function (t) {
        if (t.type === 'sep') raw += String(t.value ?? '');
        else raw += tokenValue(t.type, c);
    });
    return cleanLocal(raw);
}

if (typeof window !== 'undefined') {
    window.ms365PersonEmailPatternLogic = {
        EMAIL_PATTERN_PRESETS,
        defaultEmailBuilderPattern,
        normalizeEmailBuilderPattern,
        serializeEmailBuilderPattern,
        parseEmailPatternId,
        isBuilderPatternId,
        emailPatternTokenLabel,
        localPartFromEmailPatternId
    };
}
