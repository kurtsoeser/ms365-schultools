
// Gemeinsamer Namespace für den Kursteam-Modus (kein ES-Module/Bundler nötig).
const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

ns.STORAGE_KEY = 'webuntis-teams-creator-state-v1';
ns.INVALID_CHARS_REPLACE = /[\\%&*+\/=?{}|<>();:,\[\]"öäü]/g;
ns.INVALID_CHARS_TEST = /[\\%&*+\/=?{}|<>();:,\[\]"öäü]/;

ns.escapeHtml = function escapeHtml(text) {
    const d = document.createElement('div');
    d.textContent = text;
    return d.innerHTML;
};

ns.attrEscape = function attrEscape(text) {
    return String(text ?? '').replace(/&/g, '&amp;').replace(/"/g, '&quot;');
};

ns.csvEscapeField = function csvEscapeField(value) {
    const s = String(value ?? '');
    if (/[",\r\n]/.test(s)) {
        return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
};

ns.buildCsvRow = function buildCsvRow(cols) {
    return cols.map(ns.csvEscapeField).join(',') + '\r\n';
};

ns.psEscapeSingle = function psEscapeSingle(s) {
    return String(s ?? '').replace(/'/g, "''");
};

ns.downloadBlob = function downloadBlob(filename, text, mime) {
    const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
    URL.revokeObjectURL(a.href);
};

ns.sanitizeGruppeForMail = function sanitizeGruppeForMail(g) {
    if (!g || !String(g).trim()) return '';
    let s = String(g).replace(/[_\s]+/g, '-').replace(/-+/g, '-');
    s = s.replace(/^[^A-Za-z0-9]+|[^A-Za-z0-9]+$/g, '');
    return s;
};

/** jg20301hma → jg2030-1hma (Lesbarkeit in Kursteam-Gruppenmail). */
ns.formatKlasseSegmentForGruppenmail = function formatKlasseSegmentForGruppenmail(klasseForName) {
    const s = String(klasseForName ?? '').trim().replace(/\s+/g, '-');
    const m = s.match(/^jg(\d{4})([0-9A-Za-z].*)$/i);
    if (m) return ('jg' + m[1] + '-' + m[2]).toLowerCase();
    return s;
};

ns.buildGruppenmailBase = function buildGruppenmailBase(yearPrefix, klasseForName, fach, gruppe) {
    const km = ns.formatKlasseSegmentForGruppenmail(klasseForName);
    const fm = String(fach).replace(/\s+/g, '-');
    let base = `${yearPrefix}-${km}-${fm}`;
    const gs = gruppe ? ns.sanitizeGruppeForMail(gruppe) : '';
    if (gs) base += '-' + gs;
    return base.replace(/\s+/g, '-');
};

ns.resolveDuplicateGruppenmails = function resolveDuplicateGruppenmails(teams) {
    const seen = new Map();
    let adjusted = 0;
    teams.forEach(team => {
        const base = team.gruppenmail;
        let candidate = base;
        let n = 2;
        while (seen.has(candidate)) {
            candidate = base + '-' + n;
            n++;
        }
        if (candidate !== base) {
            team.gruppenmail = candidate;
            team.mailNicknameAdjusted = true;
            adjusted++;
        } else {
            team.mailNicknameAdjusted = false;
        }
        seen.set(candidate, true);
    });
    return adjusted;
};

function longestCommonPrefix(strings) {
    if (!strings.length) return '';
    let prefix = String(strings[0] || '');
    for (let i = 1; i < strings.length; i++) {
        const s = String(strings[i] || '');
        while (prefix && !s.startsWith(prefix)) {
            prefix = prefix.slice(0, -1);
        }
        if (!prefix) return '';
    }
    return prefix;
}

/**
 * Intelligente Zusammenführung: gemeinsames Jahrgang-/Buchstaben-Präfix + abweichende Endungen.
 * 1HMA,1HMB → 1HMAB; 1HMA,1HMB,2HMA,2HMB → 12HMAB; 1AK,1BK → 1AKBK.
 * Wenn kein kürzeres Ergebnis möglich ist → null (Caller fällt auf concat zurück).
 */
function combineClassNamesSmart(classes) {
    const normalized = classes.map((c) => c.replace(/\s+/g, ''));
    const parsed = normalized.map((c) => {
        const m = c.match(/^(\d+)([A-Za-z]+)$/);
        return m ? { year: m[1], letters: m[2].toUpperCase() } : null;
    });
    if (!parsed.every(Boolean)) return null;

    const years = [];
    const yearSeen = new Set();
    parsed.forEach((p) => {
        if (!yearSeen.has(p.year)) {
            yearSeen.add(p.year);
            years.push(p.year);
        }
    });

    const letterParts = parsed.map((p) => p.letters);
    const letterLcp = longestCommonPrefix(letterParts);
    const suffixes = [];
    const suffixSeen = new Set();
    letterParts.forEach((part) => {
        const suf = part.slice(letterLcp.length);
        if (suf && !suffixSeen.has(suf)) {
            suffixSeen.add(suf);
            suffixes.push(suf);
        }
    });

    const result = years.join('') + letterLcp + suffixes.join('');
    const concat = normalized.join('');
    if (!result || result.length >= concat.length) return null;
    return result;
}

/**
 * Mehrere Klassen zu einem Anzeige-/Mail-Segment zusammenführen.
 * Früher: 1AK,1BK → 1AKB (Buchstaben unique) → oft fälschlich „hakb“ im Alias.
 * Default: vollständige Kürzel aneinander (1AK1BK).
 * Optional: „smart“ → 1HMA,1HMB → 1HMAB; „letters“ → AKBK.
 * @param {string} classString Klassen getrennt durch Komma/;/~
 * @param {{ mode?: 'concat'|'smart'|'letters' }} [opts]
 */
ns.combineClassNames = function combineClassNames(classString, opts) {
    const rawMode = opts && opts.mode ? String(opts.mode) : 'concat';
    const mode = rawMode === 'letters' || rawMode === 'smart' ? rawMode : 'concat';
    const classes = String(classString || '')
        .split(/[,;~]+/)
        .map((c) => c.trim())
        .filter(Boolean);
    if (classes.length === 0) return String(classString || '');
    if (classes.length === 1) return classes[0];

    if (mode === 'letters') {
        const parts = classes
            .map((c) => {
                const m = c.match(/^\d+([A-Za-z]+)$/);
                return m ? m[1].toUpperCase() : c.replace(/[^A-Za-z]/g, '').toUpperCase();
            })
            .filter(Boolean);
        return parts.join('');
    }

    if (mode === 'smart') {
        const smart = combineClassNamesSmart(classes);
        if (smart) return smart;
    }

    // concat: 1AK + 1BK → 1AK1BK (eindeutig, kein „AKB“/„hakb“)
    return classes.map((c) => c.replace(/\s+/g, '')).join('');
};

/** True wenn die Klassen-Zelle mehrere Klassen enthält (geteilt / Schwerpunkte). */
ns.isCombinedClassCell = function isCombinedClassCell(raw) {
    const s = String(raw || '').trim();
    if (!s) return false;
    if (/[,;~]/.test(s)) return true;
    // bereits zusammengezogen: 1AK1BK, 3AS3BS, …
    return /(?:\d+[A-Za-z]+){2,}/.test(s);
};

