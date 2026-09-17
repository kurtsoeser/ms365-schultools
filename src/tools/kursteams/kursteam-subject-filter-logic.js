function normalizeSubjectToken(s) {
    return String(s || '').trim().toUpperCase();
}

/**
 * Trennt Endziffern ab: OMAI1 → { base: 'OMAI', suffix: '1' }.
 * Keine Ziffern / nur Ziffern → suffix leer, base = Token.
 */
function splitSubjectBaseAndSuffix(token) {
    const t = normalizeSubjectToken(token);
    if (!t) return { base: '', suffix: '' };
    const m = t.match(/^(.+?)(\d+)$/);
    if (!m) return { base: t, suffix: '' };
    const base = m[1];
    const suffix = m[2];
    if (!base || !/[A-ZÄÖÜ]/.test(base)) return { base: t, suffix: '' };
    return { base, suffix };
}

function subjectBaseToken(token) {
    return splitSubjectBaseAndSuffix(token).base;
}

function parseExcludeSubjectsFromString(value) {
    return String(value || '')
        .split(',')
        .map(normalizeSubjectToken)
        .filter((x) => x.length > 0);
}

function uniqSortedSubjectTokens(tokens) {
    const uniq = Array.from(new Set((tokens || []).map(normalizeSubjectToken).filter(Boolean)));
    uniq.sort((a, b) => a.localeCompare(b, 'de'));
    return uniq;
}

function collectSubjectsFromRows(rows) {
    const set = new Set();
    (rows || []).forEach((r) => {
        const t = normalizeSubjectToken(r && r.fach);
        if (t) set.add(t);
    });
    return Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
}

/**
 * Gruppiert Fächer nach Basis (OMAI + OMAI1/2/3 → eine Familie).
 * @returns {{ base: string, variants: string[], isFamily: boolean }[]}
 */
function groupSubjectsByBase(subjects) {
    const map = new Map();
    (subjects || []).forEach((raw) => {
        const token = normalizeSubjectToken(raw);
        if (!token) return;
        const base = subjectBaseToken(token);
        if (!map.has(base)) map.set(base, new Set());
        map.get(base).add(token);
    });
    const groups = Array.from(map.entries()).map(([base, set]) => {
        const variants = Array.from(set).sort((a, b) => a.localeCompare(b, 'de'));
        return {
            base,
            variants,
            isFamily: variants.length > 1 || (variants.length === 1 && variants[0] !== base)
        };
    });
    groups.sort((a, b) => a.base.localeCompare(b.base, 'de'));
    return groups;
}

/**
 * Fach-Endziffern → Basis + Gruppe (nur wenn Gruppe leer).
 * @returns {{ fach: string, gruppe: string, changed: boolean, suffix: string }}
 */
function normalizeNumberedSubjectFields(fach, gruppe) {
    const { base, suffix } = splitSubjectBaseAndSuffix(fach);
    const g = String(gruppe || '').trim();
    if (!suffix || !base) {
        return { fach: normalizeSubjectToken(fach) || String(fach || '').trim(), gruppe: g, changed: false, suffix: '' };
    }
    const nextGruppe = g || suffix;
    const nextFach = base;
    const changed = nextFach !== normalizeSubjectToken(fach) || nextGruppe !== g;
    return { fach: nextFach, gruppe: nextGruppe, changed, suffix };
}

function subjectFilterSummaryText(availableCount, excludedCount, familyCount) {
    const a = Number(availableCount) || 0;
    const e = Number(excludedCount) || 0;
    const f = Number(familyCount) || 0;
    if (!a) {
        return 'Noch keine Daten: Importieren Sie zuerst Zeilen in Schritt 1 oder fügen Sie manuell Unterrichtszeilen hinzu.';
    }
    const fam = f > 0 ? ` ${f} nummerierte Familie(n).` : '';
    return `${a} Fach/Fächer erkannt.${fam} ${e} ausgeschlossen.`;
}

window.ms365KursteamSubjectFilterLogic = {
    normalizeSubjectToken,
    splitSubjectBaseAndSuffix,
    subjectBaseToken,
    parseExcludeSubjectsFromString,
    uniqSortedSubjectTokens,
    collectSubjectsFromRows,
    groupSubjectsByBase,
    normalizeNumberedSubjectFields,
    subjectFilterSummaryText
};
