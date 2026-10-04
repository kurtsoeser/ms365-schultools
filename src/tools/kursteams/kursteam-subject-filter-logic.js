import {
    normalizeSubjectToken,
    splitSubjectBaseAndSuffix,
    subjectBaseToken,
    groupSubjectsByBase
} from '../../shared/subject-code-family.js';

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
 * Fach-Endziffern / Ü / Plus → Basis + Gruppe (nur wenn Gruppe leer).
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
    const fam = f > 0 ? ` ${f} Fach-Familie(n).` : '';
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
