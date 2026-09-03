/**
 * Reine Logik für die Sammel-Archivierung im Team-Archiv (Suchbegriff → passende Kursteams).
 * Kein DOM, kein Graph – nur Filtern/Sortieren der bereits geladenen Gruppenliste.
 * @file
 */

/**
 * @param {unknown} v
 * @returns {string}
 */
function normalizeSearchTerm(v) {
    return String(v ?? '').trim();
}

function normLower(v) {
    return String(v ?? '').trim().toLowerCase();
}

/**
 * Eine Gruppe passt, wenn der Suchbegriff (z. B. Schuljahr-Präfix „SJ25“) im Anzeigenamen
 * oder im mailNickname vorkommt (case-insensitive, Teilstring – Reihenfolge der Namensbausteine
 * ist bei Kursteams frei wählbar, „enthält“ ist daher robuster als „beginnt mit“).
 * @param {{ displayName?: string, mailNickname?: string }} group
 * @param {string} term bereits getrimmter Suchbegriff
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
 * @param {Array<{ id?: string, displayName?: string, mailNickname?: string }>} groups
 * @param {string} term
 * @returns {Array<{ id: string, displayName: string, mailNickname: string }>} nach Anzeigename sortiert
 */
function filterAndSortGroupsByTerm(groups, term) {
    const t = normalizeSearchTerm(term);
    if (!t) return [];
    return (Array.isArray(groups) ? groups : [])
        .filter((g) => g && g.id && groupMatchesTerm(g, t))
        .map((g) => ({
            id: String(g.id),
            displayName: String(g.displayName || ''),
            mailNickname: String(g.mailNickname || '')
        }))
        .sort((a, b) => a.displayName.localeCompare(b.displayName, 'de'));
}

/**
 * @param {Array<{ ok?: boolean, skipped?: boolean }>} results
 * @returns {{ total: number, ok: number, fail: number }}
 */
function buildBulkActionSummary(results) {
    const list = Array.isArray(results) ? results : [];
    let ok = 0;
    let fail = 0;
    list.forEach((r) => {
        if (r && r.ok) ok += 1;
        else fail += 1;
    });
    return { total: list.length, ok, fail };
}

window.ms365TeamsArchivBulkLogic = {
    normalizeSearchTerm,
    groupMatchesTerm,
    filterAndSortGroupsByTerm,
    buildBulkActionSummary
};
