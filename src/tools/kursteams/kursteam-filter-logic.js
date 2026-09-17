/**
 * Rohdaten → gefilterte Zeilen (Fach-Ausschluss, Klassenpflicht, optional Duplikate).
 * @param {Array} rawData
 * @param {string[]} excludeSubjects bereits normalisierte Fach-Kürzel (z. B. upper case)
 * @param {boolean} removeDuplicates
 * @param {object} [options]
 * @param {boolean} [options.normalizeNumberedSubjects] OMAI1 → fach OMAI, gruppe 1 (wenn leer)
 * @param {function} [options.normalizeNumberedSubjectFields] Inject aus Subject-Logic
 * @returns {{ filtered: Array, removedByFilter: number, removedByDuplicate: number, normalizedCount: number }}
 */
function applyRowFilters(rawData, excludeSubjects, removeDuplicates, options) {
    const ex = Array.isArray(excludeSubjects) ? excludeSubjects : [];
    const opts = options && typeof options === 'object' ? options : {};
    const normalizeNumbered = !!opts.normalizeNumberedSubjects;
    const normalizeFn =
        typeof opts.normalizeNumberedSubjectFields === 'function'
            ? opts.normalizeNumberedSubjectFields
            : null;

    let filtered = (rawData || []).filter((row) => {
        if (!row.fach || !row.lehrer) return false;
        const fach = row.fach.toUpperCase().trim();
        if (ex.includes(fach)) return false;
        if (!row.klasse || row.klasse.trim() === '') return false;
        return true;
    });

    const countAfterPass1 = filtered.length;
    const removedByFilter = (rawData || []).length - countAfterPass1;

    let normalizedCount = 0;
    if (normalizeNumbered && normalizeFn) {
        filtered = filtered.map((row) => {
            const n = normalizeFn(row.fach, row.gruppe);
            if (!n || !n.changed) return row;
            normalizedCount += 1;
            return Object.assign({}, row, { fach: n.fach, gruppe: n.gruppe });
        });
    }

    const countBeforeDedup = filtered.length;

    if (removeDuplicates) {
        const seen = new Set();
        filtered = filtered.filter((row) => {
            const key = `${row.klasse}-${row.fach}-${row.lehrer}-${row.gruppe}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    }

    const removedByDuplicate = countBeforeDedup - filtered.length;

    return { filtered, removedByFilter, removedByDuplicate, normalizedCount };
}

window.ms365KursteamFilterLogic = {
    applyRowFilters
};
