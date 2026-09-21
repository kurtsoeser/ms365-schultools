/**
 * Geteilte Unterrichte (gleiche Lehrkraft+Fach+Schülergruppe bzw. gleiche Mehrklassen-Zelle)
 * zu einer Zeile zusammenführen → ein Team statt eins pro Klasse.
 * @param {Array} rows
 * @returns {{ rows: Array, mergedCount: number }}
 */
function mergeSharedLessonRows(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const byKey = new Map();
    const order = [];

    list.forEach((row) => {
        const lehrer = String((row && row.lehrer) || '')
            .toUpperCase()
            .trim();
        const fach = String((row && row.fach) || '')
            .toUpperCase()
            .trim();
        const gruppe = String((row && row.gruppe) || '').trim();
        const parts = Array.isArray(row && row.klassenParts)
            ? row.klassenParts.map((c) => String(c || '').trim()).filter(Boolean)
            : String((row && row.klasse) || '')
                  .split(/[,;~]+/)
                  .map((c) => c.trim())
                  .filter(Boolean);

        // Merge-Schlüssel: WebUntis-Schülergruppe (ETH_5AK5BK_…) oder Mehrklassen-Rohzelle
        let key = '';
        if (gruppe && lehrer && fach) {
            key = 'g|' + lehrer + '|' + fach + '|' + gruppe.toUpperCase();
        } else if (parts.length > 1 && lehrer && fach) {
            const sorted = parts
                .slice()
                .map((c) => c.toUpperCase())
                .sort()
                .join('~');
            key = 'c|' + lehrer + '|' + fach + '|' + sorted;
        }

        if (!key) {
            order.push({ type: 'single', row });
            return;
        }

        if (!byKey.has(key)) {
            const entry = {
                type: 'merge',
                row: Object.assign({}, row),
                classSet: new Set(parts.length ? parts : [String(row.klasse || '').trim()].filter(Boolean))
            };
            byKey.set(key, entry);
            order.push(entry);
        } else {
            const entry = byKey.get(key);
            parts.forEach((p) => entry.classSet.add(p));
            if (!entry.row.gruppe && gruppe) entry.row.gruppe = gruppe;
            if (!entry.row.klassenRaw && row.klassenRaw) entry.row.klassenRaw = row.klassenRaw;
        }
    });

    let mergedCount = 0;
    const out = order.map((item) => {
        if (item.type === 'single') return item.row;
        const classes = Array.from(item.classSet).filter(Boolean);
        if (classes.length > 1) {
            mergedCount += 1;
            item.row.klasse = classes.join(',');
            item.row.klassenParts = classes.slice();
            item.row.sharedLesson = true;
        } else if (classes.length === 1) {
            item.row.klasse = classes[0];
            item.row.klassenParts = classes.slice();
        }
        return item.row;
    });

    return { rows: out, mergedCount };
}

/**
 * Mehrklassen-Zeilen wieder in Einzelklassen aufteilen (wenn Merge-Option aus).
 * @param {Array} rows
 * @returns {{ rows: Array, splitCount: number }}
 */
function expandSharedLessonRows(rows) {
    const list = Array.isArray(rows) ? rows : [];
    const out = [];
    let splitCount = 0;
    let nextId = Date.now();

    list.forEach((row) => {
        const parts = Array.isArray(row && row.klassenParts)
            ? row.klassenParts.map((c) => String(c || '').trim()).filter(Boolean)
            : String((row && row.klasse) || '')
                  .split(/[,;~]+/)
                  .map((c) => c.trim())
                  .filter(Boolean);

        if (parts.length <= 1) {
            out.push(row);
            return;
        }

        splitCount += 1;
        parts.forEach((klasse, i) => {
            out.push(
                Object.assign({}, row, {
                    id: row.id != null && i === 0 ? row.id : nextId++,
                    klasse,
                    klassenParts: [klasse],
                    sharedLesson: false
                })
            );
        });
    });

    return { rows: out, splitCount };
}

/**
 * Rohdaten → gefilterte Zeilen (Fach-Ausschluss, Klassenpflicht, optional Duplikate).
 * @param {Array} rawData
 * @param {string[]} excludeSubjects bereits normalisierte Fach-Kürzel (z. B. upper case)
 * @param {boolean} removeDuplicates
 * @param {object} [options]
 * @param {boolean} [options.normalizeNumberedSubjects] OMAI1 → fach OMAI, gruppe 1 (wenn leer)
 * @param {function} [options.normalizeNumberedSubjectFields] Inject aus Subject-Logic
 * @param {boolean} [options.mergeSharedLessons] Mehrklassen-/Gruppen-Unterricht → 1 Zeile
 * @returns {{ filtered: Array, removedByFilter: number, removedByDuplicate: number, normalizedCount: number, mergedSharedCount: number }}
 */
function applyRowFilters(rawData, excludeSubjects, removeDuplicates, options) {
    const ex = Array.isArray(excludeSubjects) ? excludeSubjects : [];
    const opts = options && typeof options === 'object' ? options : {};
    const normalizeNumbered = !!opts.normalizeNumberedSubjects;
    const normalizeFn =
        typeof opts.normalizeNumberedSubjectFields === 'function'
            ? opts.normalizeNumberedSubjectFields
            : null;
    const mergeShared = !!opts.mergeSharedLessons;

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

    let mergedSharedCount = 0;
    if (mergeShared) {
        const m = mergeSharedLessonRows(filtered);
        filtered = m.rows;
        mergedSharedCount = m.mergedCount;
    } else {
        const e = expandSharedLessonRows(filtered);
        filtered = e.rows;
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

    return {
        filtered,
        removedByFilter,
        removedByDuplicate,
        normalizedCount,
        mergedSharedCount
    };
}

window.ms365KursteamFilterLogic = {
    applyRowFilters,
    mergeSharedLessonRows,
    expandSharedLessonRows
};
