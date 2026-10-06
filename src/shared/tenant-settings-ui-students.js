/**
 * Schüler-Tab Helfer für tenant-settings-ui (Phase 7b).
 * Pure Zeilen-Konvertierung – Render bleibt in der UI-Closure.
 */

function normStr(v) {
    return String(v ?? '').trim();
}

/**
 * Format kompatibel zu parseLinesToStudents / school-sis recordsToSemicolonLines.
 * @param {Array<{ klasse?: string, name?: string, email?: string, externalId?: string, parentPairs?: Array<{ name?: string, email?: string }> }>} rows
 * @returns {string}
 */
export function studentsToLines(rows) {
    return (rows || [])
        .map(function (x) {
            const parts = [
                normStr(x.klasse),
                normStr(x.name),
                normStr(x.email).toLowerCase()
            ];
            if (normStr(x.externalId)) parts.push('#id:' + normStr(x.externalId));
            const pairs = Array.isArray(x.parentPairs) ? x.parentPairs : [];
            if (!pairs.length) return parts.join(';').trim();
            pairs.forEach(function (p) {
                parts.push(normStr(p && p.name));
                parts.push(normStr(p && p.email).toLowerCase());
            });
            return parts.join(';').trim();
        })
        .filter(Boolean)
        .join('\n');
}

/**
 * @param {Array<{ klasse?: string, name?: string, email?: string }>} rows
 * @param {string} key
 * @param {1|-1} dir
 */
export function sortStudentRows(rows, key, dir) {
    const d = dir === -1 ? -1 : 1;
    const k = String(key || 'name');
    return (rows || []).slice().sort(function (a, b) {
        const av = normStr(a && a[k]).toLowerCase();
        const bv = normStr(b && b[k]).toLowerCase();
        if (av < bv) return -1 * d;
        if (av > bv) return 1 * d;
        return 0;
    });
}
