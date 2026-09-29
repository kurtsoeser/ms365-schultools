/**
 * Gemeinsame CSV-Helfer (Analyse 02 Querschnitt).
 */

/** CSV-Zelle mit Semikolon-Trennung (DE/AT). */
export function csvEscape(cell) {
    const s = String(cell ?? '');
    if (/[",\n\r;]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
    return s;
}

/**
 * @param {object[]} rows
 * @param {Array<{ label: string, value: (row: object) => unknown }>} columns
 * @param {{ bom?: boolean, sep?: string }} [opts]
 */
export function rowsToCsv(rows, columns, opts) {
    const o = opts && typeof opts === 'object' ? opts : {};
    const sep = o.sep != null ? String(o.sep) : ';';
    const bom = o.bom !== false;
    const header = columns.map((c) => csvEscape(c.label)).join(sep);
    const list = Array.isArray(rows) ? rows : [];
    const lines = list.map((row) => columns.map((c) => csvEscape(c.value(row))).join(sep));
    return (bom ? '\uFEFF' : '') + header + '\n' + lines.join('\n');
}

/**
 * Lädt CSV als Download im Browser.
 * @param {string} filename
 * @param {string} csvText
 */
export function downloadCsv(filename, csvText) {
    if (typeof document === 'undefined') return;
    const blob = new Blob([csvText], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = String(filename || 'export.csv');
    a.click();
    setTimeout(function () {
        try {
            URL.revokeObjectURL(a.href);
        } catch {
            /* ignore */
        }
    }, 2000);
}
