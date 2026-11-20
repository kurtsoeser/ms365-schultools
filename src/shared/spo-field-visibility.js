/**
 * Hilfen für SharePoint-Formular-Sichtbarkeit (ohne SPO-REST-I/O).
 */

/** Formeln wie „=false“, die das Feld dauerhaft im Formular ausblenden. */
export function isAlwaysHideShowFormula(formula) {
    const s = String(formula || '')
        .trim()
        .toLowerCase()
        .replace(/\s+/g, '');
    return s === '=false' || s === "='false'" || s === "=if(true,'false','true')";
}
