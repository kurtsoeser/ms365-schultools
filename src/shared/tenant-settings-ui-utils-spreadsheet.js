/**
 * Spreadsheet / XLSX / CSV IO helpers for tenant-settings-ui.
 */
import { normHeaderKey } from './tenant-settings-ui-utils-dialog.js';
function ensureXlsxReady() {
    return typeof XLSX !== 'undefined' && XLSX.utils && typeof XLSX.read === 'function';
}

function sheetToJsonRows(workbook) {
    const sheetName = workbook.SheetNames && workbook.SheetNames[0];
    if (!sheetName) return [];
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) return [];
    return XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

function workbookFirstSheetAoa(wb) {
    const sheetName = wb && wb.SheetNames && wb.SheetNames[0];
    const sheet = sheetName && wb.Sheets ? wb.Sheets[sheetName] : null;
    return sheet ? XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' }) : [];
}

/** Maximale Spaltenzahl in den ersten Zeilen (WebUntis-CSV ist oft tab-getrennt). */
function aoaMaxColumnCount(aoa, sampleRows) {
    const n = typeof sampleRows === 'number' ? sampleRows : 5;
    let max = 0;
    const rows = Array.isArray(aoa) ? aoa : [];
    for (let i = 0; i < rows.length && i < n; i++) {
        const row = rows[i];
        if (Array.isArray(row)) max = Math.max(max, row.length);
    }
    return max;
}

function csvTextToWorkbook(text) {
    let s = String(text || '');
    if (s.charCodeAt(0) === 0xfeff) s = s.slice(1);
    const delimiters = ['\t', ';', ','];
    let best = null;
    let bestCols = 0;
    delimiters.forEach(function (FS) {
        try {
            const wb = XLSX.read(s, { type: 'string', FS: FS });
            const cols = aoaMaxColumnCount(workbookFirstSheetAoa(wb));
            if (cols > bestCols) {
                bestCols = cols;
                best = wb;
            }
        } catch {
            /* nächstes Trennzeichen */
        }
    });
    return best || XLSX.read(s, { type: 'string', FS: ';' });
}

function parseCsvTextToJsonRows(text) {
    if (!ensureXlsxReady()) return [];
    return sheetToJsonRows(csvTextToWorkbook(text));
}

function downloadXlsxTemplate(filename, aoa, sheetName) {
    if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Daten');
    XLSX.writeFile(wb, filename);
    return true;
}

/** @param {{ name: string, aoa: any[][] }[]} sheets */
function downloadXlsxMultiSheet(filename, sheets) {
    if (!ensureXlsxReady() || typeof XLSX.writeFile !== 'function') return false;
    const wb = XLSX.utils.book_new();
    (sheets || []).forEach((sh) => {
        const rawName = String(sh.name || 'Daten').replace(/[:\\/?*[\]]/g, '-');
        const sn = rawName.slice(0, 31) || 'Daten';
        const ws = XLSX.utils.aoa_to_sheet(sh.aoa || []);
        XLSX.utils.book_append_sheet(wb, ws, sn);
    });
    XLSX.writeFile(wb, filename);
    return true;
}

/**
 * Generischer Excel/CSV-Import → JSON-Zeilen (erstes Arbeitsblatt).
 * @param {File} file
 * @param {(rows: object[]) => void} onRows
 * @param {(msg: string) => void} [onError]
 */
function importSpreadsheetFileToJsonRows(file, onRows, onError) {
    if (!file) return;
    if (!ensureXlsxReady()) {
        if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
        return;
    }
    const name = String(file.name || '').toLowerCase();
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            let jsonRows = [];
            if (name.endsWith('.csv')) {
                const buf = new Uint8Array(e.target.result);
                const tryDecoders = ['utf-8', 'windows-1252'];
                for (const enc of tryDecoders) {
                    try {
                        const text = new TextDecoder(enc).decode(buf);
                        jsonRows = parseCsvTextToJsonRows(text);
                        if (jsonRows.length) break;
                    } catch {
                        // ignore
                    }
                }
            } else {
                const data = new Uint8Array(e.target.result);
                const wb = XLSX.read(data, { type: 'array' });
                jsonRows = sheetToJsonRows(wb);
            }
            Promise.resolve(onRows(jsonRows || [])).catch(function (err) {
                if (onError) onError('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)));
            });
        } catch (err) {
            if (onError) onError('Import fehlgeschlagen: ' + (err?.message || String(err)));
        }
    };
    reader.readAsArrayBuffer(file);
}

function findWorksheetName(wb, aliases) {
    const names = wb.SheetNames || [];
    const want = (aliases || []).map((a) => normHeaderKey(a));
    for (const sn of names) {
        const nk = normHeaderKey(sn);
        if (want.indexOf(nk) >= 0) return sn;
    }
    for (const sn of names) {
        const nk = normHeaderKey(sn);
        for (const w of want) {
            if (!w) continue;
            if (nk === w || nk.indexOf(w) === 0 || w.indexOf(nk) === 0) return sn;
        }
    }
    return null;
}

function importFileToWorkbook(file, onWorkbook, onError) {
    if (!file) return;
    if (!ensureXlsxReady()) {
        if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
        return;
    }
    const name = String(file.name || '').toLowerCase();
    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            let wb;
            if (name.endsWith('.csv') || name.endsWith('.txt')) {
                wb = csvTextToWorkbook(String(e.target.result || ''));
            } else {
                wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
            }
            if (onWorkbook) onWorkbook(wb);
        } catch (err) {
            if (onError) onError('Import fehlgeschlagen: ' + (err?.message || String(err)));
        }
    };
    if (name.endsWith('.csv') || name.endsWith('.txt')) reader.readAsText(file, 'utf-8');
    else reader.readAsArrayBuffer(file);
}

function workbookToSheetList(wb, fileName) {
    const names = (wb && wb.SheetNames) || [];
    const out = [];
    names.forEach(function (sn) {
        const sheet = wb.Sheets[sn];
        if (!sheet) return;
        const aoa = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (!aoa || !aoa.length) return;
        out.push({
            name: (fileName ? fileName + ' — ' : '') + sn,
            aoa: aoa
        });
    });
    return out;
}

function readFilesToAllSheets(files) {
    const list = files && files.length != null ? Array.from(files) : files ? [files] : [];
    if (!ensureXlsxReady()) return Promise.reject(new Error('Excel-Bibliothek nicht geladen'));
    return Promise.all(
        list.map(function (file) {
            return new Promise(function (resolve, reject) {
                const name = String(file.name || '').toLowerCase();
                const reader = new FileReader();
                reader.onload = function (e) {
                    try {
                        let wb;
                        if (name.endsWith('.csv') || name.endsWith('.txt')) {
                            wb = csvTextToWorkbook(String(e.target.result || ''));
                        } else {
                            wb = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                        }
                        resolve(workbookToSheetList(wb, file.name || ''));
                    } catch (err) {
                        reject(err);
                    }
                };
                reader.onerror = function () {
                    reject(new Error('Datei konnte nicht gelesen werden: ' + (file.name || '')));
                };
                if (name.endsWith('.csv') || name.endsWith('.txt')) reader.readAsText(file);
                else reader.readAsArrayBuffer(file);
            });
        })
    ).then(function (nested) {
        const flat = [];
        nested.forEach(function (arr) {
            (arr || []).forEach(function (sh) {
                flat.push(sh);
            });
        });
        return flat;
    });
}

export {
    ensureXlsxReady,
    sheetToJsonRows,
    workbookFirstSheetAoa,
    aoaMaxColumnCount,
    csvTextToWorkbook,
    parseCsvTextToJsonRows,
    downloadXlsxTemplate,
    downloadXlsxMultiSheet,
    importSpreadsheetFileToJsonRows,
    findWorksheetName,
    importFileToWorkbook,
    workbookToSheetList,
    readFilesToAllSheets
};
