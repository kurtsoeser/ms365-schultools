/**
 * CSV-Export für Freistellungen.
 */
import { rowsToCsv, downloadCsv } from '../../shared/utils/csv.js';
import { formatDeDate, statusLabel } from './freistellung-planer-state.js';
import { inclusiveDayCount } from './freistellung-planer-logic.js';

const COLUMNS = [
    { label: 'Schüler/in', value: (r) => r.schuelerName || r.titel || '' },
    { label: 'Klasse', value: (r) => r.klasse || '' },
    { label: 'Beginn', value: (r) => formatDeDate(r.beginn) },
    { label: 'Ende', value: (r) => formatDeDate(r.ende) },
    { label: 'Tage', value: (r) => inclusiveDayCount(r.beginn, r.ende) ?? '' },
    { label: 'Status', value: (r) => statusLabel(r.status) },
    { label: 'Kategorie', value: (r) => r.kategorie || '' },
    { label: 'Klassenvorstand', value: (r) => r.kvName || r.kvEmail || '' },
    { label: 'KV-E-Mail', value: (r) => r.kvEmail || '' },
    { label: 'Beschreibung', value: (r) => r.beschreibung || '' },
    { label: 'Bemerkungen', value: (r) => r.bemerkungen || '' },
    { label: 'Beantragt von', value: (r) => r.authorEmail || r.beantragtVon || '' },
    { label: 'Genehmigungspfad', value: (r) => r.approvalLabel || '' }
];

/**
 * @param {object[]} items
 */
export function buildFreistellungCsv(items) {
    return rowsToCsv(items || [], COLUMNS, { bom: true, sep: ';' });
}

/**
 * @param {object[]} items
 * @param {string} [filename]
 */
export function downloadFreistellungCsv(items, filename) {
    const csv = buildFreistellungCsv(items);
    const name =
        filename ||
        'freistellungen-' + new Date().toISOString().slice(0, 10) + '.csv';
    downloadCsv(name, csv);
}
