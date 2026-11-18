/**
 * Ergebnis-Export (CSV) – Vorbereitung für lokalen Access-Import (Linz-Auhof Schülerakte).
 */

import { ergebnisListenProAngebot } from './wahlfach-zuteilung-logic.js';

/**
 * @param {object} state
 * @returns {string[][]}
 */
export function buildErgebnisExportAoa(state) {
    const s = state || {};
    const schuljahr = (s.runde && s.runde.schuljahr) || '';
    const rundeTitel = (s.runde && s.runde.titel) || '';
    const header = [
        'Schuljahr',
        'Runde',
        'Angebot',
        'AngebotId',
        'KapazitaetMin',
        'KapazitaetMax',
        'Belegt',
        'SchuelerName',
        'Klasse',
        'AnmeldungId',
        'Wunsch1',
        'Wunsch2',
        'Wunsch3',
        'ZugeteiltAm',
        'ZugeteiltVon',
        'AccessExportKey'
    ];
    /** @type {string[][]} */
    const rows = [header];
    const angebotById = Object.fromEntries((s.angebote || []).map((a) => [a.id, a]));
    const label = (id) => (angebotById[id] && angebotById[id].bezeichnung) || '';

    const listen = ergebnisListenProAngebot(s.angebote, s.anmeldungen, s.zuteilungen);
    for (const L of listen) {
        const ang = L.angebot;
        if (!L.teilnehmende.length) {
            rows.push([
                schuljahr,
                rundeTitel,
                ang.bezeichnung,
                ang.id,
                String(ang.kapazitaetMin ?? ''),
                String(ang.kapazitaetMax ?? ''),
                '0',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                [schuljahr, ang.id, ''].join('|')
            ]);
            continue;
        }
        for (const t of L.teilnehmende) {
            const a = t.anmeldung;
            const z = t.zuteilung || {};
            rows.push([
                schuljahr,
                rundeTitel,
                ang.bezeichnung,
                ang.id,
                String(ang.kapazitaetMin ?? ''),
                String(ang.kapazitaetMax ?? ''),
                String(L.belegt),
                a.schuelerName || '',
                a.klasse || '',
                a.id || '',
                label(a.wunsch1Id),
                label(a.wunsch2Id),
                label(a.wunsch3Id),
                z.zugeteiltAm ? String(z.zugeteiltAm).slice(0, 19) : '',
                z.zugeteiltVon || '',
                [schuljahr, ang.id, a.id || ''].join('|')
            ]);
        }
    }
    return rows;
}

/**
 * @param {string[][]} aoa
 * @returns {string}
 */
export function aoaToCsv(aoa) {
    return (aoa || [])
        .map((row) =>
            row
                .map((cell) => {
                    const s = String(cell ?? '');
                    if (/[;"\n\r]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
                    return s;
                })
                .join(';')
        )
        .join('\r\n');
}

/**
 * @param {object} state
 * @param {string} [filename]
 */
export function downloadErgebnisCsv(state, filename) {
    const aoa = buildErgebnisExportAoa(state);
    const csv = '\uFEFF' + aoaToCsv(aoa);
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
    const a = document.createElement('a');
    const sj = String((state && state.runde && state.runde.schuljahr) || 'export').replace(
        /[^\w.-]+/g,
        '_'
    );
    a.href = URL.createObjectURL(blob);
    a.download = filename || `wahlfach-ergebnis-${sj}.csv`;
    a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
}
