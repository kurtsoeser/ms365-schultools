import { describe, it, expect } from 'vitest';
import {
    parseMaturaTerminPage,
    periodenForTerminJahr,
    datumForExamPart,
    htmlToMaturaText,
    mergePruefungstermine,
    refreshPruefungstermineFromWeb,
    MATURA_PAGE_TEXT
} from '../src/tools/srdp-anmeldung/srdp-pruefungstermine.js';
import { buildPruefplanDetail } from '../src/tools/srdp-anmeldung/srdp-anmeldung-logic.js';

describe('sRDP-Prüfungstermine', () => {
    const catalog = parseMaturaTerminPage(MATURA_PAGE_TEXT);

    it('liest Haupt-, Herbst- und Wintertermin 2027 aus der Matura-Seite', () => {
        const bundle = periodenForTerminJahr(2027, catalog);
        expect(bundle.schuljahr).toBe('2026/27');
        expect(bundle.perioden.map((row) => row.id)).toEqual(['haupt', 'herbst', 'winter']);
        const deutsch = bundle.perioden[0].block.faecher.find((row) => row.key === 'deutsch');
        const englisch = bundle.perioden[0].block.faecher.find((row) => row.key === 'englisch');
        const mathematik = bundle.perioden[0].block.faecher.find((row) => row.key === 'mathematik');
        expect(deutsch.daten).toEqual(['2027-05-05']);
        expect(englisch.daten).toEqual(['2027-05-11']);
        expect(mathematik.daten).toEqual(['2027-05-10']);
        expect(bundle.perioden[1].block.jahr).toBe('2027');
        expect(bundle.perioden[2].block.jahr).toBe('2028');
        const kompensation = bundle.perioden[0].block.faecher.find((row) => row.key === 'kompensation');
        expect(kompensation.daten).toEqual(['2027-06-01', '2027-06-02']);
    });

    it('hängt das Datum an Deutsch, Englisch und Mathematik, nicht an BFK', () => {
        const detail = buildPruefplanDetail(
            { variante: 'Variante 1', lfs: 'ENWS' },
            'hak',
            { terminJahr: '2027' }
        );
        const deutsch = detail.schriftlich.find((row) => row.part.startsWith('D'));
        const lfs = detail.schriftlich.find((row) => row.part.startsWith('LFS'));
        const bfk = detail.schriftlich.find((row) => row.part.startsWith('BFK'));
        expect(deutsch.datumIso).toBe('2027-05-05');
        expect(deutsch.datum).toBe('5. Mai 2027');
        expect(lfs.datumIso).toBe('2027-05-11');
        expect(lfs.fach).toBe('Englisch');
        expect(bfk.datum).toBe('');
        expect(bfk.datumHinweis).toMatch(/Schule/);
        expect(datumForExamPart('AM (4,5h)', '', '2027').iso).toBe('2027-05-10');
        const withBfk = buildPruefplanDetail(
            { variante: 'Variante 1', lfs: 'ENWS' },
            'hak',
            { terminJahr: '2027', bfkDatum: '2027-05-06' }
        );
        const bfkDated = withBfk.schriftlich.find((row) => row.part.startsWith('BFK'));
        expect(bfkDated.datumIso).toBe('2027-05-06');
        expect(bfkDated.fach).toBe('BFK schriftlich');
    });

    it('liest einen späteren Haupttermin aus dem HTML der Matura-Seite', async () => {
        const html =
            '<span class="headline">Schuljahr 2027/28</span>' +
            '<table><tr><td colspan="2"><h3>Haupttermin 2028</h3></td></tr>' +
            '<tr><td>Deutsch (AHS/BHS/BRP)</td><td>3. Mai 2028</td></tr>' +
            '<tr><td>Englisch (AHS/BHS/BRP)</td><td>9. Mai 2028</td></tr></table>';
        const parsed = parseMaturaTerminPage(htmlToMaturaText(html));
        const bundle = periodenForTerminJahr(2028, mergePruefungstermine(parsed, catalog));
        expect(bundle.schuljahr).toBe('2027/28');
        expect(bundle.perioden[0].block.faecher.find((row) => row.key === 'deutsch').daten).toEqual([
            '2028-05-03'
        ]);

        const result = await refreshPruefungstermineFromWeb({
            force: true,
            fetch: async () => ({
                ok: true,
                headers: { get: () => 'application/json' },
                json: async () => ({ html, fetchedAt: '2028-01-01T00:00:00.000Z' })
            })
        });
        expect(result.source).toBe('live');
        expect(periodenForTerminJahr(2028).perioden[0].block.faecher[0].daten[0]).toBe('2028-05-03');
    });
});
