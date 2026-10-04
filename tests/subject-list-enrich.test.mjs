import { describe, expect, it } from 'vitest';
import {
    enrichSubjectsFromCatalogSources,
    mergeSubjectCatalogRows,
    deriveSubjectEntryFromGroup,
    subjectCatalogShrinkReport,
    subjectHintsFromArges,
    subjectHintsFromStructureRows,
    subjectHintsFromUnterrichtsbelegungRows,
    subjectHintFromBelegungRow,
    subjectCodeFromTeamNamePipe,
    belegungRowsFromAppDataContainer
} from '../src/shared/subject-list-enrich.js';

describe('subject-list-enrich', () => {
    it('deriveSubjectEntryFromGroup liest fach-Alias und Anzeigenamen', () => {
        expect(deriveSubjectEntryFromGroup({ mailNickname: 'fach-d', displayName: 'Fachgruppe Deutsch' })).toEqual({
            code: 'D',
            name: 'Fachgruppe Deutsch'
        });
        expect(
            deriveSubjectEntryFromGroup({ mailNickname: 'fachgruppe-mathematik', displayName: 'Mathematik' })
        ).toMatchObject({ code: 'MATHEMATIK' });
        expect(deriveSubjectEntryFromGroup({ displayName: 'Fach Englisch' }, { mailPrefix: 'fach' })).toMatchObject({
            code: 'ENGLISCH'
        });
    });

    it('mergeSubjectCatalogRows ist additiv', () => {
        const r = mergeSubjectCatalogRows(
            [{ code: 'D', name: 'Deutsch' }],
            [{ code: 'M', name: 'Mathematik' }, { code: 'D', name: 'Deutsch neu' }]
        );
        expect(r.changed).toBe(true);
        expect(r.stats.added).toBe(1);
        expect(r.subjects.find((s) => s.code === 'D').name).toBe('Deutsch');
        expect(r.subjects.find((s) => s.code === 'M').name).toBe('Mathematik');
    });

    it('subjectHintsFromArges sammelt Zuordnungs-Kürzel', () => {
        const hints = subjectHintsFromArges([{ code: 'D', name: 'Deutsch-ARGE', subjects: ['D', 'LAT'] }]);
        expect(hints.map((h) => h.code).sort()).toEqual(['D', 'LAT']);
    });

    it('enrichSubjectsFromCatalogSources ergänzt fehlende Fächer', () => {
        const r = enrichSubjectsFromCatalogSources([{ code: 'D', name: 'Deutsch' }], {
            structureRows: [{ typ: 'Kursteam', ktFach: 'M', bezeichnung: 'Mathematik KT' }],
            arges: [{ code: 'X', subjects: ['E'] }],
            unterrichtsbelegungRows: [{ fach: 'LAT', teamName: 'SJ26 | 1A | LAT' }]
        });
        expect(r.subjects.map((s) => s.code).sort()).toEqual(['D', 'E', 'LAT', 'M']);
    });

    it('subjectHintsFromUnterrichtsbelegungRows nutzt Fach und Teamnamen', () => {
        expect(subjectCodeFromTeamNamePipe('SJ26 | 5A | M')).toBe('M');
        expect(subjectHintFromBelegungRow({ fach: 'Mathematik', teamName: 'SJ26 | 5A | M' })).toEqual({
            code: 'M',
            name: 'Mathematik'
        });
        const hints = subjectHintsFromUnterrichtsbelegungRows([
            { fach: 'BW', teamName: 'x' },
            { fach: 'D', teamName: 'SJ26 | 1A | D' }
        ]);
        expect(hints.map((h) => h.code).sort()).toEqual(['BW', 'D']);
    });

    it('belegungRowsFromAppDataContainer sammelt alle Schuljahre', () => {
        const rows = belegungRowsFromAppDataContainer({
            years: {
                byLabel: {
                    '2025/26': {
                        unterrichtsbelegung: { rows: [{ fach: 'M', klasse: '1A' }] }
                    },
                    '2024/25': {
                        unterrichtsbelegung: { rows: [{ fach: 'D', klasse: '2B' }] }
                    }
                }
            }
        });
        expect(rows).toHaveLength(2);
    });

    it('subjectCatalogShrinkReport erkennt fehlende Kürzel', () => {
        const r = subjectCatalogShrinkReport(
            [{ code: 'D' }, { code: 'M' }, { code: 'LAT' }],
            [{ code: 'D' }, { code: 'M' }]
        );
        expect(r.wouldShrink).toBe(true);
        expect(r.removed).toEqual(['LAT']);
    });
});
