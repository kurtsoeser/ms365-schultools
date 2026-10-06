import { describe, expect, it } from 'vitest';
import { bundleFromWebuntisHandoff } from '../src/shared/stammdaten-import-bundle.js';
import {
    ADAPTER_WEBUNTIS_HANDOFF,
    ADAPTER_SIS_FILE,
    ADAPTER_EDTECH_GENERIC,
    mergeStammdatenImport,
    listStammdatenImportAdapterIds,
    detectStammdatenImportAdapter,
    getStammdatenImportAdapter
} from '../src/shared/stammdaten-import-pipeline.js';
import '../src/shared/stammdaten-import-adapters.js';

describe('stammdaten-import-pipeline', () => {
    it('registriert WebUntis-, SIS- und EdTech-Adapter', () => {
        const ids = listStammdatenImportAdapterIds();
        expect(ids).toContain(ADAPTER_WEBUNTIS_HANDOFF);
        expect(ids).toContain(ADAPTER_SIS_FILE);
        expect(ids).toContain(ADAPTER_EDTECH_GENERIC);
        expect(getStammdatenImportAdapter(ADAPTER_EDTECH_GENERIC).label).toMatch(/EdTech/i);
    });

    it('bundleFromWebuntisHandoff übernimmt counts', () => {
        const b = bundleFromWebuntisHandoff({ counts: { students: 3, classes: 1 } });
        expect(b.counts.students).toBe(3);
        expect(b.adapterId).toBe('webuntis-handoff');
    });

    it('mergeStammdatenImport merged leere Zeilen bei leerem Payload', () => {
        const merged = mergeStammdatenImport(
            ADAPTER_WEBUNTIS_HANDOFF,
            { teachersLines: 'A;B', studentsLines: '', subjectsLines: '', classesLines: '' },
            { teachersLines: '', studentsLines: '', subjectsLines: '', classesLines: '' },
            {
                webuntis: {},
                schoolSis: {},
                parseTeachersLines: () => [],
                parseStudentsLines: () => [],
                parseSubjectsLines: () => [],
                parseClassesLines: () => []
            }
        );
        expect(merged).toBeTruthy();
        expect(merged.lines).toBeTruthy();
    });

    it('SIS-Adapter merged Schülerlisten', () => {
        const merged = mergeStammdatenImport(
            ADAPTER_SIS_FILE,
            { studentsLines: '1A;Alt;alt@schule.at', teachersLines: '', subjectsLines: '', classesLines: '' },
            {
                records: [{ klasse: '1A', name: 'Neu', email: 'neu@schule.at' }],
                mode: 'merge'
            },
            {
                schoolSis: {
                    applySisImport: function (existing, incoming) {
                        return existing.concat(incoming);
                    },
                    diffSisImport: function () {
                        return { counts: { added: 1, updated: 0, removed: 0 }, conflicts: [] };
                    },
                    recordsToSemicolonLines: function (rows) {
                        return rows.map((r) => [r.klasse, r.name, r.email].join(';')).join('\n');
                    }
                },
                parseStudentsLines: function (text) {
                    return String(text || '')
                        .split('\n')
                        .filter(Boolean)
                        .map(function (line) {
                            const p = line.split(';');
                            return { klasse: p[0], name: p[1], email: p[2] };
                        });
                }
            }
        );
        expect(merged.lines.studentsLines).toContain('Neu');
        expect(merged.studentDiff.counts.added).toBe(1);
    });

    it('EdTech-Stub liefert Hinweis ohne Datenverlust', () => {
        const merged = mergeStammdatenImport(
            ADAPTER_EDTECH_GENERIC,
            { studentsLines: '1A;X;x@y.at', teachersLines: '', subjectsLines: '', classesLines: '' },
            {},
            {}
        );
        expect(merged.lines.studentsLines).toContain('1A');
        expect(merged.message).toMatch(/EdTech/i);
    });

    it('detectStammdatenImportAdapter erkennt leere Sheets', () => {
        const d = detectStammdatenImportAdapter([]);
        expect(d.confidence).toBe('none');
    });
});
