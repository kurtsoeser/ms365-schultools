import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

function loadImportMappingSandbox() {
    const sandbox = {
        console,
        document: {
            readyState: 'complete',
            getElementById: () => null,
            addEventListener: () => {}
        }
    };
    sandbox.window = sandbox;
    createContext(sandbox);
    const full = join(root, 'src/tools/kursteams/kursteam-import.js');
    runInContext(readFileSync(full, 'utf8'), sandbox, { filename: full });
    return sandbox.ms365Kursteam;
}

describe('kursteam WebUntis ExportLessons mapping', () => {
    it('erkennt subject/teacher/klassen aus ExportLessons', () => {
        const ns = loadImportMappingSandbox();
        const mapped = ns.mapImportedLessonRow({
            subject: 'ANWA',
            teacher: 'RIN',
            klassen: '2BS',
            periods: 3,
            room: '214',
            foreignKey: ''
        });
        expect(mapped).toEqual({
            lehrer: 'RIN',
            fach: 'ANWA',
            klasseRaw: '2BS',
            gruppe: '',
            profile: 'webuntis-lessons'
        });
    });

    it('splittet mehrere Klassen in klassen', () => {
        const ns = loadImportMappingSandbox();
        expect(ns.splitKlassenCell('2AS, 2BS')).toEqual(['2AS', '2BS']);
        expect(ns.splitKlassenCell('1DK~1EK')).toEqual(['1DK', '1EK']);
        expect(ns.splitKlassenCell('5AK~5BK~5CK')).toEqual(['5AK', '5BK', '5CK']);
    });

    it('mappt studentgroup aus ExportLessons', () => {
        const ns = loadImportMappingSandbox();
        const mapped = ns.mapImportedLessonRow({
            subject: 'ETH',
            teacher: 'HAGMU',
            klassen: '5AK~5BK',
            studentgroup: 'ETH_5AK5BK_HAGMU'
        });
        expect(mapped.gruppe).toBe('ETH_5AK5BK_HAGMU');
        expect(mapped.klasseRaw).toBe('5AK~5BK');
    });

    it('processImportedData baut Zeilen aus ExportLessons', () => {
        const ns = loadImportMappingSandbox();
        const applied = [];
        ns.applyWebuntisRows = (rows) => {
            applied.push(rows);
        };
        ns.showToast = () => {};

        ns.processImportedData([
            { subject: 'D', teacher: 'MEI', klassen: '1A', periods: 2 },
            { subject: 'M', teacher: 'MEI', klassen: '1A,1B', periods: 1 },
            { subject: 'ETH', teacher: 'HAG', klassen: '2AK~2CK', studentgroup: 'ETH_2AK2CK_HAG', periods: 2 }
        ]);

        expect(applied).toHaveLength(1);
        expect(applied[0]).toHaveLength(5);
        expect(applied[0].map((r) => `${r.lehrer}|${r.fach}|${r.klasse}|${r.gruppe}`)).toEqual([
            'MEI|D|1A|',
            'MEI|M|1A|',
            'MEI|M|1B|',
            'HAG|ETH|2AK|ETH_2AK2CK_HAG',
            'HAG|ETH|2CK|ETH_2AK2CK_HAG'
        ]);
    });
});
