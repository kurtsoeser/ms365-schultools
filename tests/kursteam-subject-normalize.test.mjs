import { describe, it, expect, beforeAll } from 'vitest';
import { loadScript } from './kursteams-vm.mjs';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

function loadScripts(relativePaths) {
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    relativePaths.forEach((rel) => {
        const full = join(root, rel);
        runInContext(readFileSync(full, 'utf8'), sandbox, { filename: full });
    });
    return sandbox;
}

/** @type {import('../src/tools/kursteams/kursteam-subject-filter-logic.js')} */
let kursteamSubjectFilter;

describe('kursteam-subject-filter-logic', () => {
    beforeAll(async () => {
        if (!globalThis.window) globalThis.window = globalThis;
        await import('../src/tools/kursteams/kursteam-subject-filter-logic.js');
        kursteamSubjectFilter = globalThis.ms365KursteamSubjectFilterLogic;
    });

    it('normalizeNumberedSubjectFields verschiebt Ziffer in Gruppe', () => {
        const { normalizeNumberedSubjectFields } = kursteamSubjectFilter;

        expect(normalizeNumberedSubjectFields('OMAI1', '')).toEqual({
            fach: 'OMAI',
            gruppe: '1',
            changed: true,
            suffix: '1'
        });
        expect(normalizeNumberedSubjectFields('OMAI1', 'G2').changed).toBe(true);
        expect(normalizeNumberedSubjectFields('OMAI1', 'G2').gruppe).toBe('G2');
        expect(normalizeNumberedSubjectFields('D', '').changed).toBe(false);
    });

    it('normalizeNumberedSubjectFields verschiebt Ü in Gruppe', () => {
        const { normalizeNumberedSubjectFields } = kursteamSubjectFilter;
        expect(normalizeNumberedSubjectFields('RWÜ', '')).toEqual({
            fach: 'RW',
            gruppe: 'Ü',
            changed: true,
            suffix: 'Ü'
        });
    });
});

describe('kursteam-filter-logic numbered normalize', () => {
    it('normalisiert vor Dedup und zählt Änderungen', async () => {
        if (!globalThis.window) globalThis.window = globalThis;
        await import('../src/tools/kursteams/kursteam-subject-filter-logic.js');
        const ctx = loadScripts(['src/tools/kursteams/kursteam-filter-logic.js']);
        ctx.ms365KursteamSubjectFilterLogic = globalThis.ms365KursteamSubjectFilterLogic;
        const { applyRowFilters } = ctx.ms365KursteamFilterLogic;
        const { normalizeNumberedSubjectFields } = ctx.ms365KursteamSubjectFilterLogic;

        const raw = [
            { klasse: '1A', fach: 'OMAI1', lehrer: 'A', gruppe: '' },
            { klasse: '1A', fach: 'OMAI2', lehrer: 'B', gruppe: '' },
            { klasse: '1A', fach: 'ORD', lehrer: 'Z', gruppe: '' }
        ];

        const r = applyRowFilters(raw, ['ORD'], true, {
            normalizeNumberedSubjects: true,
            normalizeNumberedSubjectFields
        });
        expect(r.filtered).toHaveLength(2);
        expect(r.normalizedCount).toBe(2);
        expect(r.filtered.map((x) => `${x.fach}|${x.gruppe}`)).toEqual(['OMAI|1', 'OMAI|2']);
    });
});

describe('kursteam team-build strip digits', () => {
    it('stripSubjectTrailingDigits nutzt Basis-Fach und Gruppe', async () => {
        if (!globalThis.window) globalThis.window = globalThis;
        await import('../src/tools/kursteams/kursteam-subject-filter-logic.js');
        const ctx = loadScripts([
            'src/shared/ms365-module-guard.js',
            'src/tools/kursteams/kursteam-team-names.js',
            'src/tools/kursteams/kursteam-utils.js',
            'src/tools/kursteams/kursteam-team-build.js'
        ]);
        const KTB = ctx.ms365KursteamTeamBuild;
        const KT = ctx.ms365KursteamTeamNames;
        ctx.ms365KursteamSubjectFilterLogic = globalThis.ms365KursteamSubjectFilterLogic;
        const KS = ctx.ms365KursteamSubjectFilterLogic;
        const ns = ctx.ms365Kursteam;

        const pattern = [
            { type: 'yearPrefix' },
            { type: 'text', value: ' | ' },
            { type: 'klasse' },
            { type: 'text', value: ' | ' },
            { type: 'fach' },
            { type: 'text', value: ' | ' },
            { type: 'gruppe' }
        ];

        const teams = KTB.buildTeamEntriesFromRows(
            [{ klasse: '1AK', fach: 'OMAI1', lehrer: 'PICHL', gruppe: '' }],
            {
                yearPrefix: 'SJ26',
                emailDomain: '@schule.at',
                separator: ' | ',
                pattern,
                combineClassNames: ns.combineClassNames,
                buildGruppenmailBase: ns.buildGruppenmailBase,
                formatKlasseSegmentForGruppenmail: ns.formatKlasseSegmentForGruppenmail,
                sanitizeGruppeForMail: ns.sanitizeGruppeForMail,
                INVALID_CHARS_REPLACE: ns.INVALID_CHARS_REPLACE,
                INVALID_CHARS_TEST: ns.INVALID_CHARS_TEST,
                teacherEmailMapping: {},
                stripSubjectTrailingDigits: true,
                normalizeNumberedSubjectFields: KS.normalizeNumberedSubjectFields
            }
        );

        expect(teams[0].teamName).toBe('SJ26 | 1AK | OMAI | 1');
        expect(teams[0].fach).toBe('OMAI');
        expect(teams[0].fachOriginal).toBe('OMAI1');
        expect(teams[0].gruppe).toBe('1');
        expect(KT).toBeTruthy();
    });
});
