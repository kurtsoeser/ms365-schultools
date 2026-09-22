import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import { loadScript } from './kursteams-vm.mjs';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

function loadScriptsInSandbox(relativePaths) {
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    relativePaths.forEach((rel) => {
        const full = join(root, rel);
        runInContext(readFileSync(full, 'utf8'), sandbox, { filename: full });
    });
    return sandbox;
}

describe('kursteam-gruppenmail', () => {
    it('formatKlasseSegmentForGruppenmail trennt jg-Jahrgang und Klasse', () => {
        const ctx = loadScript('src/tools/kursteams/kursteam-utils.js');
        const { formatKlasseSegmentForGruppenmail, buildGruppenmailBase } = ctx.ms365Kursteam;

        expect(formatKlasseSegmentForGruppenmail('jg20301hma')).toBe('jg2030-1hma');
        expect(formatKlasseSegmentForGruppenmail('jg20301a')).toBe('jg2030-1a');
        expect(formatKlasseSegmentForGruppenmail('1HMA')).toBe('1HMA');
        expect(buildGruppenmailBase('SJ26', 'jg20301hma', 'E', '')).toBe('SJ26-jg2030-1hma-E');
        expect(buildGruppenmailBase('SJ26', '1AK', 'D', '')).toBe('SJ26-1AK-D');
    });

    it('combineClassNames: 1AK,1BK → 1AK1BK (nicht 1AKB/hakb)', () => {
        const ctx = loadScript('src/tools/kursteams/kursteam-utils.js');
        const { combineClassNames, isCombinedClassCell, buildGruppenmailBase } = ctx.ms365Kursteam;

        expect(combineClassNames('1AK,1BK')).toBe('1AK1BK');
        expect(combineClassNames('1AK~1BK~1CK')).toBe('1AK1BK1CK');
        expect(combineClassNames('1AK,1BK', { mode: 'letters' })).toBe('AKBK');
        expect(isCombinedClassCell('1AK,1BK')).toBe(true);
        expect(isCombinedClassCell('1AK1BK')).toBe(true);
        expect(isCombinedClassCell('1AK')).toBe(false);
        expect(buildGruppenmailBase('SJ26', combineClassNames('1AK,1BK'), 'BESPM', '')).toBe(
            'SJ26-1AK1BK-BESPM'
        );
        expect(buildGruppenmailBase('SJ26', '1AK1BK', 'BESPM', '')).not.toMatch(/hakb/i);
    });

    it('combineClassNames smart: 1HMA,1HMB → 1HMAB', () => {
        const ctx = loadScript('src/tools/kursteams/kursteam-utils.js');
        const { combineClassNames } = ctx.ms365Kursteam;

        expect(combineClassNames('1HMA,1HMB', { mode: 'smart' })).toBe('1HMAB');
        expect(combineClassNames('1HMA~1HMB~1HMC', { mode: 'smart' })).toBe('1HMABC');
        expect(combineClassNames('1HMA,1HMB,2HMA,2HMB', { mode: 'smart' })).toBe('12HMAB');
        expect(combineClassNames('1AK,1BK', { mode: 'smart' })).toBe('1AKBK');
        // Kein sinnvolles Kürzen → Fallback auf concat
        expect(combineClassNames('1HMA,2AK', { mode: 'smart' })).toBe('1HMA2AK');
    });

    it('buildGruppenmailFromPattern: Trenner im Namen → Bindestrich in Gruppenmail', () => {
        const ctx = loadScriptsInSandbox([
            'src/tools/kursteams/kursteam-team-names.js',
            'src/tools/kursteams/kursteam-utils.js'
        ]);
        const { buildGruppenmailFromPattern, defaultTeamNamePattern } = ctx.ms365KursteamTeamNames;
        const { formatKlasseSegmentForGruppenmail } = ctx.ms365Kursteam;

        const pattern = defaultTeamNamePattern();
        expect(
            buildGruppenmailFromPattern(
                pattern,
                { yearPrefix: 'SJ26', klasse: 'jg20301hma', fach: 'E', gruppe: '' },
                { formatKlasse: formatKlasseSegmentForGruppenmail }
            )
        ).toBe('SJ26-jg2030-1hma-E');

        expect(
            buildGruppenmailFromPattern(
                pattern,
                { yearPrefix: 'SJ26', klasse: '1HMA', fach: 'E', gruppe: '' },
                { formatKlasse: formatKlasseSegmentForGruppenmail }
            )
        ).toBe('SJ26-1HMA-E');

        const customSep = [
            { type: 'yearPrefix' },
            { type: 'text', value: ' _ ' },
            { type: 'klasse' },
            { type: 'text', value: ' :: ' },
            { type: 'fach' }
        ];
        expect(
            buildGruppenmailFromPattern(
                customSep,
                { yearPrefix: 'SJ26', klasse: '2AK', fach: 'M', gruppe: '' },
                { formatKlasse: formatKlasseSegmentForGruppenmail }
            )
        ).toBe('SJ26-2AK-M');
    });

    it('buildTeamEntriesFromRows nutzt dasselbe Muster wie der Team-Name', () => {
        const ctx = loadScriptsInSandbox([
            'src/shared/ms365-module-guard.js',
            'src/tools/kursteams/kursteam-team-names.js',
            'src/tools/kursteams/kursteam-utils.js',
            'src/tools/kursteams/kursteam-team-build.js'
        ]);
        const KTB = ctx.ms365KursteamTeamBuild;
        const KT = ctx.ms365KursteamTeamNames;
        const ns = ctx.ms365Kursteam;

        const teams = KTB.buildTeamEntriesFromRows(
            [{ klasse: '1HMA', fach: 'E', lehrer: 'ABC', gruppe: '' }],
            {
                yearPrefix: 'SJ26',
                emailDomain: '@schule.at',
                separator: ' | ',
                pattern: KT.defaultTeamNamePattern(),
                combineClassNames: ns.combineClassNames,
                isCombinedClassCell: ns.isCombinedClassCell,
                buildGruppenmailBase: ns.buildGruppenmailBase,
                formatKlasseSegmentForGruppenmail: ns.formatKlasseSegmentForGruppenmail,
                sanitizeGruppeForMail: ns.sanitizeGruppeForMail,
                INVALID_CHARS_REPLACE: ns.INVALID_CHARS_REPLACE,
                INVALID_CHARS_TEST: ns.INVALID_CHARS_TEST,
                teacherEmailMapping: { ABC: 'lehrer@schule.at' }
            }
        );

        expect(teams[0].teamName).toBe('SJ26 | 1HMA | E');
        expect(teams[0].gruppenmail).toBe('SJ26-1HMA-E');
    });

    it('buildTeamEntriesFromRows: Mehrklassen → 1AK1BK in Name/Mail, kein Einzelklassen-Nick', () => {
        const ctx = loadScriptsInSandbox([
            'src/shared/ms365-module-guard.js',
            'src/tools/kursteams/kursteam-team-names.js',
            'src/tools/kursteams/kursteam-utils.js',
            'src/tools/kursteams/kursteam-team-build.js'
        ]);
        const KTB = ctx.ms365KursteamTeamBuild;
        const KT = ctx.ms365KursteamTeamNames;
        const ns = ctx.ms365Kursteam;

        ctx.ms365AppDataV2 = {
            getClassTeamGruppenmailForKlasse: () => 'jg2030hak'
        };

        const teams = KTB.buildTeamEntriesFromRows(
            [{ klasse: '1AK,1BK', fach: 'BESPM', lehrer: 'ABC', gruppe: 'BESPM1AK1BK' }],
            {
                yearPrefix: 'SJ26',
                emailDomain: '@schule.at',
                separator: ' | ',
                pattern: KT.defaultTeamNamePattern(),
                combineClassNames: ns.combineClassNames,
                isCombinedClassCell: ns.isCombinedClassCell,
                buildGruppenmailBase: ns.buildGruppenmailBase,
                formatKlasseSegmentForGruppenmail: ns.formatKlasseSegmentForGruppenmail,
                sanitizeGruppeForMail: ns.sanitizeGruppeForMail,
                INVALID_CHARS_REPLACE: ns.INVALID_CHARS_REPLACE,
                INVALID_CHARS_TEST: ns.INVALID_CHARS_TEST,
                teacherEmailMapping: { ABC: 'lehrer@schule.at' }
            }
        );

        expect(teams[0].teamName).toContain('1AK1BK');
        expect(teams[0].gruppenmail).toMatch(/1AK1BK/i);
        expect(teams[0].gruppenmail).not.toMatch(/hakb/i);
        expect(teams[0].gruppenmail).not.toMatch(/jg2030hak/i);
        expect(teams[0].originalClass).toBe('1AK,1BK');
    });

    it('buildTeamEntriesFromRows: classCombineMode smart → 1HMAB', () => {
        const ctx = loadScriptsInSandbox([
            'src/shared/ms365-module-guard.js',
            'src/tools/kursteams/kursteam-team-names.js',
            'src/tools/kursteams/kursteam-utils.js',
            'src/tools/kursteams/kursteam-team-build.js'
        ]);
        const KTB = ctx.ms365KursteamTeamBuild;
        const ns = ctx.ms365Kursteam;

        const teams = KTB.buildTeamEntriesFromRows(
            [{ klasse: '1HMA,1HMB', fach: 'PH', lehrer: 'LOIE', gruppe: '' }],
            {
                yearPrefix: 'SJ26-27',
                emailDomain: '@schule.at',
                separator: ' | ',
                pattern: [
                    { type: 'yearPrefix' },
                    { type: 'text', value: ' | ' },
                    { type: 'klasse' },
                    { type: 'text', value: ' | ' },
                    { type: 'fach' },
                    { type: 'text', value: ' | ' },
                    { type: 'lehrer' }
                ],
                combineClassNames: ns.combineClassNames,
                isCombinedClassCell: ns.isCombinedClassCell,
                buildGruppenmailBase: ns.buildGruppenmailBase,
                formatKlasseSegmentForGruppenmail: ns.formatKlasseSegmentForGruppenmail,
                sanitizeGruppeForMail: ns.sanitizeGruppeForMail,
                INVALID_CHARS_REPLACE: ns.INVALID_CHARS_REPLACE,
                INVALID_CHARS_TEST: ns.INVALID_CHARS_TEST,
                teacherEmailMapping: { LOIE: 'lehrer@schule.at' },
                classCombineMode: 'smart'
            }
        );

        expect(teams[0].teamName).toBe('SJ26-27 | 1HMAB | PH | LOIE');
        expect(teams[0].gruppenmail).toMatch(/1HMAB/i);
        expect(teams[0].originalClass).toBe('1HMA,1HMB');
    });
});
