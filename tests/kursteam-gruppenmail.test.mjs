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
});
