import { describe, it, expect } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

function el(value) {
    return {
        value: value == null ? '' : String(value),
        style: { display: '' },
        textContent: '',
        innerHTML: '',
        focus() {},
        replaceChildren() {},
        appendChild() {},
        addEventListener() {}
    };
}

describe('kursteam-single', () => {
    it('baut Vorschau für ein Team (Klasse/Fach/Lehrer)', () => {
        const inputs = {
            singleTeamKlasse: el('4AK'),
            singleTeamFach: el('WINF'),
            singleTeamLehrer: el('MEI'),
            singleTeamGruppe: el(''),
            singleTeamOwner: el('mei@schule.at'),
            yearPrefix: el('SJ26'),
            teamSeparator: el(' | '),
            stripSubjectTrailingDigits: { checked: false, ...el('') }
        };

        const sandbox = {
            console,
            document: {
                readyState: 'complete',
                getElementById: (id) => inputs[id] || el(''),
                addEventListener: () => {}
            },
            window: null,
            location: { search: '' }
        };
        sandbox.window = sandbox;
        createContext(sandbox);

        const scripts = [
            'src/shared/ms365-module-guard.js',
            'src/tools/kursteams/kursteam-team-names.js',
            'src/tools/kursteams/kursteam-utils.js',
            'src/tools/kursteams/kursteam-team-build.js',
            'src/tools/kursteams/kursteam-single.js'
        ];
        scripts.forEach((rel) => {
            const full = join(root, rel);
            runInContext(readFileSync(full, 'utf8'), sandbox, { filename: full });
        });

        const ns = sandbox.ms365Kursteam;
        ns.teacherEmailMapping = { MEI: 'mei@schule.at' };
        ns.getPatternFromBuilder = null;

        const team = ns.buildSingleTeamPreviewEntry();
        expect(team).toBeTruthy();
        expect(team.teamName).toContain('4AK');
        expect(team.teamName).toContain('WINF');
        expect(team.gruppenmail.toLowerCase()).toContain('4ak');
        expect(team.gruppenmail.toLowerCase()).toContain('winf');
        expect(team.besitzer).toBe('mei@schule.at');
        expect(team.isValid).toBe(true);
    });

    it('bootSingleTeamModeFromQuery erkennt mode=single', () => {
        const sandbox = {
            console,
            URLSearchParams,
            document: {
                readyState: 'complete',
                getElementById: () => el(''),
                addEventListener: () => {}
            },
            window: null,
            location: { search: '?mode=single' }
        };
        sandbox.window = sandbox;
        createContext(sandbox);
        runInContext(
            readFileSync(join(root, 'src/tools/kursteams/kursteam-single.js'), 'utf8'),
            sandbox,
            { filename: 'kursteam-single.js' }
        );

        let started = false;
        sandbox.ms365Kursteam.startKursteamSingleTeam = () => {
            started = true;
        };
        expect(sandbox.ms365Kursteam.bootSingleTeamModeFromQuery()).toBe(true);
        expect(started).toBe(true);
    });
});
