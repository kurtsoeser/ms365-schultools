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
        addEventListener() {},
        querySelector() {
            return null;
        },
        querySelectorAll() {
            return [];
        },
        getAttribute() {
            return null;
        }
    };
}

function makeRow(fields) {
    const inputs = {};
    Object.keys(fields).forEach((k) => {
        inputs[k] = el(fields[k]);
    });
    return {
        querySelector(sel) {
            const m = String(sel).match(/data-st-field="([^"]+)"/);
            if (m) return inputs[m[1]] || null;
            return null;
        },
        querySelectorAll(sel) {
            if (String(sel).includes('data-st-field')) {
                return Object.keys(inputs).map((k) => {
                    const inp = inputs[k];
                    inp.getAttribute = (name) => (name === 'data-st-field' ? k : null);
                    inp.addEventListener = () => {};
                    return inp;
                });
            }
            return [];
        }
    };
}

describe('kursteam-single', () => {
    it('baut Vorschau für mehrere Zeilen', () => {
        const row1 = makeRow({
            klasse: '4AK',
            fach: 'WINF',
            lehrer: 'MEI',
            gruppe: '',
            owner: 'mei@schule.at'
        });
        const row2 = makeRow({
            klasse: '3AK',
            fach: 'D',
            lehrer: 'ABC',
            gruppe: '',
            owner: 'abc@schule.at'
        });

        const sandbox = {
            console,
            document: {
                readyState: 'complete',
                getElementById: (id) => {
                    if (id === 'singleTeamRows') {
                        return {
                            querySelectorAll: (sel) =>
                                String(sel).includes('single-team-row') ? [row1, row2] : [],
                            querySelector: () => row1,
                            replaceChildren() {},
                            appendChild() {}
                        };
                    }
                    if (id === 'yearPrefix') return el('SJ26');
                    if (id === 'teamSeparator') return el(' | ');
                    if (id === 'stripSubjectTrailingDigits') return { checked: false };
                    if (id === 'singleTeamPreview') return el('');
                    return el('');
                },
                addEventListener: () => {},
                querySelector: () => null,
                createElement: () => el('')
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
        ns.teacherEmailMapping = { MEI: 'mei@schule.at', ABC: 'abc@schule.at' };
        ns.getPatternFromBuilder = null;

        const all = ns.buildAllSingleTeamEntries();
        expect(all).toHaveLength(2);
        expect(all[0].team.teamName).toContain('4AK');
        expect(all[0].team.teamName).toContain('WINF');
        expect(all[0].team.besitzer).toBe('mei@schule.at');
        expect(all[0].team.isValid).toBe(true);
        expect(all[1].team.teamName).toContain('3AK');
        expect(all[1].team.isValid).toBe(true);
    });

    it('bootSingleTeamModeFromQuery erkennt mode=single', () => {
        const sandbox = {
            console,
            URLSearchParams,
            document: {
                readyState: 'complete',
                getElementById: () => el(''),
                addEventListener: () => {},
                querySelector: () => null,
                createElement: () => el('')
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
