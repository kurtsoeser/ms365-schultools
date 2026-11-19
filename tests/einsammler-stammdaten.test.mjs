import { afterEach, describe, expect, it } from 'vitest';
import {
    buildEmpfaengerLinesForKlassen,
    loadEinsammlerStammdaten,
    parseKlasseCodesField
} from '../src/tools/einsammler/einsammler-stammdaten.js';

describe('einsammler-stammdaten', () => {
    afterEach(() => {
        delete globalThis.window;
        delete globalThis.ms365AppDataV2;
        delete globalThis.ms365TenantSettingsLoad;
        delete globalThis.ms365StammdatenCanonical;
    });

    it('parst Klassen-Codes', () => {
        expect(parseKlasseCodesField('3A; 3B,4C')).toEqual(['3A', '3B', '4C']);
    });

    it('baut Empfängerzeilen aus Schülern', () => {
        const students = [
            {
                name: 'Anna',
                klasse: '3A',
                parents: [{ email: 'eltern@example.com', name: 'E' }]
            },
            { name: 'Ben', klasse: '3A', parents: [] },
            { name: 'Cara', klasse: '2B', parents: [{ email: 'x@y.at', name: 'X' }] }
        ];
        const lines = buildEmpfaengerLinesForKlassen(['3A'], students);
        expect(lines).toContain('Anna;3A;eltern@example.com');
        expect(lines).toContain('Ben;3A;');
        expect(lines.some((l) => l.includes('Cara'))).toBe(false);
    });

    it('legt je Eltern-Mail eine Empfängerzeile an', () => {
        const students = [
            {
                name: 'Anna',
                klasse: '3A',
                parents: [
                    { email: 'a@example.com', name: 'A' },
                    { email: 'b@example.com', name: 'B' }
                ]
            }
        ];
        const lines = buildEmpfaengerLinesForKlassen(['3A'], students);
        expect(lines).toEqual(['Anna;3A;a@example.com', 'Anna;3A;b@example.com']);
    });

    it('lädt Schüler aus App-Daten v2 (auch ohne Tenant-Settings)', () => {
        globalThis.window = globalThis;
        globalThis.ms365AppDataV2 = {
            getContainer() {
                return {
                    years: {
                        current: '2026/27',
                        byLabel: {
                            '2026/27': {
                                classes: [{ code: '4B', name: 'DEMO Klasse 4B' }],
                                guardians: [{ id: 'g1', name: 'Eltern', email: 'eltern@schule.at' }],
                                students: [
                                    {
                                        name: 'Max Muster',
                                        klasse: '4B',
                                        email: 'max@schule.at',
                                        guardianIds: ['g1']
                                    }
                                ]
                            }
                        }
                    }
                };
            }
        };
        const bundle = loadEinsammlerStammdaten();
        expect(bundle.classes.some((c) => c.code === '4B')).toBe(true);
        expect(bundle.students).toHaveLength(1);
        expect(bundle.students[0].name).toBe('Max Muster');
        expect(bundle.students[0].parents[0].email).toBe('eltern@schule.at');
        const lines = buildEmpfaengerLinesForKlassen(['4B'], bundle.students);
        expect(lines).toEqual(['Max Muster;4B;eltern@schule.at']);
    });
});
