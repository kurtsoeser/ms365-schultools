import { describe, it, expect } from 'vitest';
import {
    prefillStudentFreistellungForm,
    emptyForm,
    studentKlasseFromRecord,
    matchStudentByEmail,
    classesForStudentPicker,
    resolveStudentKlasseCode,
    resolveKvForClass
} from '../src/tools/freistellung-planer/freistellung-planer-state.js';

describe('prefillStudentFreistellungForm', () => {
    it('setzt Klasse und KV-E-Mail aus Stammdaten', () => {
        const state = {
            role: 'schueler',
            accountName: 'Max Mustermann',
            accountEmail: 'max@schule.at',
            studentMatch: { klasse: '3AH' },
            stammdaten: {
                classes: [
                    {
                        code: '3AH',
                        name: '3. AH',
                        headEmail: 'kv@schule.at',
                        headName: 'Frau Lehrer'
                    }
                ]
            },
            form: emptyForm()
        };
        prefillStudentFreistellungForm(state);
        expect(state.form.schuelerName).toBe('Max Mustermann');
        expect(state.form.klasse).toBe('3AH');
        expect(state.form.kvEmail).toBe('kv@schule.at');
        expect(state.form.kvName).toBe('Frau Lehrer');
    });

    it('resolveKvForClass findet KV auch über Schuljahr-Klassenliste', () => {
        const state = {
            stammdaten: { classes: [], teachers: [] },
            studentMatch: null
        };
        const orig = globalThis.ms365AppDataV2;
        globalThis.ms365AppDataV2 = {
            getContainer() {
                return {
                    core: { classTeams: [] },
                    years: {
                        byLabel: {
                            '2025/26': {
                                classes: [
                                    {
                                        code: '1A',
                                        name: '1A',
                                        headEmail: 'kv1a@schule.at',
                                        headName: 'KV 1A'
                                    }
                                ]
                            }
                        }
                    }
                };
            },
            getSetup() {
                return { classGroupMatchByKey: {} };
            },
            normalizeCoreClassTeams(arr) {
                return arr || [];
            }
        };
        const kv = resolveKvForClass(state, '1A');
        globalThis.ms365AppDataV2 = orig;
        expect(kv && kv.email).toBe('kv1a@schule.at');
        expect(kv && kv.name).toBe('KV 1A');
    });

    it('liest Klasse aus alternativen Stammdaten-Feldern', () => {
        expect(studentKlasseFromRecord({ classCode: '4BK' })).toBe('4BK');
        expect(
            resolveStudentKlasseCode({
                studentMatch: { mail: 'a@b.at', class: '2AH' }
            })
        ).toBe('2AH');
    });

    it('matchStudentByEmail über mail/upn', () => {
        const hit = matchStudentByEmail('max@schule.at', [{ name: 'Max', mail: 'max@schule.at', klasse: '1AK' }]);
        expect(hit && hit.klasse).toBe('1AK');
    });

    it('classesForStudentPicker ergänzt Klassen aus Anträgen', () => {
        const list = classesForStudentPicker({
            stammdaten: { classes: [{ code: '3AH', name: '3. AH' }] },
            items: [{ klasse: '4BK' }]
        });
        expect(list.map((c) => c.code)).toEqual(['3AH', '4BK']);
    });
});
