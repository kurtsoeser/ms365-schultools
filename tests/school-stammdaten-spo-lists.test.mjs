import { describe, it, expect } from 'vitest';
import {
    defaultStammdatenSpoListTitles,
    LEHRER_LIST_COLUMNS,
    KLASSEN_LIST_COLUMNS,
    stammdatenColumnToGraphBody,
    mapLehrerFieldsToTeacher,
    mapKlassenFieldsToClass,
    mergeStammdatenPreferLive,
    resolveStammdatenSpoListTitles,
    fetchSchoolStammdatenFromSite,
    healthSchoolStammdatenShellLists,
    teacherToLehrerListFields,
    classToKlassenListFields,
    loadLocalSchoolStammdatenFromWindow,
    siteWebUrlFromSharePointListUrl
} from '../src/shared/school-stammdaten-spo-lists.js';

describe('school-stammdaten-spo-lists', () => {
    it('liefert Standard-Titel Lehrer/Klassen', () => {
        const t = defaultStammdatenSpoListTitles();
        expect(t.lehrer).toBeTruthy();
        expect(t.klassen).toBe('Klassen');
        expect(resolveStammdatenSpoListTitles({ lehrer: 'Kollegium' }).lehrer).toBe('Kollegium');
    });

    it('siteWebUrlFromSharePointListUrl schneidet /Lists/… ab', () => {
        expect(
            siteWebUrlFromSharePointListUrl(
                'https://kurtrocks.sharepoint.com/sites/MS365-Schultools/Lists/Lehrerinnen'
            )
        ).toBe('https://kurtrocks.sharepoint.com/sites/MS365-Schultools');
        expect(
            siteWebUrlFromSharePointListUrl('https://kurtrocks.sharepoint.com/sites/MS365-Schultools')
        ).toBe('https://kurtrocks.sharepoint.com/sites/MS365-Schultools');
    });

    it('Lehrer- und Klassen-Spalten enthalten KV / Kürzel / M365-Gruppe', () => {
        expect(LEHRER_LIST_COLUMNS.some((c) => c.name === 'LehrerCode')).toBe(true);
        expect(KLASSEN_LIST_COLUMNS.some((c) => c.name === 'KVEmail')).toBe(true);
        expect(KLASSEN_LIST_COLUMNS.some((c) => c.name === 'Klassenvorstand')).toBe(true);
        expect(KLASSEN_LIST_COLUMNS.some((c) => c.name === 'M365GroupId')).toBe(true);
        expect(KLASSEN_LIST_COLUMNS.some((c) => c.name === 'M365GroupMail')).toBe(true);
        const gruppe = KLASSEN_LIST_COLUMNS.find((c) => c.name === 'M365Gruppe');
        expect(gruppe && gruppe.personOrGroup && gruppe.personOrGroup.chooseFromType).toBe(
            'peopleAndGroups'
        );
    });

    it('Graph-Body für Text und Person', () => {
        const text = stammdatenColumnToGraphBody(LEHRER_LIST_COLUMNS[0]);
        expect(text.text).toBeTruthy();
        const person = stammdatenColumnToGraphBody(
            LEHRER_LIST_COLUMNS.find((c) => c.name === 'Lehrkraft')
        );
        expect(person.personOrGroup).toBeTruthy();
    });

    it('mapLehrerFieldsToTeacher bevorzugt Kürzel und Display Name', () => {
        const t = mapLehrerFieldsToTeacher({
            LehrerCode: 'BAU',
            Vorname: 'Julia',
            Nachname: 'Bauer',
            EMail: 'julia.bauer@schule.at',
            Title: ''
        });
        expect(t).toEqual({
            code: 'BAU',
            name: 'Julia Bauer',
            email: 'julia.bauer@schule.at',
            upn: 'julia.bauer@schule.at'
        });
    });

    it('mapLehrerFieldsToTeacher liest Personenfeld Lehrkraft und Title-only', () => {
        const fromPerson = mapLehrerFieldsToTeacher({
            Title: '',
            Lehrkraft: { Email: 'max@schule.at', LookupValue: 'Max Muster' }
        });
        expect(fromPerson && fromPerson.email).toBe('max@schule.at');
        expect(fromPerson && fromPerson.name).toBe('Max Muster');
        expect(fromPerson && fromPerson.code).toBe('max@schule.at');

        const titleOnly = mapLehrerFieldsToTeacher({ Title: 'Andrea Hagauer' });
        expect(titleOnly && titleOnly.code).toBe('Andrea Hagauer');
        expect(titleOnly && titleOnly.name).toBe('Andrea Hagauer');
    });

    it('mapKlassenFieldsToClass mappt KV, Abschlussjahr und M365-Gruppe', () => {
        const c = mapKlassenFieldsToClass({
            Title: '2AK',
            KlassenCode: '2AK',
            Abschlussjahr: '2027',
            KVName: 'Dave Grohl',
            KVEmail: 'dave.grohl@schule.at',
            M365GroupId: 'gid-123',
            M365GroupMail: '2ak@schule.at',
            M365GroupName: 'Klasse 2AK',
            Schuelerinnen: [
                { Email: 'a@schule.at', LookupValue: 'Anna' },
                { Email: 'b@schule.at', LookupValue: 'Ben' }
            ]
        });
        expect(c.code).toBe('2AK');
        expect(c.year).toBe('2027');
        expect(c.headEmail).toBe('dave.grohl@schule.at');
        expect(c.kvEmail).toBe('dave.grohl@schule.at');
        expect(c.m365GroupId).toBe('gid-123');
        expect(c.m365GroupMail).toBe('2ak@schule.at');
        expect(c.m365GroupName).toBe('Klasse 2AK');
        expect(c.studentEmails).toEqual(['a@schule.at', 'b@schule.at']);
    });

    it('classToKlassenListFields schreibt M365-Felder', () => {
        const fields = classToKlassenListFields({
            code: '1A',
            name: '1A',
            year: '2028',
            m365GroupId: 'abc',
            m365GroupMail: '1a@schule.at',
            m365GroupName: 'Klasse 1A'
        });
        expect(fields.M365GroupId).toBe('abc');
        expect(fields.M365GroupMail).toBe('1a@schule.at');
        expect(fields.M365GroupName).toBe('Klasse 1A');
    });

    it('mergeStammdatenPreferLive nimmt Live wenn vorhanden', () => {
        const merged = mergeStammdatenPreferLive(
            { teachers: [{ code: 'OLD' }], classes: [], students: [{ id: 1 }] },
            { teachers: [{ code: 'NEW', name: 'Neu', email: 'n@x.at' }], classes: [{ code: '1A' }] }
        );
        expect(merged.teachers[0].code).toBe('NEW');
        expect(merged.classes[0].code).toBe('1A');
        expect(merged.students).toHaveLength(1);
    });

    it('fetchSchoolStammdatenFromSite liest und mappt Listenzeilen', async () => {
        const lists = {
            Lehrerinnen: {
                id: 'L1',
                displayName: 'Lehrerinnen',
                webUrl: 'https://x/lists/lehrer',
                items: [
                    {
                        fields: {
                            LehrerCode: 'MUE',
                            Title: 'Mag. Müller',
                            EMail: 'mueller@schule.at'
                        }
                    }
                ]
            },
            Klassen: {
                id: 'K1',
                displayName: 'Klassen',
                webUrl: 'https://x/lists/klassen',
                items: [
                    {
                        fields: {
                            KlassenCode: '5AK',
                            Title: '5AK',
                            Abschlussjahr: '2027',
                            KVEmail: 'kv@schule.at',
                            KVName: 'KV'
                        }
                    }
                ]
            }
        };
        const graph = {
            graphPathSite: () => '/sites/s',
            graphJson: async (_method, path) => {
                if (String(path).includes('/lists?$filter=')) {
                    if (String(path).includes('Lehrerinnen')) {
                        return { value: [{ id: 'L1', displayName: 'Lehrerinnen', webUrl: lists.Lehrerinnen.webUrl }] };
                    }
                    if (String(path).includes('Klassen')) {
                        return { value: [{ id: 'K1', displayName: 'Klassen', webUrl: lists.Klassen.webUrl }] };
                    }
                    return { value: [] };
                }
                if (String(path).includes('/lists/L1/items')) {
                    return { value: lists.Lehrerinnen.items };
                }
                if (String(path).includes('/lists/K1/items')) {
                    return { value: lists.Klassen.items };
                }
                return { value: [] };
            }
        };
        const result = await fetchSchoolStammdatenFromSite({
            token: 't',
            siteId: 's',
            graph
        });
        expect(result.missing).toEqual([]);
        expect(result.teachers).toHaveLength(1);
        expect(result.teachers[0].code).toBe('MUE');
        expect(result.classes).toHaveLength(1);
        expect(result.classes[0].year).toBe('2027');
        expect(result.classes[0].headEmail).toBe('kv@schule.at');
    });

    it('teacherToLehrerListFields setzt Kürzel und E-Mail', () => {
        const f = teacherToLehrerListFields({
            code: 'HAG',
            name: 'Andrea Beispiel-Riml',
            email: 'andrea.beispiel-riml@schule.example'
        });
        expect(f.LehrerCode).toBe('HAG');
        expect(f.Title).toBe('Andrea Beispiel-Riml');
        expect(f.EMail).toBe('andrea.beispiel-riml@schule.example');
        expect(f.Vorname).toBe('Andrea');
        expect(f.Nachname).toBe('Beispiel-Riml');
    });

    it('classToKlassenListFields setzt KV und Abschlussjahr', () => {
        const f = classToKlassenListFields({
            code: '1AK',
            name: '1AK',
            year: '2031',
            headName: 'Andrea Beispiel-Riml',
            headEmail: 'andrea.beispiel-riml@schule.example'
        });
        expect(f.KlassenCode).toBe('1AK');
        expect(f.Abschlussjahr).toBe('2031');
        expect(f.KVName).toBe('Andrea Beispiel-Riml');
        expect(f.KVEmail).toBe('andrea.beispiel-riml@schule.example');
    });

    it('loadLocalSchoolStammdatenFromWindow liest Tenant-Settings', () => {
        const win = {
            ms365TenantSettingsLoad: () => ({
                teachers: [{ code: 'X' }],
                classes: [{ code: '1A' }]
            })
        };
        const local = loadLocalSchoolStammdatenFromWindow(win);
        expect(local.teachers).toHaveLength(1);
        expect(local.classes[0].code).toBe('1A');
        expect(local.source).toContain('tenant-settings');
    });

    it('loadLocalSchoolStammdatenFromWindow liest app-data-v2 getContainer (nicht getCore)', () => {
        const win = {
            ms365AppDataV2: {
                getContainer: () => ({
                    core: {
                        teachers: [
                            { code: 'YOU', name: 'Angus Young', email: 'angus.young@kurtrocks.com' },
                            { code: 'GRO', name: 'Dave Grohl', email: 'dave.grohl@kurtrocks.com' }
                        ]
                    },
                    years: {
                        current: '2025/26',
                        byLabel: {
                            '2025/26': {
                                classes: [
                                    { code: '1A', name: 'DEMO Klasse 1A' },
                                    { code: '4B', name: 'DEMO Klasse 4B' }
                                ]
                            }
                        }
                    }
                })
            }
        };
        const local = loadLocalSchoolStammdatenFromWindow(win);
        expect(local.teachers).toHaveLength(2);
        expect(local.classes).toHaveLength(2);
        expect(local.source).toContain('app-data-v2');
    });

    it('loadLocalSchoolStammdatenFromWindow liest ms365-schooltool-data-v2 aus localStorage', () => {
        const store = {
            'ms365-schooltool-data-v2': JSON.stringify({
                core: { teachers: [{ code: 'SCO', name: 'Bon Scott' }] },
                years: {
                    current: '2025/26',
                    byLabel: { '2025/26': { classes: [{ code: '2A', name: 'Klasse 2A' }] } }
                }
            })
        };
        const win = {
            localStorage: {
                getItem: (k) => (Object.prototype.hasOwnProperty.call(store, k) ? store[k] : null)
            }
        };
        const local = loadLocalSchoolStammdatenFromWindow(win);
        expect(local.teachers[0].code).toBe('SCO');
        expect(local.classes[0].code).toBe('2A');
        expect(local.source).toContain('localStorage-v2');
    });

    it('fetchSchoolStammdatenFromSite nutzt SPO-REST wenn Graph 0 Zeilen liefert', async () => {
        const graph = {
            graphPathSite: () => '/sites/s',
            getGraphToken: async () => 'spo-tok',
            getSpoRequestDigest: async () => 'digest',
            spoRestFetch: async (_url, _tok, _dig, _method, api) => {
                if (String(api).includes("getbytitle('Klassen')")) {
                    return {
                        ok: true,
                        status: 200,
                        text: '',
                        data: {
                            value: [
                                {
                                    Id: 1,
                                    Title: 'DEMO Klasse 4B',
                                    KlassenCode: '4B',
                                    Abschlussjahr: '2029',
                                    KVName: 'Lehrer01 DEMO',
                                    KVEmail: 'lehrer01.demo@schule.at'
                                }
                            ]
                        }
                    };
                }
                if (String(api).includes("getbytitle('Lehrerinnen')")) {
                    return {
                        ok: true,
                        status: 200,
                        text: '',
                        data: {
                            value: [
                                {
                                    Id: 1,
                                    Title: 'Lehrer01 DEMO',
                                    LehrerCode: 'L01',
                                    EMail: 'lehrer01.demo@schule.at'
                                }
                            ]
                        }
                    };
                }
                return { ok: false, status: 404, text: 'missing', data: null };
            },
            graphJson: async (_method, path) => {
                if (String(path).includes('/lists?$filter=')) {
                    if (String(path).includes('Klassen')) {
                        return { value: [{ id: 'K1', displayName: 'Klassen', webUrl: 'https://x/lists/klassen' }] };
                    }
                    if (String(path).includes('Lehrerinnen')) {
                        return {
                            value: [{ id: 'L1', displayName: 'Lehrerinnen', webUrl: 'https://x/lists/lehrer' }]
                        };
                    }
                    return { value: [] };
                }
                if (String(path).includes('/items')) return { value: [] };
                return { value: [] };
            }
        };
        const result = await fetchSchoolStammdatenFromSite({
            token: 't',
            siteId: 's',
            siteWebUrl: 'https://contoso.sharepoint.com/sites/MS365-Schultools',
            graph
        });
        expect(result.source).toBe('spo-rest');
        expect(result.classes).toHaveLength(1);
        expect(result.classes[0].code).toBe('4B');
        expect(result.teachers).toHaveLength(1);
        expect(result.teachers[0].email).toBe('lehrer01.demo@schule.at');
        expect(result.readErrors).toEqual([]);
    });

    it('Health meldet leere Klassenliste als nicht ok', async () => {
        const graph = {
            graphPathSite: () => '/sites/s',
            graphJson: async (_m, path) => {
                const p = String(path);
                if (p.includes('/lists?') && p.includes('Klassen')) {
                    return { value: [{ id: 'K1', displayName: 'Klassen' }] };
                }
                if (p.includes('/lists?') && p.includes('Lehrer')) {
                    return { value: [{ id: 'L1', displayName: 'Lehrerinnen' }] };
                }
                if (p.includes('/columns')) {
                    return {
                        value: [
                            { name: 'Vorname' },
                            { name: 'Nachname' },
                            { name: 'LehrerCode' },
                            { name: 'EMail' },
                            { name: 'UPN' },
                            { name: 'Lehrkraft' },
                            { name: 'KlassenCode' },
                            { name: 'Abschlussjahr' },
                            { name: 'KVName' },
                            { name: 'KVEmail' },
                            { name: 'Klassenvorstand' },
                            { name: 'M365GroupId' },
                            { name: 'M365GroupMail' },
                            { name: 'M365GroupName' },
                            { name: 'M365Gruppe' }
                        ]
                    };
                }
                if (p.includes('/items')) return { value: [] };
                return { value: [] };
            }
        };
        const report = await healthSchoolStammdatenShellLists({
            token: 't',
            siteId: 's',
            graph
        });
        const klassen = report.find((r) => r.key === 'klassen');
        expect(klassen).toBeTruthy();
        expect(klassen.ok).toBe(false);
        expect(klassen.itemCount).toBe(0);
        expect(String(klassen.detail)).toMatch(/0 Einträge/);
    });
});
