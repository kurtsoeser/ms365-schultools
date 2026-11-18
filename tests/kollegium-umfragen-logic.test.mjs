import { beforeEach, describe, expect, it } from 'vitest';
import {
    buildAuswertung,
    parseMails,
    teacherInAudience,
    umfrageNochOffen,
    validateAntwortPayload,
    validateUmfrage
} from '../src/tools/kollegium-umfragen/kollegium-umfragen-logic.js';
import {
    ANTWORTEN_COLUMNS,
    LIST_TITLES,
    UMFRAGEN_COLUMNS
} from '../src/tools/kollegium-umfragen/kollegium-umfragen-schema.js';
import { draftFromVorlage, VORLAGEN } from '../src/tools/kollegium-umfragen/kollegium-umfragen-templates.js';
import {
    DEMO_SEED_TAG,
    FRAGE_TYP,
    STATUS,
    ZIELGRUPPE,
    isDemoSeedItem
} from '../src/tools/kollegium-umfragen/kollegium-umfragen-model.js';
import { buildKollegiumDemoState } from '../src/tools/kollegium-umfragen/kollegium-umfragen-demo.js';
import {
    mapAntwortFromItem,
    mapAntwortToFields,
    mapUmfrageFromItem,
    mapUmfrageToFields
} from '../src/tools/kollegium-umfragen/kollegium-umfragen-graph.js';
import {
    getKollegiumSyncMeta,
    isKollegiumAccessDeniedError,
    pullKollegiumFromSharePoint,
    pushKollegiumAntwort,
    pushKollegiumUmfrage
} from '../src/tools/kollegium-umfragen/kollegium-umfragen-sync.js';
import { SCHOOL_TOOL_SITE_KINDS } from '../src/shared/school-tool-sites.js';

describe('kollegium-umfragen-templates', () => {
    it('liefert Vorlagen und Drafts', () => {
        expect(VORLAGEN.length).toBeGreaterThanOrEqual(5);
        const d = draftFromVorlage('ja-nein-beschluss');
        expect(d?.titel).toBeTruthy();
        expect(d?.fragen?.length).toBeGreaterThan(0);
    });
});

describe('kollegium-umfragen-logic', () => {
    it('validiert Umfragen', () => {
        expect(validateUmfrage({ titel: '', fragen: [] }).ok).toBe(false);
        const d = draftFromVorlage('ja-nein-beschluss');
        expect(validateUmfrage(d).ok).toBe(true);
    });

    it('prüft Zielgruppe manuell / alle / KV ohne Hardcode', () => {
        const u = {
            status: STATUS.AKTIV,
            zielgruppe: ZIELGRUPPE.MANUELL,
            empfaengerMails: 'a@schule.at; b@schule.at',
            fragen: [{ id: 'q1', typ: 'single', text: 'X', optionen: ['Ja', 'Nein'] }]
        };
        expect(teacherInAudience(u, { email: 'a@schule.at' })).toBe(true);
        expect(teacherInAudience(u, { email: 'c@schule.at' })).toBe(false);
        expect(
            teacherInAudience(
                { ...u, zielgruppe: ZIELGRUPPE.ALLE, empfaengerMails: '' },
                { email: 'x@schule.at' }
            )
        ).toBe(false);
        expect(
            teacherInAudience(
                { ...u, zielgruppe: ZIELGRUPPE.ALLE, empfaengerMails: '' },
                { email: 'x@schule.at', isTeacher: true }
            )
        ).toBe(true);
        const kvU = {
            status: STATUS.AKTIV,
            zielgruppe: ZIELGRUPPE.KV,
            empfaengerMails: '',
            fragen: [{ id: 'q1', typ: 'single', text: 'X', optionen: ['Ja', 'Nein'] }]
        };
        expect(teacherInAudience(kvU, { email: 'x@schule.at', isKv: false })).toBe(false);
        expect(teacherInAudience(kvU, { email: 'kv@schule.at', isKv: true })).toBe(true);
    });

    it('baut Auswertung für Single Choice', () => {
        const d = draftFromVorlage('ja-nein-beschluss');
        d.id = 'u1';
        d.status = STATUS.AKTIV;
        const qid = d.fragen[0].id;
        const aus = buildAuswertung(d, [
            {
                umfrageId: 'u1',
                status: 'abgegeben',
                lehrerEmail: 'a@x.at',
                lehrerName: 'A',
                werte: { [qid]: 'Ja' }
            },
            {
                umfrageId: 'u1',
                status: 'abgegeben',
                lehrerEmail: 'b@x.at',
                lehrerName: 'B',
                werte: { [qid]: 'Ja' }
            }
        ]);
        expect(aus.abgegeben).toBe(2);
        expect(aus.fragen[0].buckets[0].label).toBe('Ja');
        expect(aus.fragen[0].buckets[0].count).toBe(2);
    });

    it('parst Mails', () => {
        expect(parseMails('a@x.at; b@y.at')).toEqual(['a@x.at', 'b@y.at']);
    });

    it('respektiert Frist und Datei-Pflicht', () => {
        expect(
            umfrageNochOffen({ status: STATUS.AKTIV, faelligAm: '2020-01-01' }, '2026-11-06')
        ).toBe(false);
        expect(umfrageNochOffen({ status: STATUS.AKTIV, faelligAm: '' }, '2026-11-06')).toBe(true);
        const u = {
            id: 'u-datei',
            status: STATUS.AKTIV,
            faelligAm: '',
            fragen: [
                {
                    id: 'q1',
                    typ: FRAGE_TYP.DATEI,
                    text: 'Beleg',
                    pflicht: true,
                    optionen: []
                }
            ]
        };
        expect(
            validateAntwortPayload(u, {
                umfrageId: 'u-datei',
                lehrerEmail: 'a@x.at',
                werte: {}
            }).ok
        ).toBe(false);
        expect(
            validateAntwortPayload(u, {
                umfrageId: 'u-datei',
                lehrerEmail: 'a@x.at',
                werte: { q1: { name: 'beleg.pdf', key: 'x' } }
            }).ok
        ).toBe(true);
    });
});

describe('kollegium-umfragen-schema', () => {
    it('hat Listen und Spalten inkl. Anhang + SeedTag', () => {
        expect(LIST_TITLES.umfragen).toBe('KollegiumUmfragen');
        expect(UMFRAGEN_COLUMNS.some((c) => c.name === 'FragenJson')).toBe(true);
        expect(UMFRAGEN_COLUMNS.some((c) => c.name === 'AnhangErlaubt')).toBe(true);
        expect(UMFRAGEN_COLUMNS.some((c) => c.name === 'SeedTag')).toBe(true);
        expect(ANTWORTEN_COLUMNS.some((c) => c.name === 'AnhangName')).toBe(true);
        expect(ANTWORTEN_COLUMNS.some((c) => c.name === 'SeedTag')).toBe(true);
    });
});

describe('kollegium site kind', () => {
    it('kollegium ist in SCHOOL_TOOL_SITE_KINDS', () => {
        expect(SCHOOL_TOOL_SITE_KINDS.some((k) => k.id === 'kollegium')).toBe(true);
    });
});

describe('kollegium demo seed', () => {
    it('markiert Demo mit SeedTag / localDemoOnly', () => {
        const demo = buildKollegiumDemoState();
        expect(demo.localDemoOnly).toBe(true);
        expect(demo.umfragen.every(isDemoSeedItem)).toBe(true);
        expect(demo.antworten.every(isDemoSeedItem)).toBe(true);
        expect(DEMO_SEED_TAG).toBe('kollegium-demo-v1');
    });
});

describe('kollegium graph map', () => {
    it('mappt Umfrage Item ↔ Felder', () => {
        const mapped = mapUmfrageFromItem({
            id: '55',
            createdDateTime: '2026-11-18T10:00:00Z',
            fields: {
                UmfrageId: 'ku-1',
                Title: 'Beschluss',
                Beschreibung: 'Text',
                Zielgruppe: 'alle_lehrkraefte',
                FragenJson: JSON.stringify([{ id: 'q1', typ: 'single', text: 'Ja?', optionen: ['Ja', 'Nein'] }]),
                Status: 'aktiv',
                Anonym: true,
                AnhangErlaubt: false,
                ErstellerEmail: 'dir@schule.at',
                SeedTag: 'x'
            }
        });
        expect(mapped.id).toBe('ku-1');
        expect(mapped.spoItemId).toBe('55');
        expect(mapped.anonym).toBe(true);
        expect(mapped.fragen[0].text).toBe('Ja?');
        const fields = mapUmfrageToFields(mapped);
        expect(fields.Title).toBe('Beschluss');
        expect(fields.UmfrageId).toBe('ku-1');
        expect(fields.Anonym).toBe(true);
        expect(fields.SeedTag).toBe('x');
        const slim = mapUmfrageToFields(mapped, { omitOptional: true });
        expect(slim.SeedTag).toBeUndefined();
        expect(slim.Title).toBe('Beschluss');
    });

    it('mappt Antwort Item ↔ Felder', () => {
        const mapped = mapAntwortFromItem({
            id: '77',
            fields: {
                AntwortId: 'ka-1',
                UmfrageId: 'ku-1',
                LehrerEmail: 'a@schule.at',
                LehrerName: 'A',
                AntwortenJson: JSON.stringify({ q1: 'Ja' }),
                Status: 'abgegeben',
                SeedTag: 'x'
            }
        });
        expect(mapped.id).toBe('ka-1');
        expect(mapped.spoItemId).toBe('77');
        expect(mapped.werte.q1).toBe('Ja');
        const fields = mapAntwortToFields(mapped);
        expect(fields.AntwortId).toBe('ka-1');
        expect(fields.LehrerEmail).toBe('a@schule.at');
        expect(fields.SeedTag).toBe('x');
    });
});

describe('kollegium SPO sync', () => {
    beforeEach(() => {
        const mem = new Map();
        globalThis.localStorage = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => mem.set(k, String(v)),
            removeItem: (k) => mem.delete(k)
        };
        globalThis.window = globalThis;
        delete globalThis.window.ms365SpoGraph;
    });

    it('erkennt Access-Denied / 403', () => {
        expect(isKollegiumAccessDeniedError(new Error('403 Forbidden'))).toBe(true);
        expect(isKollegiumAccessDeniedError(new Error('accessDenied'))).toBe(true);
        expect(isKollegiumAccessDeniedError(new Error('Liste fehlt'))).toBe(false);
    });

    it('push überspringt Demo-Seeds', async () => {
        const demo = buildKollegiumDemoState();
        const rU = await pushKollegiumUmfrage(demo.umfragen[0]);
        expect(rU.ok).toBe(true);
        expect(rU.skipped).toBe(true);
        const rA = await pushKollegiumAntwort(demo.antworten[0]);
        expect(rA.ok).toBe(true);
        expect(rA.skipped).toBe(true);
    });

    it('403 beim Pull wipe’t lokalen Cache nicht', async () => {
        const store = await import('../src/tools/kollegium-umfragen/kollegium-umfragen-store.js');
        store.clearAll();
        store.upsertUmfrage({
            id: 'ku-local-keep',
            titel: 'Lokal behalten',
            zielgruppe: ZIELGRUPPE.MANUELL,
            empfaengerMails: 'me@schule.at',
            fragen: [{ id: 'q1', typ: 'single', text: 'X', optionen: ['Ja', 'Nein'] }],
            status: STATUS.ENTWURF
        });
        localStorage.setItem(
            'ms365-kollegium-umfragen-sp-settings-v1',
            JSON.stringify({
                siteUrl: 'https://contoso.sharepoint.com/sites/schule',
                listUmfragen: 'KollegiumUmfragen',
                listAntworten: 'KollegiumUmfrageAntworten'
            })
        );
        globalThis.window.ms365SpoGraph = {
            getGraphToken: async () => {
                throw new Error('403 Forbidden');
            },
            graphPathSite: () => '/sites/x',
            graphJson: async () => {
                throw new Error('should not run');
            },
            resolveSiteFromWebUrl: async () => ({ id: 'site', displayName: 'x' })
        };

        expect(store.loadState().umfragen).toHaveLength(1);
        const r = await pullKollegiumFromSharePoint({ forceDemoIfEmpty: false });
        expect(r.ok).toBe(false);
        expect(r.softError).toMatch(/Kein SharePoint-Zugriff/i);
        expect(store.loadState().umfragen).toHaveLength(1);
        expect(store.loadState().umfragen[0].id).toBe('ku-local-keep');
        expect(getKollegiumSyncMeta().source).toBe('local');
    });
});
