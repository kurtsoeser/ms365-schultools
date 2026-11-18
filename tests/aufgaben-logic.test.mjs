import { describe, expect, it, beforeEach } from 'vitest';
import {
    canManageAufgabe,
    filterAufgaben,
    aufgabenStats,
    sortAufgaben,
    statusPipelineActions,
    validateAufgabe
} from '../src/tools/aufgaben/aufgaben-logic.js';
import {
    DEMO_SEED_TAG,
    PRIORITAET,
    STATUS,
    isDemoSeedItem
} from '../src/tools/aufgaben/aufgaben-model.js';
import { buildAufgabenDemoState } from '../src/tools/aufgaben/aufgaben-demo.js';
import {
    AUFGABEN_VORLAGEN,
    draftFromAufgabeVorlage,
    getAufgabeVorlage
} from '../src/tools/aufgaben/aufgaben-templates.js';
import {
    buildMailtoHref,
    buildStatusMailto,
    buildZuweisungMailto
} from '../src/tools/aufgaben/aufgaben-notify.js';
import { mapAufgabeFromItem, mapAufgabeToFields } from '../src/tools/aufgaben/aufgaben-graph.js';
import {
    getAufgabenSyncMeta,
    isAufgabenAccessDeniedError,
    pullAufgabenFromSharePoint,
    pushAufgabe
} from '../src/tools/aufgaben/aufgaben-sync.js';
import { resolveSchoolToolSiteUrl, SCHOOL_TOOL_SITE_KINDS } from '../src/shared/school-tool-sites.js';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { dirname, join } from 'node:path';

describe('aufgaben-logic', () => {
    it('validateAufgabe verlangt Titel und Zuweisung', () => {
        expect(validateAufgabe({ title: '', zuweisungName: '' }).ok).toBe(false);
        expect(
            validateAufgabe({
                title: 'Unterschrift',
                zuweisungName: 'Müller',
                zuweisungEmail: 'mueller@schule.at',
                prioritaet: 'mittel'
            }).ok
        ).toBe(true);
    });

    it('statusPipelineActions für offen und erledigt', () => {
        const offen = statusPipelineActions(STATUS.OFFEN).map((a) => a.to);
        expect(offen).toContain(STATUS.IN_ARBEIT);
        expect(offen).toContain(STATUS.ERLEDIGT);
        expect(statusPipelineActions(STATUS.ERLEDIGT).map((a) => a.to)).toEqual([STATUS.OFFEN]);
    });

    it('canManageAufgabe für Assignee/Creator/Admin', () => {
        const a = {
            id: '1',
            zuweisungEmail: 'me@schule.at',
            erstelltVonEmail: 'other@schule.at',
            status: STATUS.OFFEN
        };
        expect(canManageAufgabe(a, { email: 'me@schule.at' })).toBe(true);
        expect(canManageAufgabe(a, { email: 'other@schule.at' })).toBe(true);
        expect(canManageAufgabe(a, { email: 'x@schule.at' })).toBe(false);
        expect(canManageAufgabe(a, { email: 'x@schule.at', admin: true })).toBe(true);
    });

    it('sortiert nach Priorität dann Frist', () => {
        const sorted = sortAufgaben([
            {
                id: 'a',
                title: 'A',
                prioritaet: PRIORITAET.MITTEL,
                frist: '2026-01-02'
            },
            {
                id: 'b',
                title: 'B',
                prioritaet: PRIORITAET.DRINGEND,
                frist: '2026-01-10'
            }
        ]);
        expect(sorted.map((x) => x.id)).toEqual(['b', 'a']);
    });

    it('aufgabenStats zählt meine offenen und überfällig', () => {
        const st = aufgabenStats(
            [
                {
                    status: STATUS.OFFEN,
                    prioritaet: PRIORITAET.DRINGEND,
                    zuweisungEmail: 'me@schule.at',
                    frist: '2000-01-01'
                },
                {
                    status: STATUS.IN_ARBEIT,
                    prioritaet: PRIORITAET.MITTEL,
                    zuweisungEmail: 'other@schule.at',
                    frist: '2099-01-01'
                },
                { status: STATUS.ERLEDIGT, prioritaet: PRIORITAET.HOCH, zuweisungEmail: 'me@schule.at' }
            ],
            'me@schule.at'
        );
        expect(st.offen).toBe(2);
        expect(st.meineOffen).toBe(1);
        expect(st.ueberfaellig).toBe(1);
        expect(st.dringend).toBe(1);
        expect(st.erledigt).toBe(1);
    });
});

describe('aufgaben-demo / store actor', () => {
    beforeEach(() => {
        const mem = new Map();
        globalThis.localStorage = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => mem.set(k, String(v)),
            removeItem: (k) => mem.delete(k)
        };
    });

    it('Demo enthält seedTag und localDemoOnly', () => {
        const demo = buildAufgabenDemoState();
        expect(demo.localDemoOnly).toBe(true);
        expect(demo.aufgaben.every(isDemoSeedItem)).toBe(true);
        expect(DEMO_SEED_TAG).toBe('aufgaben-demo-v1');
    });

    it('setAufgabeStatus speichert Actor', async () => {
        const store = await import('../src/tools/aufgaben/aufgaben-store.js');
        store.clearAll();
        const created = store.upsertAufgabe(
            {
                id: 'af-actor-1',
                title: 'Test',
                zuweisungName: 'Me',
                zuweisungEmail: 'me@schule.at',
                prioritaet: 'mittel'
            },
            { actorName: 'Melder', actorEmail: 'melder@schule.at' }
        );
        expect(created.ok).toBe(true);
        expect(created.item.erstelltVonEmail).toBe('melder@schule.at');
        const r = store.setAufgabeStatus('af-actor-1', STATUS.IN_ARBEIT, {
            actorName: 'Admin',
            actorEmail: 'admin@schule.at'
        });
        expect(r.ok).toBe(true);
        expect(r.item.bearbeitetVonEmail).toBe('admin@schule.at');
        expect(r.item.bearbeitetVonName).toBe('Admin');
        expect(r.item.status).toBe(STATUS.IN_ARBEIT);
    });

    it('clearDemoSeed entfernt nur Seed-Zeilen', async () => {
        const store = await import('../src/tools/aufgaben/aufgaben-store.js');
        store.loadDemoData();
        store.upsertAufgabe({
            id: 'af-real-1',
            title: 'Echt',
            zuweisungName: 'Real',
            zuweisungEmail: 'real@schule.at',
            prioritaet: 'niedrig'
        });
        const after = store.clearDemoSeed();
        expect(after.aufgaben).toHaveLength(1);
        expect(after.aufgaben[0].id).toBe('af-real-1');
    });
});

describe('aufgaben site kind', () => {
    it('aufgaben ist in SCHOOL_TOOL_SITE_KINDS', () => {
        expect(SCHOOL_TOOL_SITE_KINDS.some((k) => k.id === 'aufgaben')).toBe(true);
        expect(typeof resolveSchoolToolSiteUrl).toBe('function');
    });
});

describe('aufgaben vorlagen & notify', () => {
    it('liefert Roadmap-Vorlagen', () => {
        expect(AUFGABEN_VORLAGEN.length).toBeGreaterThanOrEqual(4);
        expect(getAufgabeVorlage('unterschrift')?.titel).toMatch(/Unterschrift/i);
        const d = draftFromAufgabeVorlage('elternkontakt');
        expect(d?.title).toMatch(/Elternkontakt/i);
        expect(d?.prioritaet).toBe('dringend');
    });

    it('baut mailto für Zuweisung und Erledigt', () => {
        const href = buildZuweisungMailto(
            {
                title: 'Test',
                zuweisungName: 'Müller',
                zuweisungEmail: 'mueller@schule.at',
                erstelltVon: 'Admin',
                prioritaet: 'hoch',
                frist: '2026-12-01'
            },
            { fromName: 'Admin' }
        );
        expect(href).toMatch(/^mailto:mueller@schule\.at\?/);
        expect(href).toMatch(/Aufgabe%20zugewiesen/);
        expect(buildMailtoHref('', 'x', 'y')).toBe('');
        const done = buildStatusMailto({
            title: 'Test',
            status: 'erledigt',
            erstelltVonEmail: 'admin@schule.at',
            erstelltVon: 'Admin',
            zuweisungName: 'Müller'
        });
        expect(done).toMatch(/^mailto:admin@schule\.at\?/);
        expect(buildStatusMailto({ status: 'offen', erstelltVonEmail: 'a@b.at' })).toBe('');
    });
});

describe('aufgaben graph map', () => {
    it('mappt Item ↔ Felder', () => {
        const mapped = mapAufgabeFromItem({
            id: '99',
            fields: {
                AufgabeId: 'af-1',
                Title: 'Hallo',
                ZuweisungEmail: 'a@schule.at',
                Status: 'offen',
                Prioritaet: 'mittel',
                Frist: '2026-12-01T00:00:00Z',
                BearbeitetVonName: 'Admin',
                SeedTag: 'x'
            }
        });
        expect(mapped.id).toBe('af-1');
        expect(mapped.spoItemId).toBe('99');
        expect(mapped.frist).toBe('2026-12-01');
        const fields = mapAufgabeToFields(mapped);
        expect(fields.Title).toBe('Hallo');
        expect(fields.AufgabeId).toBe('af-1');
        expect(fields.ZuweisungEmail).toBe('a@schule.at');
        expect(fields.Frist).toBe('2026-12-01');
        expect(fields.BearbeitetVonName).toBe('Admin');
        expect(fields.SeedTag).toBe('x');
        const slim = mapAufgabeToFields(mapped, { omitOptional: true });
        expect(slim.BearbeitetVonName).toBeUndefined();
        expect(slim.SeedTag).toBeUndefined();
        expect(slim.Title).toBe('Hallo');
    });
});

describe('aufgaben SPO sync', () => {
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
        expect(isAufgabenAccessDeniedError(new Error('403 Forbidden'))).toBe(true);
        expect(isAufgabenAccessDeniedError(new Error('accessDenied'))).toBe(true);
        expect(isAufgabenAccessDeniedError(new Error('Liste fehlt'))).toBe(false);
    });

    it('push überspringt Demo-Seeds', async () => {
        const demo = buildAufgabenDemoState().aufgaben[0];
        const r = await pushAufgabe(demo);
        expect(r.ok).toBe(true);
        expect(r.skipped).toBe(true);
        expect(r.item?.seedTag).toBe(DEMO_SEED_TAG);
    });

    it('403 beim Pull wipe’t lokalen Cache nicht', async () => {
        const store = await import('../src/tools/aufgaben/aufgaben-store.js');
        store.clearAll();
        store.upsertAufgabe({
            id: 'af-local-keep',
            title: 'Lokal behalten',
            zuweisungName: 'Me',
            zuweisungEmail: 'me@schule.at',
            prioritaet: 'mittel'
        });
        localStorage.setItem(
            'ms365-aufgaben-sp-settings-v1',
            JSON.stringify({
                siteUrl: 'https://contoso.sharepoint.com/sites/schule',
                listAufgaben: 'Aufgaben'
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

        const before = store.loadState();
        expect(before.aufgaben).toHaveLength(1);

        const r = await pullAufgabenFromSharePoint({ forceDemoIfEmpty: false });
        expect(r.ok).toBe(false);
        expect(r.softError).toMatch(/Kein SharePoint-Zugriff/i);
        expect(store.loadState().aufgaben).toHaveLength(1);
        expect(store.loadState().aufgaben[0].id).toBe('af-local-keep');
        expect(getAufgabenSyncMeta().source).toBe('local');
    });
});

describe('aufgaben PA recipe', () => {
    it('katalog enthält aufgaben-mail', () => {
        const here = dirname(fileURLToPath(import.meta.url));
        const src = readFileSync(
            join(here, '../src/tools/power-automate/pa-recipes-catalog.js'),
            'utf8'
        );
        expect(src).toMatch(/id:\s*'aufgaben-mail'/);
    });
});

describe('aufgaben filter smoke roles', () => {
    const items = [
        {
            id: '1',
            title: 'A',
            status: 'offen',
            zuweisungEmail: 'me@schule.at',
            erstelltVonEmail: 'other@schule.at',
            prioritaet: 'mittel',
            beschreibung: '',
            zuweisungName: 'Me',
            rolle: '',
            kommentar: ''
        },
        {
            id: '2',
            title: 'B',
            status: 'offen',
            zuweisungEmail: 'other@schule.at',
            erstelltVonEmail: 'me@schule.at',
            prioritaet: 'mittel',
            beschreibung: '',
            zuweisungName: 'Other',
            rolle: '',
            kommentar: ''
        },
        {
            id: '3',
            title: 'C',
            status: 'offen',
            zuweisungEmail: 'x@schule.at',
            erstelltVonEmail: 'y@schule.at',
            prioritaet: 'mittel',
            beschreibung: '',
            zuweisungName: 'X',
            rolle: '',
            kommentar: ''
        }
    ];

    it('Melder sieht ohne showAll nur eigene (mir+von_mir)', () => {
        const list = filterAufgaben(items, {
            scope: '',
            meEmail: 'me@schule.at',
            admin: false,
            showAll: false
        });
        expect(list.map((a) => a.id).sort()).toEqual(['1', '2']);
    });

    it('Scope mir nur zugewiesen', () => {
        expect(
            filterAufgaben(items, {
                scope: 'mir',
                meEmail: 'me@schule.at',
                admin: false
            }).map((a) => a.id)
        ).toEqual(['1']);
    });

    it('Admin+showAll sieht alle', () => {
        expect(
            filterAufgaben(items, {
                scope: '',
                meEmail: 'me@schule.at',
                admin: true,
                showAll: true
            })
        ).toHaveLength(3);
    });
});
