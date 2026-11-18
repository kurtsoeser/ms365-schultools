import { describe, expect, it, beforeEach } from 'vitest';
import {
    canChangeStatus,
    canCommentOnMeldung,
    filterMeldungen,
    meldungenStats,
    sortMeldungen,
    statusPipelineActions,
    validateMeldung
} from '../src/tools/technik-meldungen/technik-meldungen-logic.js';
import {
    DEMO_SEED_TAG,
    PRIORITAET,
    STATUS,
    isDemoSeedItem
} from '../src/tools/technik-meldungen/technik-meldungen-model.js';
import { buildTechnikDemoState } from '../src/tools/technik-meldungen/technik-meldungen-demo.js';
import { resolveSchoolToolSiteUrl, SCHOOL_TOOL_SITE_KINDS } from '../src/shared/school-tool-sites.js';
import {
    mapMeldungFromItem,
    mapMeldungToFields
} from '../src/tools/technik-meldungen/technik-meldungen-graph.js';

describe('technik-meldungen-logic', () => {
    it('validateMeldung verlangt Titel und Raum', () => {
        expect(validateMeldung({ title: '', raumCode: '' }).ok).toBe(false);
        expect(
            validateMeldung({
                title: 'Beamer',
                raumCode: 'PHY1',
                kategorie: 'beamer',
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

    it('canChangeStatus nur Admin; Kommentar Melder oder Admin', () => {
        const m = { id: '1', melderEmail: 'me@schule.at', status: STATUS.OFFEN };
        expect(canChangeStatus(m, { admin: false })).toBe(false);
        expect(canChangeStatus(m, { admin: true })).toBe(true);
        expect(canCommentOnMeldung(m, { meEmail: 'me@schule.at' })).toBe(true);
        expect(canCommentOnMeldung(m, { meEmail: 'other@schule.at' })).toBe(false);
        expect(canCommentOnMeldung(m, { meEmail: 'other@schule.at', admin: true })).toBe(true);
    });

    it('sortiert nach Priorität dann Datum', () => {
        const sorted = sortMeldungen([
            {
                id: 'a',
                title: 'A',
                prioritaet: PRIORITAET.MITTEL,
                gemeldetAm: '2026-01-02'
            },
            {
                id: 'b',
                title: 'B',
                prioritaet: PRIORITAET.DRINGEND,
                gemeldetAm: '2026-01-01'
            }
        ]);
        expect(sorted.map((m) => m.id)).toEqual(['b', 'a']);
    });

    it('meldungenStats zählt offen/dringend', () => {
        const st = meldungenStats([
            { status: STATUS.OFFEN, prioritaet: PRIORITAET.DRINGEND },
            { status: STATUS.IN_ARBEIT, prioritaet: PRIORITAET.MITTEL },
            { status: STATUS.ERLEDIGT, prioritaet: PRIORITAET.HOCH }
        ]);
        expect(st.offen).toBe(2);
        expect(st.dringend).toBe(1);
        expect(st.erledigt).toBe(1);
    });
});

describe('technik-meldungen-demo / store actor', () => {
    beforeEach(() => {
        const mem = new Map();
        globalThis.localStorage = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => mem.set(k, String(v)),
            removeItem: (k) => mem.delete(k)
        };
    });

    it('Demo enthält seedTag und Schüler-Melder', () => {
        const demo = buildTechnikDemoState();
        expect(demo.localDemoOnly).toBe(true);
        expect(demo.meldungen.every(isDemoSeedItem)).toBe(true);
        expect(demo.meldungen.some((m) => m.melderEmail === 'schueler.demo@schule.demo')).toBe(
            true
        );
        expect(DEMO_SEED_TAG).toBe('technik-demo-v1');
    });

    it('setMeldungStatus speichert Actor', async () => {
        const store = await import('../src/tools/technik-meldungen/technik-meldungen-store.js');
        store.clearAll();
        const created = store.upsertMeldung({
            id: 'tm-actor-1',
            title: 'Test',
            raumCode: 'A1',
            kategorie: 'pc',
            prioritaet: 'mittel',
            melderEmail: 'me@schule.at',
            melderName: 'Me'
        });
        expect(created.ok).toBe(true);
        const r = store.setMeldungStatus('tm-actor-1', STATUS.IN_ARBEIT, {
            actorName: 'IT',
            actorEmail: 'it@schule.at'
        });
        expect(r.ok).toBe(true);
        expect(r.item.bearbeitetVonEmail).toBe('it@schule.at');
        expect(r.item.bearbeitetVonName).toBe('IT');
        expect(r.item.status).toBe(STATUS.IN_ARBEIT);
    });

    it('clearDemoSeed entfernt nur Seed-Zeilen', async () => {
        const store = await import('../src/tools/technik-meldungen/technik-meldungen-store.js');
        store.loadDemoData();
        store.upsertMeldung({
            id: 'tm-real-1',
            title: 'Echt',
            raumCode: 'B1',
            kategorie: 'pc',
            prioritaet: 'niedrig',
            melderEmail: 'real@schule.at',
            melderName: 'Real'
        });
        const after = store.clearDemoSeed();
        expect(after.meldungen).toHaveLength(1);
        expect(after.meldungen[0].id).toBe('tm-real-1');
    });
});

describe('technik site kind', () => {
    it('technik ist in SCHOOL_TOOL_SITE_KINDS', () => {
        expect(SCHOOL_TOOL_SITE_KINDS.some((k) => k.id === 'technik')).toBe(true);
        expect(typeof resolveSchoolToolSiteUrl).toBe('function');
    });
});

describe('technik SPO mapping', () => {
    it('mapMeldungToFields / FromItem rundtrip Kernfelder', () => {
        const fields = mapMeldungToFields({
            id: 'tm-1',
            title: 'Beamer',
            raumCode: 'PHY1',
            kategorie: 'beamer',
            prioritaet: 'hoch',
            status: 'offen',
            melderEmail: 'a@schule.at',
            melderName: 'A',
            beschreibung: 'flackert'
        });
        expect(fields.Title).toBe('Beamer');
        expect(fields.MeldungId).toBe('tm-1');
        expect(fields.Kategorie).toBe('beamer');
        expect(fields.MelderEmail).toBe('a@schule.at');

        const mapped = mapMeldungFromItem({
            id: 42,
            fields: { ...fields, Status: 'in_arbeit' }
        });
        expect(mapped.spoItemId).toBe('42');
        expect(mapped.id).toBe('tm-1');
        expect(mapped.status).toBe('in_arbeit');
        expect(mapped.title).toBe('Beamer');
    });
});

describe('technik filter smoke roles', () => {
    const meldungen = [
        {
            id: '1',
            title: 'A',
            status: 'offen',
            melderEmail: 'schueler@schule.at',
            kategorie: 'pc',
            prioritaet: 'mittel',
            raumCode: 'A',
            raumName: '',
            beschreibung: '',
            melderName: '',
            kommentar: ''
        },
        {
            id: '2',
            title: 'B',
            status: 'offen',
            melderEmail: 'lehrer@schule.at',
            kategorie: 'beamer',
            prioritaet: 'hoch',
            raumCode: 'B',
            raumName: '',
            beschreibung: '',
            melderName: '',
            kommentar: ''
        }
    ];

    it('Schüler sieht nur eigene', () => {
        expect(
            filterMeldungen(meldungen, {
                meEmail: 'schueler@schule.at',
                admin: false,
                showAll: false
            }).map((m) => m.id)
        ).toEqual(['1']);
    });

    it('Lehrer sieht nur eigene', () => {
        expect(
            filterMeldungen(meldungen, {
                meEmail: 'lehrer@schule.at',
                admin: false,
                showAll: false
            }).map((m) => m.id)
        ).toEqual(['2']);
    });

    it('Admin Inbox sieht alle', () => {
        expect(
            filterMeldungen(meldungen, {
                meEmail: 'it@schule.at',
                admin: true,
                showAll: true
            })
        ).toHaveLength(2);
    });
});
