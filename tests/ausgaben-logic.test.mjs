import { describe, expect, it, beforeEach } from 'vitest';
import {
    findOpenVorgang,
    filterGegenstaende,
    inventarStats,
    parseGegenstaendeCsv,
    validateGegenstand,
    validateVorgang
} from '../src/tools/ausgaben/ausgaben-logic.js';
import { LIST_TITLES, GEGENSTAENDE_COLUMNS } from '../src/tools/ausgaben/ausgaben-schema.js';
import { TYP_VORLAGEN, materializeTypVorlage } from '../src/tools/ausgaben/ausgaben-templates.js';
import {
    DEMO_SEED_TAG,
    VORGANG_STATUS,
    isDemoSeedItem
} from '../src/tools/ausgaben/ausgaben-model.js';
import { buildAusgabenDemoState } from '../src/tools/ausgaben/ausgaben-demo.js';
import {
    mapGegenstandFromItem,
    mapGegenstandToFields,
    mapVorgangFromItem,
    mapVorgangToFields
} from '../src/tools/ausgaben/ausgaben-graph.js';
import { mergeTypenFromGegenstaende } from '../src/tools/ausgaben/ausgaben-sync.js';
import { resolveSchoolToolSiteUrl } from '../src/shared/school-tool-sites.js';

describe('ausgaben-templates', () => {
    it('liefert Vorlagen und materialisiert Beispiele', () => {
        expect(TYP_VORLAGEN.length).toBeGreaterThanOrEqual(4);
        const m = materializeTypVorlage('schluessel', { withExamples: true });
        expect(m?.typ.id).toBe('schluessel');
        expect(m?.gegenstaende.length).toBeGreaterThan(0);
    });
});

describe('ausgaben-logic', () => {
    it('validiert Gegenstand und Vorgang', () => {
        expect(validateGegenstand({ kennung: '', typ: 'schluessel' }).ok).toBe(false);
        expect(validateGegenstand({ kennung: 'S-1', typ: 'schluessel' }).ok).toBe(true);
        expect(
            validateVorgang({
                gegenstandId: 'g1',
                empfaengerName: 'Müller',
                empfaengerTyp: 'lehrer',
                status: 'ausgegeben'
            }).ok
        ).toBe(true);
    });

    it('findet offene Vorgänge und filtert verfügbar', () => {
        const gegenstaende = [
            { id: 'g1', typ: 'schluessel', kennung: 'S-1', bezeichnung: 'A', ortZone: '', aktiv: true, typLabel: '' },
            { id: 'g2', typ: 'schluessel', kennung: 'S-2', bezeichnung: 'B', ortZone: '', aktiv: true, typLabel: '' }
        ];
        const vorgaenge = [
            { id: 'v1', gegenstandId: 'g1', status: VORGANG_STATUS.AUSGEGEBEN, kennung: 'S-1', empfaengerName: 'X' }
        ];
        expect(findOpenVorgang('g1', vorgaenge)).toBeTruthy();
        expect(findOpenVorgang('g2', vorgaenge)).toBeFalsy();
        expect(filterGegenstaende(gegenstaende, vorgaenge, { nurVerfuegbar: true })).toHaveLength(1);
    });

    it('berechnet Stats und parst CSV', () => {
        const st = inventarStats(
            [{ id: '1', aktiv: true }, { id: '2', aktiv: false }],
            [{ status: 'ausgegeben', faelligAm: '2000-01-01' }]
        );
        expect(st.total).toBe(2);
        expect(st.ausgegeben).toBe(1);
        expect(parseGegenstaendeCsv('Typ;Kennung\nschluessel;S-9;Raum 9;A')).toHaveLength(1);
    });
});

describe('ausgaben-schema', () => {
    it('hat Listen und flexible Typ-Spalte', () => {
        expect(LIST_TITLES.gegenstaende).toBe('AusgabenGegenstaende');
        expect(GEGENSTAENDE_COLUMNS.some((c) => c.name === 'Typ' && c.text)).toBe(true);
    });
});

describe('ausgaben-demo / SPO mapping', () => {
    beforeEach(() => {
        const mem = new Map();
        globalThis.localStorage = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => mem.set(k, String(v)),
            removeItem: (k) => mem.delete(k)
        };
    });

    it('Demo enthält seedTag und localDemoOnly', () => {
        const demo = buildAusgabenDemoState();
        expect(demo.localDemoOnly).toBe(true);
        expect(demo.gegenstaende.every(isDemoSeedItem)).toBe(true);
        expect(demo.vorgaenge.every(isDemoSeedItem)).toBe(true);
        expect(DEMO_SEED_TAG).toBe('ausgaben-demo-v1');
    });

    it('mapGegenstand ToFields / FromItem rundtrip Kernfelder', () => {
        const fields = mapGegenstandToFields({
            id: 'gg-1',
            typ: 'schluessel',
            typLabel: 'Schlüssel',
            kennung: 'S-12b',
            bezeichnung: 'Klassenzimmer 12b',
            ortZone: 'Trakt A',
            aktiv: true,
            notiz: 'Demo'
        });
        expect(fields.GegenstandId).toBe('gg-1');
        expect(fields.Typ).toBe('schluessel');
        expect(fields.Kennung).toBe('S-12b');
        expect(fields.Aktiv).toBe(true);
        const mapped = mapGegenstandFromItem({
            id: '99',
            fields
        });
        expect(mapped.spoItemId).toBe('99');
        expect(mapped.kennung).toBe('S-12b');
        expect(mapped.typ).toBe('schluessel');
    });

    it('mapVorgang ToFields / FromItem rundtrip Kernfelder', () => {
        const fields = mapVorgangToFields({
            id: 'vg-1',
            gegenstandId: 'gg-1',
            kennung: 'S-12b',
            empfaengerTyp: 'lehrer',
            empfaengerName: 'Müller',
            empfaengerEmail: 'mueller@schule.demo',
            status: 'ausgegeben',
            ausgegebenAm: '2026-11-18',
            ausgegebenVon: 'Hausverwaltung'
        });
        expect(fields.VorgangId).toBe('vg-1');
        expect(fields.EmpfaengerName).toBe('Müller');
        expect(fields.Status).toBe('ausgegeben');
        const mapped = mapVorgangFromItem({ id: '55', fields });
        expect(mapped.spoItemId).toBe('55');
        expect(mapped.gegenstandId).toBe('gg-1');
        expect(mapped.empfaengerEmail).toBe('mueller@schule.demo');
    });

    it('mergeTypenFromGegenstaende ergänzt fehlende Typen', () => {
        const typen = mergeTypenFromGegenstaende(
            [{ typ: 'beamer_koffer', typLabel: 'Beamer-Koffer' }],
            [{ id: 'schluessel', label: 'Schlüssel' }]
        );
        expect(typen.some((t) => t.id === 'schluessel')).toBe(true);
        expect(typen.some((t) => t.id === 'beamer_koffer' && t.label === 'Beamer-Koffer')).toBe(
            true
        );
    });

    it('resolveSchoolToolSiteUrl(ausgaben) liest SP-Settings', () => {
        expect(resolveSchoolToolSiteUrl('ausgaben')).toBe('');
        globalThis.localStorage.setItem(
            'ms365-ausgaben-sp-settings-v1',
            JSON.stringify({ siteUrl: 'https://contoso.sharepoint.com/sites/ausgaben-local' })
        );
        expect(resolveSchoolToolSiteUrl('ausgaben')).toBe(
            'https://contoso.sharepoint.com/sites/ausgaben-local'
        );
    });
});
