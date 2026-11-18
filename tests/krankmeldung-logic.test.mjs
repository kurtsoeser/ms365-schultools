import { describe, expect, it } from 'vitest';
import {
    buildPublicFormUrl,
    parseKlassenList,
    resolvePublicKlassen
} from '../src/tools/krankmeldung/krankmeldung-klassen.js';
import {
    filterMeldungen,
    findDuplicate,
    inboxStats,
    meldungenToCsv,
    validateMeldung
} from '../src/tools/krankmeldung/krankmeldung-logic.js';
import {
    mapMeldungFromItem,
    mapMeldungToFields
} from '../src/tools/krankmeldung/krankmeldung-graph.js';
import { isDemoSeedItem, DEMO_SEED_TAG } from '../src/tools/krankmeldung/krankmeldung-model.js';
import { LIST_TITLES, MELDUNGEN_COLUMNS } from '../src/tools/krankmeldung/krankmeldung-schema.js';

describe('krankmeldung-logic', () => {
    it('validiert Pflichtfelder', () => {
        expect(validateMeldung({}).ok).toBe(false);
        expect(
            validateMeldung({
                schuelerName: 'Anna',
                klasseCode: '3A',
                von: '2026-11-06',
                bis: '2026-11-06',
                kategorie: 'krankheit'
            }, { today: '2026-11-06' }).ok
        ).toBe(true);
    });

    it('lehnt zu lange Zeiträume und weit entfernte Daten ab', () => {
        expect(
            validateMeldung(
                {
                    schuelerName: 'Anna',
                    klasseCode: '3A',
                    von: '2026-11-01',
                    bis: '2026-12-15',
                    kategorie: 'krankheit'
                },
                { today: '2026-11-06' }
            ).ok
        ).toBe(false);
        expect(
            validateMeldung(
                {
                    schuelerName: 'Anna',
                    klasseCode: '3A',
                    von: '2026-01-01',
                    bis: '2026-01-02',
                    kategorie: 'krankheit'
                },
                { today: '2026-11-06' }
            ).ok
        ).toBe(false);
        expect(
            validateMeldung(
                {
                    schuelerName: 'Anna',
                    klasseCode: '3A',
                    von: '2026-11-06',
                    bis: '2026-11-06',
                    kategorie: 'krankheit'
                },
                { today: '2026-11-06', requireAnhang: true }
            ).ok
        ).toBe(false);
    });

    it('findet Duplikate bei Überlappung', () => {
        const existing = [
            {
                id: '1',
                schuelerName: 'Anna',
                klasseCode: '3A',
                von: '2026-11-05',
                bis: '2026-11-07',
                status: 'neu'
            }
        ];
        expect(
            findDuplicate(
                { schuelerName: 'Anna', klasseCode: '3A', von: '2026-11-06', bis: '2026-11-06' },
                existing
            )
        ).toBeTruthy();
        expect(
            findDuplicate(
                { schuelerName: 'Ben', klasseCode: '3A', von: '2026-11-06', bis: '2026-11-06' },
                existing
            )
        ).toBeFalsy();
    });

    it('filtert nach Tag, Klasse, Suche und Allow-Liste', () => {
        const list = [
            {
                id: '1',
                schuelerName: 'Anna',
                klasseCode: '3A',
                von: '2026-11-06',
                bis: '2026-11-06',
                status: 'neu',
                kategorie: 'krankheit'
            },
            {
                id: '2',
                schuelerName: 'Ben',
                klasseCode: '2B',
                von: '2026-11-07',
                bis: '2026-11-07',
                status: 'neu',
                kategorie: 'arzt'
            }
        ];
        expect(filterMeldungen(list, { day: '2026-11-06' })).toHaveLength(1);
        expect(filterMeldungen(list, { klasse: '2B' })).toHaveLength(1);
        expect(filterMeldungen(list, { q: 'anna' })).toHaveLength(1);
        expect(filterMeldungen(list, { klassenAllow: ['3A'] })).toHaveLength(1);
        expect(filterMeldungen(list, { kategorie: 'arzt' })).toHaveLength(1);
        expect(meldungenToCsv(list).split('\n')).toHaveLength(3);
    });

    it('berechnet Inbox-Stats', () => {
        const list = [
            { status: 'neu', von: '2026-11-06', bis: '2026-11-06', klasseCode: '3A', schuelerName: 'A' },
            { status: 'zur_kenntnis', von: '2026-11-06', bis: '2026-11-06', klasseCode: '3A', schuelerName: 'B' }
        ];
        const st = inboxStats(list, '2026-11-06');
        expect(st.neu).toBe(1);
        expect(st.heute).toBe(2);
        expect(st.offenHeute).toBe(1);
    });
});

describe('krankmeldung-schema', () => {
    it('hat Liste und Spalten inkl. Anhang', () => {
        expect(LIST_TITLES.meldungen).toBe('Krankmeldungen');
        expect(MELDUNGEN_COLUMNS.some((c) => c.name === 'MeldungId')).toBe(true);
        expect(MELDUNGEN_COLUMNS.some((c) => c.name === 'AnhangName')).toBe(true);
    });
});

describe('krankmeldung SPO mapping', () => {
    it('mapMeldungToFields / FromItem rundtrip Kernfelder', () => {
        const fields = mapMeldungToFields({
            id: 'km-1',
            schuelerName: 'Anna Berger',
            klasseCode: '3AHW',
            von: '2026-11-18',
            bis: '2026-11-19',
            kategorie: 'krankheit',
            status: 'neu',
            gemeldetVonEmail: 'berger@example.com',
            quelle: 'public'
        });
        expect(fields.Title).toBe('Anna Berger (3AHW)');
        expect(fields.MeldungId).toBe('km-1');
        expect(fields.SchuelerName).toBe('Anna Berger');
        expect(fields.KlasseCode).toBe('3AHW');
        expect(fields.Kategorie).toBe('krankheit');
        expect(fields.Quelle).toBe('public');
        expect(fields.Von).toBe('2026-11-18');

        const mapped = mapMeldungFromItem({
            id: 77,
            fields: { ...fields, Status: 'zur_kenntnis' }
        });
        expect(mapped.spoItemId).toBe('77');
        expect(mapped.id).toBe('km-1');
        expect(mapped.status).toBe('zur_kenntnis');
        expect(mapped.schuelerName).toBe('Anna Berger');
        expect(mapped.klasseCode).toBe('3AHW');
    });

    it('Demo-Seeds werden erkannt und Quelle demo nicht nach SPO gemappt', () => {
        expect(isDemoSeedItem({ seedTag: DEMO_SEED_TAG, schuelerName: 'X' })).toBe(true);
        expect(isDemoSeedItem({ quelle: 'demo' })).toBe(true);
        const fields = mapMeldungToFields({
            id: 'km-d',
            schuelerName: 'Demo',
            klasseCode: '1A',
            von: '2026-11-18',
            bis: '2026-11-18',
            kategorie: 'krankheit',
            status: 'neu',
            quelle: 'demo',
            seedTag: DEMO_SEED_TAG
        });
        expect(fields.Quelle).toBeUndefined();
    });
});

describe('krankmeldung-store actor', () => {
    it('markKenntnis speichert die Actor-E-Mail', async () => {
        const store = await import('../src/tools/krankmeldung/krankmeldung-store.js');
        const { STATUS } = await import('../src/tools/krankmeldung/krankmeldung-model.js');
        const mem = new Map();
        globalThis.localStorage = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => mem.set(k, String(v)),
            removeItem: (k) => mem.delete(k)
        };
        store.clearAll();
        const created = store.upsertMeldung({
            id: 'km-actor-1',
            schuelerName: 'Anna Test',
            klasseCode: '3A',
            von: '2026-11-18',
            bis: '2026-11-18',
            kategorie: 'krankheit',
            status: STATUS.NEU
        });
        expect(created.ok).toBe(true);
        const r = store.markKenntnis('km-actor-1', 'kv.mueller@schule.at');
        expect(r.ok).toBe(true);
        expect(r.item.kenntnisLehrerEmail).toBe('kv.mueller@schule.at');
        expect(r.item.status).toBe(STATUS.ZUR_KENNTNIS);
    });
});

describe('krankmeldung-klassen', () => {
    it('parst Klassenlisten', () => {
        expect(parseKlassenList('3A; 2B,1C\n1C')).toEqual(['3A', '2B', '1C']);
    });

    it('baut öffentlichen Link mit Klassen', () => {
        expect(buildPublicFormUrl('krankmeldung-formular.html', '3A;2B')).toBe(
            'krankmeldung-formular.html?klassen=3A%3B2B'
        );
        expect(buildPublicFormUrl('krankmeldung-formular.html', '')).toBe(
            'krankmeldung-formular.html'
        );
    });

    it('nimmt Klassen aus der URL', () => {
        expect(resolvePublicKlassen('?klassen=1A;2B')).toEqual(['1A', '2B']);
    });
});
