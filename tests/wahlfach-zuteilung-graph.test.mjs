import { describe, expect, it } from 'vitest';
import {
    mapAnmeldungFromItem,
    mapAnmeldungToFields,
    mapAngebotFromItem,
    mapAngebotToFields,
    mapRundeFromItem,
    mapRundeToFields,
    mapZuteilungFromItem,
    mapZuteilungToFields
} from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-graph.js';
import { isWzDemoEntity } from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-sync.js';
import { DEMO_SEED_TAG } from '../src/tools/wahlfach-zuteilung/wahlfach-zuteilung-demo.js';
import { SCHOOL_TOOL_SITE_KINDS } from '../src/shared/school-tool-sites.js';

describe('wahlfach SPO mapping', () => {
    it('mapRunde ToFields / FromItem behält spoItemId', () => {
        const fields = mapRundeToFields({
            id: 'runde-1',
            schuljahr: '2026/27',
            titel: 'Wahlfach Auhof',
            anmeldeVon: '2026-11-01',
            anmeldeBis: '2026-11-30',
            status: 'offen'
        });
        expect(fields.RundeId).toBe('runde-1');
        expect(fields.Title).toBe('Wahlfach Auhof');
        expect(fields.Status).toBe('offen');

        const mapped = mapRundeFromItem({ id: 11, fields });
        expect(mapped.spoItemId).toBe('11');
        expect(mapped.id).toBe('runde-1');
        expect(mapped.status).toBe('offen');
    });

    it('mapAngebot / mapAnmeldung / mapZuteilung Rundtrip Kernfelder', () => {
        const angFields = mapAngebotToFields(
            {
                id: 'ang-1',
                bezeichnung: 'Chor',
                kapazitaetMin: 10,
                kapazitaetMax: 20,
                zielJahrgaenge: '1–4',
                aktiv: true
            },
            { runde: { id: 'runde-1', schuljahr: '2026/27' } }
        );
        expect(angFields.AngebotId).toBe('ang-1');
        expect(angFields.RundeId).toBe('runde-1');
        expect(angFields.KapazitaetMax).toBe(20);

        const ang = mapAngebotFromItem({ id: 22, fields: angFields });
        expect(ang.spoItemId).toBe('22');
        expect(ang.bezeichnung).toBe('Chor');

        const anmFields = mapAnmeldungToFields(
            {
                id: 'anm-1',
                schuelerName: 'Anna',
                klasse: '3AK',
                schuelerEmail: 'anna@schule.at',
                wunsch1Id: 'ang-1',
                wunsch2Id: 'ang-2',
                status: 'eingereicht'
            },
            {
                runde: { id: 'runde-1', schuljahr: '2026/27' },
                angebote: [
                    { id: 'ang-1', bezeichnung: 'Chor' },
                    { id: 'ang-2', bezeichnung: 'Robotik' }
                ]
            }
        );
        expect(anmFields.AnmeldungId).toBe('anm-1');
        expect(anmFields.SchuelerEmail).toBe('anna@schule.at');
        expect(anmFields.Wunsch1Text).toBe('Chor');

        const anm = mapAnmeldungFromItem({ id: 33, fields: anmFields });
        expect(anm.spoItemId).toBe('33');
        expect(anm.wunsch1Id).toBe('ang-1');

        const zFields = mapZuteilungToFields(
            {
                anmeldungId: 'anm-1',
                angebotId: 'ang-1',
                begruendung: 'Manuell',
                zugeteiltVon: 'admin@schule.at',
                zugeteiltAm: '2026-11-18T10:00:00.000Z'
            },
            {
                runde: { id: 'runde-1', schuljahr: '2026/27' },
                angebote: [{ id: 'ang-1', bezeichnung: 'Chor' }],
                anmeldungen: [{ id: 'anm-1', schuelerName: 'Anna', klasse: '3AK' }]
            }
        );
        expect(zFields.AngebotText).toBe('Chor');
        expect(zFields.AccessExportKey).toContain('anm-1');

        const z = mapZuteilungFromItem({ id: 44, fields: zFields });
        expect(z.spoItemId).toBe('44');
        expect(z.angebotId).toBe('ang-1');
        expect(z.zugeteiltVon).toBe('admin@schule.at');
    });

    it('Demo-Entities werden erkannt (kein Push)', () => {
        expect(isWzDemoEntity({ id: 'ang-demo-chor' })).toBe(true);
        expect(isWzDemoEntity({ id: 'anm-1', seedTag: DEMO_SEED_TAG })).toBe(true);
        expect(isWzDemoEntity({ id: 'ang-live-1' })).toBe(false);
    });

    it('Site-Kind wahlfach ist registriert', () => {
        expect(SCHOOL_TOOL_SITE_KINDS.some((k) => k.id === 'wahlfach')).toBe(true);
    });
});
