import { describe, expect, it } from 'vitest';
import {
    mapBriefFromItem,
    mapBriefToFields,
    mapEmpfaengerFromItem,
    mapEmpfaengerToFields
} from '../src/tools/elternbrief/elternbrief-graph.js';
import { isDemoSeedItem, ELTERNBRIEF_DEMO_SEED_TAG } from '../src/tools/elternbrief/elternbrief-model.js';
import { resolveSchoolToolSiteUrl, SCHOOL_TOOL_SITE_KINDS } from '../src/shared/school-tool-sites.js';

describe('elternbrief site kind', () => {
    it('elternbrief ist in SCHOOL_TOOL_SITE_KINDS', () => {
        expect(SCHOOL_TOOL_SITE_KINDS.some((k) => k.id === 'elternbrief')).toBe(true);
        expect(typeof resolveSchoolToolSiteUrl).toBe('function');
    });
});

describe('elternbrief SPO mapping', () => {
    it('mapBriefToFields / FromItem rundtrip Kernfelder', () => {
        const fields = mapBriefToFields({
            id: 'eb-1',
            betreff: 'Schulbeginn',
            inhalt: 'Willkommen im neuen Jahr',
            klasseCodes: '3A;3B',
            absenderModus: 'eigenes_konto',
            absenderEmail: 'lehrer@schule.at',
            absenderName: 'Lehrkraft',
            status: 'gesendet',
            erstellerEmail: 'lehrer@schule.at',
            erstellerName: 'Lehrkraft',
            lesebestaetigung: true,
            faelligAm: '2026-12-01'
        });
        expect(fields.Title).toBe('Schulbeginn');
        expect(fields.BriefId).toBe('eb-1');
        expect(fields.Betreff).toBe('Schulbeginn');
        expect(fields.KlasseCodes).toBe('3A;3B');
        expect(fields.ErstellerEmail).toBe('lehrer@schule.at');
        expect(fields.Lesebestaetigung).toBe(true);

        const mapped = mapBriefFromItem({
            id: 99,
            fields: { ...fields, Status: 'geschlossen' }
        });
        expect(mapped.spoItemId).toBe('99');
        expect(mapped.id).toBe('eb-1');
        expect(mapped.status).toBe('geschlossen');
        expect(mapped.betreff).toBe('Schulbeginn');
    });

    it('mapEmpfaengerToFields / FromItem rundtrip', () => {
        const fields = mapEmpfaengerToFields({
            id: 'el-1',
            briefId: 'eb-1',
            schuelerName: 'Max Muster',
            klasseCode: '3A',
            elternEmail: 'eltern@example.com',
            status: 'offen'
        });
        expect(fields.Title).toBe('Max Muster');
        expect(fields.LeseId).toBe('el-1');
        expect(fields.BriefId).toBe('eb-1');
        expect(fields.ElternEmail).toBe('eltern@example.com');

        const mapped = mapEmpfaengerFromItem({
            id: 7,
            fields: { ...fields, Status: 'gelesen', GelesenAm: '2026-11-18T10:00:00.000Z' }
        });
        expect(mapped.spoItemId).toBe('7');
        expect(mapped.id).toBe('el-1');
        expect(mapped.status).toBe('gelesen');
        expect(mapped.gelesenAm).toContain('2026-11-18');
    });

    it('Demo-Seeds werden erkannt und nicht gepusht', () => {
        expect(
            isDemoSeedItem({
                id: 'x',
                seedTag: ELTERNBRIEF_DEMO_SEED_TAG,
                betreff: 'Demo'
            })
        ).toBe(true);
        expect(isDemoSeedItem({ id: 'y', quelle: 'demo' })).toBe(true);
        expect(isDemoSeedItem({ id: 'z', betreff: 'Echt', quelle: 'ui' })).toBe(false);
    });
});
