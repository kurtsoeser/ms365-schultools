import { describe, expect, it } from 'vitest';
import {
    mapAntwortFromItem,
    mapAntwortToFields,
    mapKopfFromItem,
    mapKopfToFields
} from '../src/tools/einsammler/einsammler-graph.js';
import { isDemoSeedItem } from '../src/tools/einsammler/einsammler-model.js';

describe('einsammler-graph mapping', () => {
    it('mapKopf roundtrip fields + spoItemId', () => {
        const fields = mapKopfToFields({
            id: 'ein-1',
            titel: 'Museumsbesuch',
            modus: 'einwilligung',
            typ: 'exkursion',
            textInhalt: 'Bitte zustimmen',
            klasseCodes: '3AHW',
            faelligAm: '2026-12-01',
            status: 'aktiv',
            uploadErlaubt: false,
            uploadPflicht: false,
            erstellerEmail: 'lehrer@schule.at',
            erstellerName: 'Lehrkraft',
            quelle: 'ui'
        });
        expect(fields.Title).toBe('Museumsbesuch');
        expect(fields.EinsammlerId).toBe('ein-1');
        expect(fields.Modus).toBe('einwilligung');
        expect(fields.ErstellerEmail).toBe('lehrer@schule.at');
        expect(fields.UploadErlaubt).toBe(false);

        const mapped = mapKopfFromItem({
            id: 77,
            createdDateTime: '2026-11-18T10:00:00Z',
            fields: { ...fields, Status: 'geschlossen' }
        });
        expect(mapped.spoItemId).toBe('77');
        expect(mapped.id).toBe('ein-1');
        expect(mapped.status).toBe('geschlossen');
        expect(mapped.titel).toBe('Museumsbesuch');
    });

    it('mapAntwort roundtrip + AnhangUrl', () => {
        const fields = mapAntwortToFields({
            id: 'ant-1',
            einsammlerId: 'ein-1',
            schuelerName: 'Anna Berger',
            klasseCode: '3AHW',
            elternEmail: 'berger@example.com',
            status: 'zugestimmt',
            antwortAm: '2026-11-18T12:00:00.000Z',
            kommentar: 'ok',
            anhangName: 'ausweis.pdf',
            anhangMime: 'application/pdf',
            anhangUrl: 'https://contoso.sharepoint.com/sites/x/ausweis.pdf',
            quelle: 'public'
        });
        expect(fields.AntwortId).toBe('ant-1');
        expect(fields.EinsammlerId).toBe('ein-1');
        expect(fields.Status).toBe('zugestimmt');
        expect(fields.AnhangUrl).toContain('ausweis.pdf');
        expect(fields.Quelle).toBe('public');

        const mapped = mapAntwortFromItem({
            id: 12,
            fields
        });
        expect(mapped.spoItemId).toBe('12');
        expect(mapped.id).toBe('ant-1');
        expect(mapped.schuelerName).toBe('Anna Berger');
        expect(mapped.anhangUrl).toContain('ausweis.pdf');
    });

    it('Demo-Quelle wird für SPO auf ui normalisiert (Choice-Set)', () => {
        const fields = mapKopfToFields({
            id: 'ein-d',
            titel: 'Demo',
            modus: 'einwilligung',
            typ: 'sonstiges',
            status: 'entwurf',
            erstellerEmail: 'x@y.at',
            quelle: 'demo',
            seedTag: 'einsammler-demo-v1'
        });
        expect(fields.Quelle).toBe('ui');
        expect(isDemoSeedItem({ quelle: 'demo', seedTag: 'einsammler-demo-v1' })).toBe(true);
    });
});
