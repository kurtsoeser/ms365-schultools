import { describe, expect, it } from 'vitest';
import {
    emptyPwSpoSyncMeta,
    isPwLocalOnlyItem,
    shouldSkipPwSpoPush,
    mergePwSpoSyncMeta,
    formatPwSpoSyncLine
} from '../src/tools/projektwochen/projektwochen-sync.js';
import {
    mapAngebotFromItem,
    mapAngebotToFields,
    mapAktionFromItem,
    mapAktionToFields,
    resolveWeekdayFieldName,
    adaptAngebotFieldsForList
} from '../src/tools/projektwochen/projektwochen-graph.js';

describe('projektwochen-sync (SPO Live Meta)', () => {
    it('emptyPwSpoSyncMeta startet ohne Verbindung', () => {
        const m = emptyPwSpoSyncMeta();
        expect(m.source).toBe('none');
        expect(m.connected).toBe(false);
        expect(m.lastPullAt).toBe('');
    });

    it('isPwLocalOnlyItem erkennt local-/demo-IDs', () => {
        expect(isPwLocalOnlyItem('local-ang-1')).toBe(true);
        expect(isPwLocalOnlyItem({ itemId: 'local-aktion-x' })).toBe(true);
        expect(isPwLocalOnlyItem('demo-biz-1')).toBe(true);
        expect(isPwLocalOnlyItem('42')).toBe(false);
        expect(isPwLocalOnlyItem({ itemId: '99' })).toBe(false);
        expect(isPwLocalOnlyItem('')).toBe(true);
    });

    it('shouldSkipPwSpoPush nur bei Demo oder ohne Context', () => {
        expect(shouldSkipPwSpoPush({ localDemoOnly: true, ctx: { siteId: 'x' } })).toBe(true);
        expect(shouldSkipPwSpoPush({ localDemoOnly: false, ctx: null })).toBe(true);
        expect(shouldSkipPwSpoPush({ localDemoOnly: false, ctx: { siteId: 'x' } })).toBe(false);
    });

    it('mergePwSpoSyncMeta und formatPwSpoSyncLine', () => {
        const m = mergePwSpoSyncMeta(null, {
            source: 'live',
            connected: true,
            lists: 'PW-Aktionen, PW-Angebote',
            lastPullAt: '2026-11-18T10:00:00.000Z'
        });
        expect(m.source).toBe('live');
        expect(formatPwSpoSyncLine(m)).toMatch(/Quelle: live/);
        expect(formatPwSpoSyncLine(m)).toMatch(/verbunden/);
    });
});

describe('projektwochen-graph Mapping', () => {
    it('mapAngebot From/To Roundtrip Kernfelder', () => {
        const item = {
            id: '17',
            fields: {
                Title: 'Zoo Wien',
                AngebotId: 'ang-zoo',
                AktionId: 'pw-2026',
                Beschreibung: 'Exkursion',
                HinweisEltern: 'Einverständnis empfohlen',
                Ort: 'Schönbrunn',
                Treffpunkt: 'Tor 1',
                Wochentag: 'Mo',
                Datum: '2026-06-08',
                Slot: 'vormittag',
                Startzeit: '08:00',
                Endzeit: '12:00',
                Kapazitaet: 25,
                PreisEuro: 12,
                Zielklassen: '1AK,1BK',
                LehrerCode: 'MU',
                LehrerEmail: 'mu@schule.at',
                Kategorie: 'exkursion',
                Status: 'freigegeben',
                BuchungAb: '2026-05-01T08:00:00Z',
                BeantragtVon: 'mu@schule.at'
            }
        };
        const mapped = mapAngebotFromItem(item);
        expect(mapped.itemId).toBe('17');
        expect(mapped.title).toBe('Zoo Wien');
        expect(mapped.tag).toBe('Mo');
        expect(mapped.kapazitaet).toBe(25);
        expect(mapped.status).toBe('freigegeben');

        const fields = mapAngebotToFields(mapped, { weekdayField: 'Wochentag' });
        expect(fields.Title).toBe('Zoo Wien');
        expect(fields.AngebotId).toBe('ang-zoo');
        expect(fields.Wochentag).toBe('Mo');
        expect(fields.Kapazitaet).toBe(25);
        expect(fields.Status).toBe('freigegeben');
        expect(fields.BuchungAb).toBeTruthy();
    });

    it('mapAktion From/To', () => {
        const mapped = mapAktionFromItem({
            id: '3',
            fields: {
                Title: 'PW 26/27',
                AktionId: 'pw-2027-06',
                Startdatum: '2026-06-08',
                Enddatum: '2026-06-12',
                Status: 'offen',
                BookingsBusinessId: 'biz-1'
            }
        });
        expect(mapped.itemId).toBe('3');
        expect(mapped.aktionId).toBe('pw-2027-06');
        const fields = mapAktionToFields(mapped);
        expect(fields.AktionId).toBe('pw-2027-06');
        expect(fields.Status).toBe('offen');
        expect(fields.BookingsBusinessId).toBe('biz-1');
    });

    it('Legacy Tag-Spalte beim Schreiben', () => {
        expect(resolveWeekdayFieldName(['Tag'])).toBe('Tag');
        const adapted = adaptAngebotFieldsForList({ Title: 'Kino', Wochentag: 'Di' }, 'Tag');
        expect(adapted.Tag).toBe('Di');
        expect(adapted.Wochentag).toBeUndefined();
        const fields = mapAngebotToFields(
            { title: 'Kino', tag: 'Di', angebotId: 'a1', aktionId: 'p1', kapazitaet: 10 },
            { weekdayField: 'Tag', includeAdminFields: false }
        );
        expect(fields.Tag).toBe('Di');
        expect(fields.Wochentag).toBeUndefined();
        expect(fields.BuchungAb).toBeUndefined();
    });
});
