import { describe, it, expect } from 'vitest';
import { mapItemFromSp, mapItemToFields } from '../src/tools/lehrer-freistellung-planer/lfr-graph.js';
import { LFR_COLUMNS, STATUS_CHOICES } from '../src/tools/lehrer-freistellung-planer/lfr-schema.js';

describe('lfr-graph mapping', () => {
    it('mapItemFromSp liest Listenfelder inkl. Audit', () => {
        const row = mapItemFromSp({
            id: '77',
            fields: {
                Title: 'Fortbildung Graph',
                AntragId: 'lfr-abc',
                Beginn: '2026-11-12T08:00:00Z',
                Ende: '2026-11-12T16:00:00Z',
                Status: 'Genehmigt',
                Kategorie: 'Fortbildung',
                Beschreibung: 'PH-Workshop',
                LehrerName: 'Ada Lehrer',
                LehrerEmail: 'Ada.Lehrer@Schule.at',
                GenehmigtVonDirektion: 'Direktion',
                GenehmigtAmDirektion: '2026-11-01',
                BemerkungDirektion: 'ok'
            }
        });
        expect(row.itemId).toBe('77');
        expect(row.antragId).toBe('lfr-abc');
        expect(row.titel).toBe('Fortbildung Graph');
        expect(row.status).toBe('Genehmigt');
        expect(row.kategorie).toBe('Fortbildung');
        expect(row.lehrerEmail).toBe('ada.lehrer@schule.at');
        expect(row.genehmigtVon).toBe('Direktion');
        expect(row.genehmigtAm).toBe('2026-11-01');
        expect(row.bemerkungDirektion).toBe('ok');
    });

    it('mapItemFromSp akzeptiert Alias Start und normalisiert Status', () => {
        const row = mapItemFromSp({
            id: '1',
            fields: {
                Title: 'Arzt',
                Start: '2026-12-03T10:00:00',
                Ende: '2026-12-03T12:00:00',
                Status: 'genehmigt'
            }
        });
        expect(row.beginn).toContain('2026-12-03');
        expect(row.status).toBe('Genehmigt');
    });

    it('mapItemToFields schreibt Schema-Spalten und AntragId', () => {
        const fields = mapItemToFields({
            antragId: 'lfr-fixed',
            titel: 'Dienstfreistellung',
            beginn: '2026-11-20T08:00',
            ende: '2026-11-20T12:00',
            status: 'Ausstehend',
            kategorie: 'Dienstfreistellung',
            beschreibung: 'Amtsweg',
            lehrerName: 'Bert Beispiel',
            lehrerEmail: 'Bert@Schule.at',
            genehmigtVon: '',
            genehmigtAm: '',
            abgelehntVon: '',
            abgelehntAm: '',
            bemerkungDirektion: ''
        });
        expect(fields.Title).toBe('Dienstfreistellung');
        expect(fields.AntragId).toBe('lfr-fixed');
        expect(fields.Status).toBe('Ausstehend');
        expect(fields.Kategorie).toBe('Dienstfreistellung');
        expect(fields.LehrerEmail).toBe('bert@schule.at');
        expect(fields.Beginn).toBeTruthy();
        expect(fields.Ende).toBeTruthy();
        expect(fields.GenehmigtVonDirektion).toBeUndefined();
        expect(Object.keys(fields)).toEqual(
            expect.arrayContaining(['Title', 'Beginn', 'Ende', 'Status', 'Kategorie', 'LehrerName', 'LehrerEmail', 'AntragId'])
        );
    });

    it('mapItemToFields kann AntragId weglassen und setzt Ablehnung', () => {
        const fields = mapItemToFields(
            {
                titel: 'X',
                beginn: '2026-11-21T09:00',
                ende: '2026-11-21T10:00',
                status: 'Abgelehnt',
                kategorie: 'Sonstiges',
                lehrerEmail: 'a@b.at',
                abgelehntVon: 'Direktion',
                abgelehntAm: '2026-11-18'
            },
            { includeId: false }
        );
        expect(fields.AntragId).toBeUndefined();
        expect(fields.AbgelehntVon).toBe('Direktion');
        expect(fields.AbgelehntAm).toBe('2026-11-18');
        expect(fields.Status).toBe('Abgelehnt');
    });

    it('Schema-Spalten decken gemappte Persistenzfelder ab', () => {
        const names = LFR_COLUMNS.map((c) => c.name);
        expect(names).toEqual(
            expect.arrayContaining([
                'Beginn',
                'Ende',
                'Status',
                'Kategorie',
                'Beschreibung',
                'LehrerName',
                'LehrerEmail',
                'AntragId',
                'GenehmigtVonDirektion',
                'GenehmigtAmDirektion',
                'AbgelehntVon',
                'AbgelehntAm',
                'BemerkungDirektion'
            ])
        );
        expect(STATUS_CHOICES).toEqual(['Ausstehend', 'Genehmigt', 'Abgelehnt']);
    });
});
