import { describe, it, expect } from 'vitest';
import {
    AKTIVITAETEN_COLUMNS,
    REQUIRED_COLUMNS,
    STATUS_CHOICES,
    TEILNAHME_MODUS_CHOICES
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-schema.js';
import {
    mapAktivitaetFromItem,
    mapAktivitaetToFields
} from '../src/tools/schulaktivitaeten-planer/schulaktivitaeten-planer-graph.js';

/** App-Felder, die dauerhaft in SharePoint liegen müssen. */
const PERSISTED_APP_KEYS = [
    'aktivitaetId',
    'titel',
    'typ',
    'klasseCode',
    'klasseCodes',
    'teilnahmeModus',
    'gruppenText',
    'teilnehmendeText',
    'nichtTeilnehmende',
    'lehrerCode',
    'lehrerEmail',
    'begleitung',
    'begleitungCodes',
    'externeVortragende',
    'ort',
    'ablauf',
    'unterrichtsbezug',
    'startdatum',
    'enddatum',
    'startUe',
    'endUe',
    'startZeit',
    'endZeit',
    'status',
    'notiz',
    'ablehnungsGrund',
    'rueckfrageText',
    'kvStellungnahme',
    'kvStellungnahmeVon',
    'aenderungsLog',
    'beantragtVon',
    'genehmigtVon',
    'untisErledigt',
    'untisErledigtVon',
    'verkehrsmittel',
    'kostenHinweis',
    'schulterminKey'
];

describe('Schulaktivitäten SharePoint-Schema', () => {
    it('REQUIRED_COLUMNS deckt alle Aktivitäts-Spalten ab', () => {
        expect(REQUIRED_COLUMNS.Schulaktivitaeten).toEqual(AKTIVITAETEN_COLUMNS.map((c) => c.name));
        expect(AKTIVITAETEN_COLUMNS.map((c) => c.name)).toEqual(
            expect.arrayContaining([
                'AktivitaetId',
                'LehrerCode',
                'LehrerEmail',
                'BeantragtVon',
                'Begleitung',
                'BegleitungCodes',
                'KlasseCodes',
                'TeilnahmeModus',
                'Status',
                'KvStellungnahme',
                'AblehnungsGrund',
                'Startdatum',
                'Enddatum'
            ])
        );
    });

    it('Status- und Teilnahme-Choices sind stabil', () => {
        expect(STATUS_CHOICES).toContain('kv_stellungnahme');
        expect(STATUS_CHOICES).toContain('nachforderung');
        expect(TEILNAHME_MODUS_CHOICES).toEqual(['klasse', 'klassen', 'gruppe', 'schueler']);
    });
});

describe('mapAktivitaet Roundtrip', () => {
    it('schreibt und liest Personenbeziehungen + Begleitung korrekt', () => {
        const fields = mapAktivitaetToFields(
            {
                aktivitaetId: 'akt-test-1',
                titel: 'Technisches Museum Wien',
                typ: 'Exkursion',
                klasseCodes: '3BK',
                teilnahmeModus: 'klasse',
                lehrerCode: 'MAY',
                lehrerEmail: 'may@schule.at',
                beantragtVon: 'May@Schule.at',
                begleitungCodes: 'COL;WEI',
                begleitung: 'COL;WEI',
                startdatum: '2026-11-25',
                enddatum: '2026-11-25',
                startUe: 3,
                endUe: 6,
                status: 'beantragt',
                ort: 'Wien',
                verkehrsmittel: 'Bus',
                kostenHinweis: '12 €'
            },
            { teacherLabels: { COL: 'Dave Cole', WEI: 'Anna Wei' } }
        );

        expect(fields.Title).toBe('Technisches Museum Wien');
        expect(fields.LehrerEmail).toBe('may@schule.at');
        expect(fields.BeantragtVon).toBe('may@schule.at');
        expect(fields.BegleitungCodes.replace(/\s/g, '')).toBe('COL;WEI');
        expect(fields.Begleitung).toBe('Dave Cole; Anna Wei');
        expect(fields.KlasseCodes.replace(/\s/g, '')).toBe('3BK');
        expect(fields.TeilnahmeModus).toBe('klasse');
        expect(fields.Status).toBe('beantragt');
        expect(fields.StartUe).toBe(3);
        expect(fields.EndUe).toBe(6);

        const mapped = mapAktivitaetFromItem({ id: '99', fields });
        for (const key of PERSISTED_APP_KEYS) {
            expect(mapped).toHaveProperty(key);
        }
        expect(mapped.beantragtVon).toBe('may@schule.at');
        expect(mapped.lehrerEmail).toBe('may@schule.at');
        expect(mapped.begleitungCodes.replace(/\s/g, '')).toBe('COL;WEI');
        expect(mapped.klasseCodes).toContain('3BK');
        expect(mapped.status).toBe('beantragt');
        expect(mapped.ort).toBe('Wien');
        expect(Number(mapped.startUe)).toBe(3);
    });

    it('liest Legacy-Begleitung mit Codes, wenn BegleitungCodes fehlt', () => {
        const mapped = mapAktivitaetFromItem({
            id: '1',
            fields: {
                Title: 'Alt',
                Typ: 'Exkursion',
                Status: 'genehmigt',
                Startdatum: '2026-12-01',
                Begleitung: 'COL;MAY',
                LehrerEmail: 'a@schule.at',
                BeantragtVon: 'a@schule.at'
            }
        });
        expect(mapped.begleitungCodes.replace(/\s/g, '')).toBe('COL;MAY');
    });

    it('bildet Antragsteller und Verantwortliche getrennt ab', () => {
        const fields = mapAktivitaetToFields({
            titel: 'Split',
            lehrerCode: 'MAY',
            lehrerEmail: 'responsible@schule.at',
            beantragtVon: 'applicant@schule.at',
            startdatum: '2026-11-01',
            status: 'beantragt'
        });
        expect(fields.LehrerEmail).toBe('responsible@schule.at');
        expect(fields.BeantragtVon).toBe('applicant@schule.at');
        const mapped = mapAktivitaetFromItem({ id: '2', fields });
        expect(mapped.lehrerEmail).toBe('responsible@schule.at');
        expect(mapped.beantragtVon).toBe('applicant@schule.at');
    });

    it('erzwingt Ownership aus Account-E-Mail wenn Draft leer ist', () => {
        const fields = mapAktivitaetToFields(
            {
                titel: 'Ohne Ownership im Draft',
                startdatum: '2026-11-01',
                status: 'beantragt'
            },
            { accountEmail: 'lehrer01.demo@schule.at', teachers: [] }
        );
        expect(fields.BeantragtVon).toBe('lehrer01.demo@schule.at');
        expect(fields.LehrerEmail).toBe('lehrer01.demo@schule.at');
        expect(fields.LehrerCode).toBeTruthy();
        expect(fields.LehrerCode).toBe('LEHRER01DEMO');
    });

    it('lehnt Speichern ohne E-Mail ab', () => {
        expect(() =>
            mapAktivitaetToFields({
                titel: 'X',
                startdatum: '2026-11-01',
                status: 'beantragt'
            })
        ).toThrow(/Lehrer-E-Mail fehlt/i);
    });
});
