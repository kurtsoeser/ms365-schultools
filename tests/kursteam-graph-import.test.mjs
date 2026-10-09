import { describe, it, expect } from 'vitest';
import {
    parseKursteamDisplayName,
    parseFachLehrerSegment,
    mailNicknameMatchesKursteamFilter,
    graphGroupToBelegungRow,
    mergeBelegungWithGraphImport
} from '../src/shared/kursteam-graph-import-logic.js';

describe('kursteam-graph-import-logic', () => {
    it('parst HAK-Steyr-Anzeigenamen', () => {
        const p = parseKursteamDisplayName('SJ26-27 | 1A | BW-FRECH');
        expect(p.ok).toBe(true);
        expect(p.yearPrefix).toBe('SJ26-27');
        expect(p.klasse).toBe('1A');
        expect(p.fach).toBe('BW');
        expect(p.lehrerCode).toBe('FRECH');
    });

    it('parst Fach mit mehreren Buchstaben', () => {
        expect(parseFachLehrerSegment('BESPK-RINNE')).toEqual({ fach: 'BESPK', lehrerCode: 'RINNE' });
        expect(parseFachLehrerSegment('ENWS-AUER')).toEqual({ fach: 'ENWS', lehrerCode: 'AUER' });
    });

    it('filtert mailNickname nach Schuljahr-Präfix', () => {
        const nick = 'SJ26-27-jg2031-hakb-BW-FRECH';
        expect(mailNicknameMatchesKursteamFilter(nick, { yearPrefix: 'SJ26-27' })).toBe(true);
        expect(mailNicknameMatchesKursteamFilter(nick, { yearPrefix: 'SJ25-26' })).toBe(false);
        expect(
            mailNicknameMatchesKursteamFilter('demo-sj26-27-jgb-ges-pag', { yearPrefix: 'DEMO SJ26-27' })
        ).toBe(true);
    });

    it('erkennt DEMO-Teams auch per Anzeigename', () => {
        const row = graphGroupToBelegungRow(
            {
                id: 'gid-demo',
                displayName: 'DEMO SJ26-27 | 2B | GES | PAG',
                mailNickname: 'demo-sj26-27-jgb-ges-pag'
            },
            { yearPrefix: 'DEMO SJ26-27' }
        );
        expect(row).not.toBeNull();
        expect(row.graphGroupId).toBe('gid-demo');
        expect(row.gruppenmail).toBe('demo-sj26-27-jgb-ges-pag');
    });

    it('wandelt Graph-Gruppe in Belegungszeile', () => {
        const row = graphGroupToBelegungRow(
            {
                id: 'gid-1',
                displayName: 'SJ26-27 | 1AK | BESPK-RINNE',
                mailNickname: 'SJ26-27-jg2031-ak-BESPK-RINNE'
            },
            { yearPrefix: 'SJ26-27' }
        );
        expect(row).not.toBeNull();
        expect(row.graphGroupId).toBe('gid-1');
        expect(row.klasse).toBe('1AK');
        expect(row.fach).toBe('BESPK');
        expect(row.lehrerCode).toBe('RINNE');
        expect(row.gruppenmail).toBe('sj26-27-jg2031-ak-bespk-rinne');
    });

    it('merged Import in bestehende Belegung per gruppenmail', () => {
        const { snapshot, stats } = mergeBelegungWithGraphImport(
            {
                yearPrefix: 'SJ26-27',
                rows: [{ klasse: '1A', fach: 'BW', gruppenmail: 'sj26-27-jg2031-hakb-bw-frech', teamName: 'x' }]
            },
            [
                {
                    klasse: '1A',
                    fach: 'BW',
                    lehrerCode: 'FRECH',
                    gruppenmail: 'sj26-27-jg2031-hakb-bw-frech',
                    graphGroupId: 'g-new',
                    teamName: 'SJ26-27 | 1A | BW-FRECH'
                }
            ],
            { yearPrefix: 'SJ26-27' }
        );
        expect(stats.updated).toBe(1);
        expect(snapshot.rows[0].graphGroupId).toBe('g-new');
    });
});
