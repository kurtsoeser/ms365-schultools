import { describe, it, expect } from 'vitest';
import {
    filterSchularbeiten,
    matchTeacherByEmail,
    matchStudentByEmail,
    viewsForRole,
    resolveStudentKlasseCode
} from '../src/tools/schularbeiten-planer/schularbeiten-planer-state.js';
import { buildIcs } from '../src/tools/schularbeiten-planer/schularbeiten-planer-export.js';

describe('schularbeiten-planer-state filter', () => {
    const items = [
        {
            itemId: '1',
            lehrerCode: 'BAU',
            lehrerEmail: 'bau@schule.at',
            fachCode: 'D',
            klasseCode: '3AK',
            status: 'beantragt',
            datum: '2026-11-10',
            thema: 'A'
        },
        {
            itemId: '2',
            lehrerCode: 'MUE',
            lehrerEmail: 'mue@schule.at',
            fachCode: 'M',
            klasseCode: '3AK',
            status: 'fixiert',
            datum: '2026-11-12',
            thema: 'B'
        },
        {
            itemId: '3',
            lehrerCode: 'MUE',
            lehrerEmail: 'mue@schule.at',
            fachCode: 'M',
            klasseCode: '5BK',
            status: 'fixiert',
            datum: '2026-11-20',
            thema: 'C'
        }
    ];

    it('Lehrer sieht nur eigene', () => {
        const mine = filterSchularbeiten(
            items,
            {},
            {
                role: 'lehrer',
                teacherMatch: { code: 'BAU', email: 'bau@schule.at', name: 'Bauer' },
                accountEmail: 'bau@schule.at',
                onlyMine: true
            }
        );
        expect(mine).toHaveLength(1);
        expect(mine[0].itemId).toBe('1');
    });

    it('ohne Lehrer-Match zeigt alle (Demo)', () => {
        const all = filterSchularbeiten(
            items,
            {},
            {
                role: 'lehrer',
                teacherMatch: null,
                accountEmail: '',
                onlyMine: true
            }
        );
        expect(all).toHaveLength(3);
    });

    it('Admin sieht alle, Filter Status', () => {
        const all = filterSchularbeiten(
            items,
            { status: 'fixiert' },
            {
                role: 'admin',
                teacherMatch: null,
                accountEmail: 'admin@schule.at'
            }
        );
        expect(all).toHaveLength(2);
        expect(all.every((x) => x.status === 'fixiert')).toBe(true);
    });

    it('Schüler sieht nur fixierte der eigenen Klasse', () => {
        const mine = filterSchularbeiten(
            items,
            {},
            {
                role: 'schueler',
                studentMatch: { klasse: '3AK', name: 'Max', email: 'max@schule.at' },
                accountEmail: 'max@schule.at'
            }
        );
        expect(mine).toHaveLength(1);
        expect(mine[0].itemId).toBe('2');
    });

    it('Schüler mit Demo-Klasse', () => {
        const mine = filterSchularbeiten(
            items,
            {},
            {
                role: 'schueler',
                studentMatch: null,
                demoKlasseCode: '5BK',
                accountEmail: ''
            }
        );
        expect(mine).toHaveLength(1);
        expect(mine[0].klasseCode).toBe('5BK');
    });

    it('matchTeacherByEmail', () => {
        expect(
            matchTeacherByEmail('BAU@schule.at', [{ code: 'BAU', email: 'bau@schule.at', name: 'B' }])
        ).toMatchObject({ code: 'BAU' });
    });

    it('matchStudentByEmail', () => {
        expect(
            matchStudentByEmail('Max@schule.at', [{ klasse: '3AK', name: 'Max', email: 'max@schule.at' }])
        ).toMatchObject({ klasse: '3AK' });
    });

    it('viewsForRole blendet Staff für Schüler aus', () => {
        const ids = viewsForRole('schueler').map((v) => v.id);
        expect(ids).toContain('dashboard');
        expect(ids).toContain('kalender');
        expect(ids).toContain('export');
        expect(ids).not.toContain('neu');
        expect(ids).not.toContain('admin');
        expect(ids).not.toContain('meine');
        expect(ids).not.toContain('regeln');
    });

    it('resolveStudentKlasseCode priorisiert Match', () => {
        expect(
            resolveStudentKlasseCode({
                studentMatch: { klasse: '3AK' },
                demoKlasseCode: '5BK'
            })
        ).toBe('3AK');
        expect(resolveStudentKlasseCode({ studentMatch: null, demoKlasseCode: '5BK' })).toBe('5BK');
    });
});

describe('buildIcs', () => {
    it('erzeugt VEVENT', () => {
        const ics = buildIcs(
            [
                {
                    schularbeitId: 'sa-1',
                    datum: '2026-11-10',
                    fachCode: 'D',
                    klasseCode: '3AK',
                    thema: 'Erörterung',
                    status: 'fixiert',
                    dauerMinuten: 100
                }
            ],
            { subjects: [{ code: 'D', name: 'Deutsch' }], classes: [{ code: '3AK', name: '3AK' }], teachers: [] }
        );
        expect(ics).toContain('BEGIN:VCALENDAR');
        expect(ics).toContain('BEGIN:VEVENT');
        expect(ics).toContain('DTSTART;VALUE=DATE:20261110');
        expect(ics).toContain('Deutsch');
    });
});
