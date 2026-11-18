import { describe, expect, it } from 'vitest';
import {
    teacherPickOptions,
    teachersWithSubjectsByKlasse,
    listClassTeachingBoards,
    summarizeClassTeaching,
    teachingTeacherEmailsByKlasse,
    teachingEmailsForClassRow,
    belegungRowsForSpoSync,
    addTeacherToClassBelegung,
    removeTeacherFromClassBelegung,
    listBelegungRowsForKlasse,
    patchBelegungRowAt,
    removeBelegungRowAt,
    belegungSnapshotFromRows
} from '../src/shared/klassenlehrer-stammdaten-logic.js';
import { buildKvByKlasse, buildClassChatPlans } from '../src/tools/klassenchats/klassenchats-logic.js';
import { parseRegisterHash, buildRegisterHash } from '../src/shared/stammdaten-register-mode.js';

describe('klassenlehrer-stammdaten-logic (Unterrichtsbelegung)', () => {
    const classes = [
        { code: '1A', name: '1A', year: '2031', headName: 'KV Anna', headEmail: 'anna@schule.at' },
        { code: '1B', name: '1B', year: '2031', headName: '', headEmail: '' },
        { code: '2A', name: '2A', year: '2030', headName: 'Bert', headEmail: 'bert@schule.at' }
    ];
    const belegungRows = [
        {
            klasse: '1A',
            lehrerCode: 'MUEL',
            lehrerEmail: 'mueller@schule.at',
            lehrerName: 'Müller',
            fach: 'M',
            gruppe: '',
            teamName: '',
            gruppenmail: ''
        },
        {
            klasse: '1A',
            lehrerCode: 'MUEL',
            lehrerEmail: 'mueller@schule.at',
            lehrerName: 'Müller',
            fach: 'PH',
            gruppe: '',
            teamName: '',
            gruppenmail: ''
        },
        {
            klasse: '1A',
            lehrerCode: 'ANNA',
            lehrerEmail: 'anna@schule.at',
            lehrerName: 'Anna',
            fach: 'D',
            gruppe: '',
            teamName: '',
            gruppenmail: ''
        },
        {
            klasse: '2A',
            lehrerCode: 'BERT',
            lehrerEmail: 'bert@schule.at',
            lehrerName: 'Bert',
            fach: 'E',
            gruppe: '',
            teamName: '',
            gruppenmail: ''
        }
    ];

    it('teachersWithSubjectsByKlasse gruppiert Fächer', () => {
        const map = teachersWithSubjectsByKlasse(belegungRows);
        expect(map.get('1A')).toHaveLength(2);
        const mueller = map.get('1A').find((t) => t.email === 'mueller@schule.at');
        expect(mueller.faecher).toEqual(['M', 'PH']);
    });

    it('listClassTeachingBoards filtert und zeigt KV-Hinweis', () => {
        const boards = listClassTeachingBoards(classes, belegungRows, 'all');
        expect(boards.map((b) => b.klasse)).toEqual(['1A', '1B', '2A']);
        expect(boards.find((b) => b.klasse === '1B').teachers).toEqual([]);
        expect(boards.find((b) => b.klasse === '1A').kvEmail).toBe('anna@schule.at');
        expect(listClassTeachingBoards(classes, belegungRows, 'without').map((b) => b.klasse)).toEqual([
            '1B'
        ]);
        expect(
            listClassTeachingBoards(classes, belegungRows, 'all', 'mueller').map((b) => b.klasse)
        ).toEqual(['1A']);
    });

    it('summarizeClassTeaching', () => {
        expect(summarizeClassTeaching(classes, belegungRows)).toEqual({
            classTotal: 3,
            withTeachers: 2,
            withoutTeachers: 1,
            teacherAssignments: 3,
            rowCount: 4
        });
    });

    it('teachingTeacherEmailsByKlasse löst E-Mails und Stammdaten-Kürzel', () => {
        const rows = belegungRows.concat([
            {
                klasse: '1B',
                lehrerCode: 'NEU',
                lehrerEmail: '',
                lehrerName: 'Neu',
                fach: 'INF',
                gruppe: '',
                teamName: '',
                gruppenmail: ''
            }
        ]);
        const map = teachingTeacherEmailsByKlasse(rows, [
            { code: 'NEU', email: 'neu@schule.at' },
            { code: 'MUEL', email: 'ignored@schule.at' }
        ]);
        expect(map.get('1A').slice().sort()).toEqual(['anna@schule.at', 'mueller@schule.at']);
        expect(teachingEmailsForClassRow(map, { code: '1B' })).toEqual(['neu@schule.at']);
        expect(teachingEmailsForClassRow(map, { code: '1A' })).toHaveLength(2);
        expect(teachingEmailsForClassRow(map, { code: '2A' })).toEqual(['bert@schule.at']);
    });

    it('belegungRowsForSpoSync baut Sync-Schlüssel und Title', () => {
        const rows = belegungRowsForSpoSync(belegungRows, [
            { code: 'MUEL', email: 'mueller@schule.at', name: 'Müller' }
        ]);
        expect(rows.length).toBe(4);
        const one = rows.find((r) => r.syncKey === '1a|mueller@schule.at|m');
        expect(one).toBeTruthy();
        expect(one.klasse).toBe('1A');
        expect(one.fach).toBe('M');
        expect(one.title).toContain('1A');
        expect(one.title).toContain('M');
    });

    it('add/removeTeacherToClassBelegung', () => {
        const added = addTeacherToClassBelegung(belegungRows, '1B', {
            email: 'chris@schule.at',
            name: 'Chris',
            code: 'CHRI'
        });
        expect(added).toHaveLength(5);
        expect(
            listClassTeachingBoards(classes, added, 'all').find((b) => b.klasse === '1B').teachers
        ).toHaveLength(1);
        const again = addTeacherToClassBelegung(added, '1B', {
            email: 'chris@schule.at',
            name: 'Chris'
        });
        expect(again).toHaveLength(5);
        const removed = removeTeacherFromClassBelegung(added, '1B', { email: 'chris@schule.at' });
        expect(removed).toHaveLength(4);
    });

    it('belegungSnapshotFromRows setzt Meta', () => {
        const snap = belegungSnapshotFromRows(belegungRows, { yearPrefix: '25', source: 'kursteams' });
        expect(snap.yearPrefix).toBe('25');
        expect(snap.classCount).toBe(2);
        expect(snap.teacherCount).toBe(3);
        expect(snap.rows).toHaveLength(4);
        expect(snap.source).toBe('kursteams');
    });

    it('listBelegungRowsForKlasse und patch/remove', () => {
        const listed = listBelegungRowsForKlasse(belegungRows, '1A');
        expect(listed).toHaveLength(3);
        const idx = listed[0].index;
        const patched = patchBelegungRowAt(belegungRows, idx, { fach: 'BIO' });
        expect(patched[idx].fach).toBe('BIO');
        expect(removeBelegungRowAt(patched, idx)).toHaveLength(3);
    });

    it('teacherPickOptions', () => {
        expect(
            teacherPickOptions([
                { name: 'Bert', email: 'bert@schule.at', code: 'B' },
                { name: 'Anna', email: 'anna@schule.at' },
                { name: 'X', email: 'bert@schule.at' }
            ]).map((o) => o.email)
        ).toEqual(['anna@schule.at', 'bert@schule.at']);
    });

    it('Klassenchats nutzen Belegung + KV getrennt', () => {
        const kv = buildKvByKlasse(classes);
        expect(kv.get('1A')?.email).toBe('anna@schule.at');
        const plans = buildClassChatPlans(
            { rows: belegungRows },
            kv,
            { knownClassCodes: classes }
        );
        const plan = plans.find((p) => String(p.klasse || '').toUpperCase() === '1A');
        expect(plan).toBeTruthy();
        const emails = (plan.members || []).map((m) => String(m.email || '').toLowerCase());
        expect(emails).toContain('mueller@schule.at');
        expect(emails).toContain('anna@schule.at');
        const anna = plan.members.find((m) => m.email === 'anna@schule.at');
        expect(anna?.role).toBe('kv');
    });
});

describe('lehrer-je-klasse tab hash', () => {
    it('mappt #klassenlehrer', () => {
        expect(parseRegisterHash('#klassenlehrer').tabBtnId).toBe('tabMainKlassenlehrer');
        expect(buildRegisterHash('tabMainKlassenlehrer')).toBe('klassenlehrer');
    });
});
