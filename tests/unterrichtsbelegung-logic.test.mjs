import { describe, it, expect } from 'vitest';
import {
    buildRowsFromTeamsData,
    buildSnapshotFromTeamsData,
    teachersByKlasse,
    normalizeBelegungSnapshot,
    summarizeBelegung
} from '../src/shared/unterrichtsbelegung-logic.js';

describe('unterrichtsbelegung-logic', () => {
    const sampleTeams = [
        {
            teamName: 'SJ26 | 1A | D',
            gruppenmail: 'sj26-1a-d@school.at',
            besitzer: 'mueller@school.at',
            isValid: true,
            originalClass: '1A',
            fach: 'D',
            gruppe: '',
            lehrerCode: 'MUE'
        },
        {
            teamName: 'SJ26 | 1A | M',
            gruppenmail: 'sj26-1a-m@school.at',
            besitzer: 'schmidt@school.at',
            isValid: true,
            originalClass: '1A',
            fach: 'M',
            gruppe: '',
            lehrerCode: 'SCH'
        },
        {
            teamName: 'SJ26 | 2B | D',
            gruppenmail: 'sj26-2b-d@school.at',
            besitzer: 'mueller@school.at',
            isValid: true,
            originalClass: '2B',
            fach: 'D',
            gruppe: '',
            lehrerCode: 'MUE'
        },
        {
            teamName: 'draft',
            gruppenmail: '',
            besitzer: '',
            isValid: false,
            ktManualDraft: true,
            originalClass: '9Z',
            fach: 'X',
            lehrerCode: 'XX'
        }
    ];

    it('baut Zeilen aus gültigen Teams und ignoriert Drafts', () => {
        const rows = buildRowsFromTeamsData(sampleTeams);
        expect(rows).toHaveLength(3);
        expect(rows.map((r) => r.klasse)).toEqual(['1A', '1A', '2B']);
        expect(rows[0].lehrerEmail).toBe('mueller@school.at');
    });

    it('aggregiert Lehrkräfte pro Klasse', () => {
        const rows = buildRowsFromTeamsData(sampleTeams);
        const by = teachersByKlasse(rows);
        expect(by.get('1A')).toHaveLength(2);
        expect(by.get('2B')).toHaveLength(1);
        expect(by.get('1A').map((t) => t.code).sort()).toEqual(['MUE', 'SCH']);
    });

    it('baut Snapshot mit Metadaten', () => {
        const snap = buildSnapshotFromTeamsData(sampleTeams, { yearPrefix: 'SJ26' });
        expect(snap).not.toBeNull();
        expect(snap.yearPrefix).toBe('SJ26');
        expect(snap.source).toBe('kursteams');
        expect(snap.classCount).toBe(2);
        expect(snap.teacherCount).toBe(2);
        expect(snap.rows).toHaveLength(3);
        const sum = summarizeBelegung(snap);
        expect(sum.classes).toBe(2);
        expect(sum.rows).toBe(3);
    });

    it('normalizeBelegungSnapshot verdichtet und zählt', () => {
        const n = normalizeBelegungSnapshot({
            yearPrefix: 'SJ25',
            source: 'kursteams',
            rows: [
                { klasse: '1A', lehrerCode: 'A', lehrerEmail: 'a@x.at', fach: 'D' },
                { klasse: '1A', lehrerCode: 'A', lehrerEmail: 'a@x.at', fach: 'D' }
            ]
        });
        expect(n.rows).toHaveLength(1);
        expect(n.classCount).toBe(1);
        expect(n.teacherCount).toBe(1);
    });

    it('leere Teams → null Snapshot', () => {
        expect(buildSnapshotFromTeamsData([])).toBeNull();
        expect(buildSnapshotFromTeamsData([{ isValid: false }])).toBeNull();
    });
});
