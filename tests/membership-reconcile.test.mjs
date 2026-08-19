import { describe, expect, it } from 'vitest';
import {
    applyAdminImportSelection,
    buildMembershipImportPreview,
    buildStudentImportRow,
    buildTeacherImportRow,
    diffClassMemberships,
    diffMemberships,
    indexGraphMembersByEmail,
    memberEmailFromGraph,
    suggestTeacherCode
} from '../src/shared/membership-reconcile.js';

describe('membership-reconcile', () => {
    it('diffMemberships trennt lokal, online und gemeinsam', () => {
        const d = diffMemberships(['ada@s.at', 'neu@s.at'], ['ada@s.at', 'alt@s.at']);
        expect(d.onlyLocal).toEqual(['neu@s.at']);
        expect(d.onlyGraph).toEqual(['alt@s.at']);
        expect(d.both).toEqual(['ada@s.at']);
    });

    it('diffClassMemberships trennt Klassenschüler, fremde Schüler und Lehrkräfte', () => {
        const d = diffClassMemberships(
            ['ada@s.at', 'neu@s.at'],
            ['ada@s.at', 'neu@s.at', 'bob@s.at'],
            ['ada@s.at', 'bob@s.at', 'lehrer@s.at', 'extra@s.at']
        );
        expect(d.onlyLocal).toEqual(['neu@s.at']);
        expect(d.onlyGraph).toEqual(['bob@s.at']);
        expect(d.both).toEqual(['ada@s.at']);
        expect(d.preserved).toEqual(['extra@s.at', 'lehrer@s.at']);
    });

    it('applyAdminImportSelection legt neue Verwaltungskontakte an', () => {
        const r = applyAdminImportSelection([], [
            { selected: true, email: 'sek@s.at', name: 'Sek', role: 'Sekretariat', graphUserId: 'u1' }
        ]);
        expect(r.added).toHaveLength(1);
        expect(r.admin[0]).toEqual({ role: 'Sekretariat', name: 'Sek', email: 'sek@s.at' });
    });

    it('memberEmailFromGraph bevorzugt mail vor upn', () => {
        expect(memberEmailFromGraph({ mail: 'a@s.at', userPrincipalName: 'b@s.at' })).toBe('a@s.at');
        expect(memberEmailFromGraph({ userPrincipalName: 'b@s.at' })).toBe('b@s.at');
    });

    it('suggestTeacherCode vermeidet Kollisionen', () => {
        expect(suggestTeacherCode('max.mustermann@s.at', ['MAXMUSTERMAN'])).toBe('MAXMUSTERMA2');
    });

    it('buildTeacherImportRow nutzt Graph-Anzeigename', () => {
        const row = buildTeacherImportRow(
            { displayName: 'Max M.', mail: 'max.mustermann@s.at' },
            [{ code: 'KV' }]
        );
        expect(row.name).toBe('Max M.');
        expect(row.email).toBe('max.mustermann@s.at');
        expect(row.code).toBe('MAXMUSTERMAN');
    });

    it('buildStudentImportRow setzt Klasse', () => {
        const row = buildStudentImportRow({ displayName: 'Ada', mail: 'ada@s.at' }, '3A');
        expect(row).toEqual({ klasse: '3A', name: 'Ada', email: 'ada@s.at' });
    });

    it('indexGraphMembersByEmail dedupliziert nach E-Mail', () => {
        const map = indexGraphMembersByEmail([
            { mail: 'a@s.at', displayName: 'A' },
            { userPrincipalName: 'a@s.at', displayName: 'B' }
        ]);
        expect(map.size).toBe(1);
        expect(map.get('a@s.at').displayName).toBe('A');
    });

    it('buildMembershipImportPreview markiert fehlende Education-Lizenz', () => {
        const facultySku = '94763226-9b3c-4e75-a931-5c89701abe66';
        const users = [
            {
                id: 'u1',
                displayName: 'Lehrer Liz',
                mail: 'liz@s.at',
                userPrincipalName: 'liz@s.at',
                accountEnabled: true,
                userType: 'Member',
                assignedLicenses: [{ skuId: facultySku }]
            },
            {
                id: 'u2',
                displayName: 'Gast Ohne',
                mail: 'gast@s.at',
                userPrincipalName: 'gast@s.at',
                accountEnabled: true,
                userType: 'Member',
                assignedLicenses: []
            }
        ];
        const rows = buildMembershipImportPreview('lehrer', users, [], null);
        expect(rows).toHaveLength(2);
        const liz = rows.find((r) => r.email === 'liz@s.at');
        const gast = rows.find((r) => r.email === 'gast@s.at');
        expect(liz.selected).toBe(true);
        expect(liz.licenseWarning).toBeFalsy();
        expect(gast.licenseWarning).toBe(true);
        expect(gast.selected).toBe(false);
    });
});
