import { describe, expect, it } from 'vitest';
import {
    applyPlaceKindsToRows,
    buildPlaceEmailKindMap,
    classifyDirectorySharedMailboxCandidate,
    countMailboxKinds,
    filterRowsByMailboxKind,
    inferMailboxKind,
    isSharedRecipientType,
    mapDirectoryUserToMailboxRow,
    mergeReportWithDirectory,
    parseCsvLine,
    parseMailboxUsageCsv,
    sharedMailboxesFromUsageJson,
    sharedMailboxesFromUsageReport
} from '../src/tools/postfaecher/postfaecher-load.js';

describe('postfaecher-load CSV', () => {
    it('parst CSV-Zeilen mit Anführungszeichen', () => {
        expect(parseCsvLine('a,"b,c",d')).toEqual(['a', 'b,c', 'd']);
        expect(parseCsvLine('"a""b",c')).toEqual(['a"b', 'c']);
    });

    it('erkennt Shared-Recipient-Typen', () => {
        expect(isSharedRecipientType('Shared')).toBe(true);
        expect(isSharedRecipientType('SharedMailbox')).toBe(true);
        expect(isSharedRecipientType(' shared mailbox ')).toBe(true);
        expect(isSharedRecipientType('User')).toBe(false);
        expect(isSharedRecipientType('UserMailbox')).toBe(false);
    });

    it('filtert Shared und überspringt gelöschte aus dem Usage-Report', () => {
        const csv = [
            'Report Refresh Date,User Principal Name,Display Name,Is Deleted,Recipient Type',
            '2026-01-01,sek@schule.at,Sekretariat,False,Shared',
            '2026-01-01,alt@schule.at,Alt,True,Shared',
            '2026-01-01,lehrer@schule.at,Lehrer,False,User'
        ].join('\n');
        const result = sharedMailboxesFromUsageReport(csv);
        expect(result.ok).toBe(true);
        expect(result.rows).toHaveLength(1);
        expect(result.rows[0].upn).toBe('sek@schule.at');
        expect(result.rows[0].highConfidence).toBe(true);
        expect(result.rows[0].source).toBe('report');
    });

    it('meldet fehlende Recipient-Type-Spalte', () => {
        const csv = [
            'User Principal Name,Display Name,Is Deleted',
            'sek@schule.at,Sekretariat,False'
        ].join('\n');
        const result = sharedMailboxesFromUsageReport(csv);
        expect(result.ok).toBe(false);
        expect(result.reason).toBe('missing-recipient-type');
    });

    it('entfernt BOM und mappt Header-Varianten', () => {
        const csv =
            '\uFEFFUser Principal Name,Display Name,Is Deleted,RecipientType\n' +
            'office@schule.at,Office,false,SharedMailbox\n';
        const parsed = parseMailboxUsageCsv(csv);
        expect(parsed.hasRecipientType).toBe(true);
        expect(parsed.rows[0].upn).toBe('office@schule.at');
        expect(parsed.rows[0].recipientType).toBe('SharedMailbox');
    });
});

describe('postfaecher-load JSON-Report', () => {
    it('filtert Shared aus JSON-Payload', () => {
        const result = sharedMailboxesFromUsageJson({
            value: [
                { userPrincipalName: 'sek@schule.at', displayName: 'Sek', recipientType: 'Shared', isDeleted: false },
                { userPrincipalName: 'u@schule.at', displayName: 'User', recipientType: 'User', isDeleted: false }
            ]
        });
        expect(result.ok).toBe(true);
        expect(result.rows).toHaveLength(1);
        expect(result.rows[0].upn).toBe('sek@schule.at');
    });

    it('lehnt JSON ohne Recipient Type ab', () => {
        const result = sharedMailboxesFromUsageJson({
            value: [{ userPrincipalName: 'sek@schule.at', displayName: 'Sek', isDeleted: false }]
        });
        expect(result.ok).toBe(false);
        expect(result.reason).toBe('missing-recipient-type');
    });
});

describe('postfaecher-load Typen (Raum/Gerät)', () => {
    it('erkennt Räume und Geräte per Name', () => {
        expect(inferMailboxKind({ name: 'Raum 101', mail: 'r101@schule.at' })).toBe('room');
        expect(inferMailboxKind({ name: 'Beamer Wagen 2', alias: 'beamer2' })).toBe('equipment');
        expect(inferMailboxKind({ name: 'Sekretariat', mail: 'sek@schule.at' })).toBe('shared');
    });

    it('Places überschreibt Heuristik', () => {
        expect(inferMailboxKind({ name: 'Sekretariat', placeKind: 'room' })).toBe('room');
    });

    it('filtert nach Typ', () => {
        const rows = [
            { id: '1', name: 'Sek', kind: 'shared' },
            { id: '2', name: 'R1', kind: 'room' },
            { id: '3', name: 'Beamer', kind: 'equipment' }
        ];
        expect(filterRowsByMailboxKind(rows, 'shared')).toHaveLength(1);
        expect(filterRowsByMailboxKind(rows, 'resources')).toHaveLength(2);
        expect(filterRowsByMailboxKind(rows, 'all')).toHaveLength(3);
    });

    it('wendet Places-E-Mails an', () => {
        const map = buildPlaceEmailKindMap([{ emailAddress: 'r1@schule.at', kind: 'room' }]);
        const out = applyPlaceKindsToRows(
            [{ id: '1', name: 'R1', mail: 'r1@schule.at', upn: 'r1@schule.at', kind: 'shared' }],
            map
        );
        expect(out[0].kind).toBe('room');
    });

    it('zählt Typen', () => {
        expect(countMailboxKinds([{ kind: 'shared' }, { kind: 'room' }, { kind: 'room' }])).toEqual({
            shared: 1,
            room: 2,
            equipment: 0,
            other: 0
        });
    });
});

describe('postfaecher-load Verzeichnis', () => {
    it('nimmt deaktivierte Konten mit Mail als Kandidaten', () => {
        const row = classifyDirectorySharedMailboxCandidate({
            id: '1',
            displayName: 'Sekretariat',
            mail: 'sek@schule.at',
            userPrincipalName: 'sek@schule.at',
            mailNickname: 'sek',
            accountEnabled: false,
            userType: 'Member',
            assignedLicenses: []
        });
        expect(row).not.toBeNull();
        expect(row.highConfidence).toBe(true);
        expect(row.kind).toBe('shared');
    });

    it('nimmt aktivierte Konten ohne Personendaten/Lizenz', () => {
        const row = classifyDirectorySharedMailboxCandidate({
            id: '2',
            displayName: 'Office',
            mail: 'office@schule.at',
            userPrincipalName: 'office@schule.at',
            accountEnabled: true,
            assignedLicenses: []
        });
        expect(row).not.toBeNull();
        expect(row.highConfidence).toBe(false);
    });

    it('lehnt Gäste ab', () => {
        expect(
            classifyDirectorySharedMailboxCandidate({
                id: '3',
                mail: 'g@ext.at',
                userPrincipalName: 'g@ext.at',
                accountEnabled: false,
                userType: 'Guest',
                assignedLicenses: []
            })
        ).toBeNull();
    });

    it('reichert Report-Zeilen mit Verzeichnisdaten an', () => {
        const merged = mergeReportWithDirectory(
            [{ id: '', name: 'Sek', mail: '', upn: 'sek@schule.at', alias: '', highConfidence: true, source: 'report' }],
            [
                {
                    id: 'guid-1',
                    displayName: 'Sekretariat',
                    mail: 'sek@schule.at',
                    userPrincipalName: 'sek@schule.at',
                    mailNickname: 'sek'
                }
            ]
        );
        expect(merged[0].id).toBe('guid-1');
        expect(merged[0].alias).toBe('sek');
        expect(merged[0].name).toBe('Sekretariat');
    });

    it('markiert Heuristik mit Personen-Signalen als unsicher', () => {
        const row = mapDirectoryUserToMailboxRow(
            {
                id: '1',
                displayName: 'Max',
                mail: 'max@schule.at',
                userPrincipalName: 'max@schule.at',
                mailNickname: 'max',
                givenName: 'Max',
                surname: 'Muster',
                accountEnabled: false,
                assignedLicenses: []
            },
            'heuristic'
        );
        expect(row.highConfidence).toBe(false);
    });
});
