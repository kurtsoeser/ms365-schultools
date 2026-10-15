import { describe, expect, it } from 'vitest';
import {
    SHARED_MAILBOX_CMD_PS_MARKER,
    buildSharedMailboxCreateCmd,
    buildSharedMailboxCreateScript,
    normalizeSharedMailboxDelegates,
    normalizeSharedMailboxRights,
    suggestSharedMailboxIdentity
} from '../src/shared/exchange-shared-mailbox-script.js';

describe('exchange-shared-mailbox-script', () => {
    it('schlägt Alias und SMTP aus Name + Domain vor', () => {
        const sug = suggestSharedMailboxIdentity('Erweiterte Schulleitung', 'dsrom.de');
        expect(sug.alias).toBe('erweiterte-schulleitung');
        expect(sug.primarySmtp).toBe('erweiterte-schulleitung@dsrom.de');
    });

    it('baut geführten New-Mailbox-Assistenten mit SMTP', () => {
        const script = buildSharedMailboxCreateScript({
            displayName: 'Test Postfach',
            alias: 'test-pf',
            primarySmtp: 'test-pf@schule.at',
            description: 'Organigramm',
            suggestedAdminUpn: 'admin@schule.at'
        });
        expect(script).toContain('New-Mailbox');
        expect(script).toContain('Shared = $true');
        expect(script).toContain('test-pf@schule.at');
        expect(script).toContain('Test Postfach');
        expect(script).toContain('Connect-ExchangeOnline');
        expect(script).toContain('Admin-Konto wählen');
        expect(script).toContain('SuggestedAdminUpn');
        expect(script).toContain('admin@schule.at');
        expect(script).toContain('Read-Host');
        expect(script).toContain('Postfach jetzt in diesem Tenant anlegen');
    });

    it('normalisiert Delegierte und Rechte', () => {
        expect(
            normalizeSharedMailboxDelegates([
                'a@schule.at',
                { email: 'B@Schule.at' },
                'a@schule.at',
                'kein-mail',
                { upn: 'c@schule.at' }
            ])
        ).toEqual(['a@schule.at', 'B@Schule.at', 'c@schule.at']);
        expect(normalizeSharedMailboxRights(['FullAccess', 'sendas', 'SendOnBehalf', 'x'])).toEqual([
            'FullAccess',
            'SendAs',
            'SendOnBehalf'
        ]);
    });

    it('setzt FullAccess und SendAs für Organigramm-Personen', () => {
        const script = buildSharedMailboxCreateScript({
            displayName: 'Sekretariat',
            alias: 'sekretariat',
            primarySmtp: 'sekretariat@schule.at',
            delegates: ['dori@schule.at', 'hutterer@schule.at'],
            rights: ['FullAccess', 'SendAs'],
            autoMapping: true
        });
        expect(script).toContain("'dori@schule.at'");
        expect(script).toContain("'hutterer@schule.at'");
        expect(script).toContain('Add-MailboxPermission');
        expect(script).toContain('AccessRights FullAccess');
        expect(script).toContain('AutoMapping:$AutoMapping');
        expect(script).toContain('Add-RecipientPermission');
        expect(script).toContain('AccessRights SendAs');
        expect(script).toContain('Get-MailboxPermission');
    });

    it('ohne Delegierte keine Rechte-Cmdlets', () => {
        const script = buildSharedMailboxCreateScript({
            displayName: 'Leer',
            alias: 'leer',
            primarySmtp: 'leer@schule.at'
        });
        expect(script).not.toContain('Add-MailboxPermission');
        expect(script).toContain('Keine Rechte ausgewählt');
    });

    it('baut doppelklickbare .cmd mit eingebettetem Assistenten', () => {
        const cmd = buildSharedMailboxCreateCmd({
            displayName: 'Sekretariat',
            alias: 'sekretariat',
            primarySmtp: 'sekretariat@modeebensee.at',
            delegates: ['a@modeebensee.at'],
            suggestedAdminUpn: 'admin@modeebensee.at'
        });
        expect(cmd).toContain('@echo off');
        expect(cmd).toContain('MS365_CMD_PATH');
        expect(cmd).toContain(SHARED_MAILBOX_CMD_PS_MARKER);
        expect(cmd).toContain('powershell.exe');
        expect(cmd).toContain('New-Mailbox');
        expect(cmd).toContain('Admin-Konto wählen');
        expect(cmd).toContain('sekretariat@modeebensee.at');
        expect(cmd.indexOf('@echo off')).toBeLessThan(cmd.indexOf(SHARED_MAILBOX_CMD_PS_MARKER));
        expect(cmd.indexOf(SHARED_MAILBOX_CMD_PS_MARKER)).toBeLessThan(cmd.indexOf('New-Mailbox'));
    });
});
