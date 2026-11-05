import { describe, expect, it, beforeEach, afterEach } from 'vitest';
import {
    CALENDAR_PRESETS,
    buildCalendarFromPreset,
    getCalendarPreset,
    looksLikeCalendarMailbox,
    slugAliasFromName
} from '../src/tools/kalender-postfaecher/kalender-postfaecher-presets.js';
import {
    CALENDAR_POSTFAECHER_CMD_PS_MARKER,
    buildCalendarPostfaecherCmd,
    buildCalendarPostfaecherScript,
    normalizeCalendarPlan,
    normalizeCalendarPlanItem,
    normalizeCalendarPrincipals
} from '../src/tools/kalender-postfaecher/kalender-postfaecher-script.js';
import {
    mergePrincipalMails,
    resolveSchoolDomain
} from '../src/tools/kalender-postfaecher/kalender-postfaecher-ui.js';

describe('kalender-postfaecher presets', () => {
    it('liefert Kernvorlagen', () => {
        expect(CALENDAR_PRESETS.map((p) => p.id)).toEqual(
            expect.arrayContaining(['schulkalender', 'lehrerkalender', 'schularbeiten', 'matura', 'custom'])
        );
        expect(getCalendarPreset('schulkalender')?.publishDefault).toBe(true);
    });

    it('baut Eintrag aus Vorlage mit Domain', () => {
        const item = buildCalendarFromPreset('matura', 'schule.at');
        expect(item.displayName).toBe('Matura-Kalender');
        expect(item.alias).toBe('matura-kalender');
        expect(item.primarySmtp).toBe('matura-kalender@schule.at');
        expect(item.mode).toBe('create');
        expect(item.publish).toBe(false);
    });

    it('erkennt Kalender-Heuristik und Alias-Slug', () => {
        expect(looksLikeCalendarMailbox({ name: 'Schulkalender', upn: 'x@y.at' })).toBe(true);
        expect(looksLikeCalendarMailbox({ name: 'Sekretariat', upn: 'sek@schule.at' })).toBe(false);
        expect(slugAliasFromName('Schularbeiten Kalender')).toBe('schularbeiten-kalender');
    });
});

describe('kalender-postfaecher script', () => {
    it('normalisiert Principals und entfernt Reader-Duplikate bei Editors', () => {
        expect(
            normalizeCalendarPrincipals(['a@schule.at', { email: 'B@Schule.at' }, 'kein-mail'])
        ).toEqual(['a@schule.at', 'B@Schule.at']);
        const item = normalizeCalendarPlanItem({
            displayName: 'Schulkalender',
            alias: 'schulkalender',
            editors: ['redaktion@schule.at'],
            readers: ['redaktion@schule.at', 'alle@schule.at'],
            publish: true
        });
        expect(item.editors).toEqual(['redaktion@schule.at']);
        expect(item.readers).toEqual(['alle@schule.at']);
        expect(item.publish).toBe(true);
    });

    it('baut Batch-Skript mit New-Mailbox, FolderPermission und Publish', () => {
        const plan = normalizeCalendarPlan([
            {
                displayName: 'Schulkalender',
                alias: 'schulkalender',
                primarySmtp: 'schulkalender@schule.at',
                editors: ['sek@schule.at'],
                readers: ['lehrer@schule.at'],
                sendAs: true,
                publish: true
            },
            {
                displayName: 'Matura',
                alias: 'matura-kalender',
                primarySmtp: 'matura@schule.at',
                mode: 'permissionsOnly',
                editors: ['koord@schule.at'],
                publish: false
            }
        ]);
        expect(plan.items).toHaveLength(2);

        const script = buildCalendarPostfaecherScript(plan.items, {
            suggestedAdminUpn: 'admin@schule.at'
        });
        expect(script).toContain('New-Mailbox');
        expect(script).toContain('Shared = $true');
        expect(script).toContain('Set-MailboxFolderPermission');
        expect(script).toContain('Add-MailboxFolderPermission');
        expect(script).toContain('AccessRights Editor');
        expect(script).toContain('AccessRights Reviewer');
        expect(script).toContain('Add-MailboxPermission');
        expect(script).toContain('AccessRights FullAccess');
        expect(script).toContain('Add-RecipientPermission');
        expect(script).toContain('Set-MailboxCalendarFolder');
        expect(script).toContain('PublishEnabled');
        expect(script).toContain('PublishedCalendarUrl');
        expect(script).toContain('schulkalender@schule.at');
        expect(script).toContain('sek@schule.at');
        expect(script).toContain('lehrer@schule.at');
        expect(script).toContain('permissionsOnly');
        expect(script).toContain('admin@schule.at');
        expect(script).toContain('Resolve-CalendarFolderIdentity');
        expect(script).toContain('Kalender');
        expect(script).toContain('Calendar');
    });

    it('baut doppelklickbare .cmd mit Marker', () => {
        const cmd = buildCalendarPostfaecherCmd(
            [
                {
                    displayName: 'Schulkalender',
                    alias: 'schulkalender',
                    primarySmtp: 'schulkalender@schule.at',
                    editors: ['a@schule.at'],
                    publish: true
                }
            ],
            { suggestedAdminUpn: 'admin@schule.at' }
        );
        expect(cmd).toContain('@echo off');
        expect(cmd).toContain(CALENDAR_POSTFAECHER_CMD_PS_MARKER);
        expect(cmd).toContain('New-Mailbox');
        expect(cmd).toContain('Set-MailboxCalendarFolder');
        expect(cmd.indexOf('@echo off')).toBeLessThan(cmd.indexOf(CALENDAR_POSTFAECHER_CMD_PS_MARKER));
        expect(cmd.indexOf(CALENDAR_POSTFAECHER_CMD_PS_MARKER)).toBeLessThan(cmd.indexOf('New-Mailbox'));
    });

    it('ohne gültige Einträge leeres Plan-Array', () => {
        expect(normalizeCalendarPlan([{ displayName: '', alias: '' }]).items).toEqual([]);
    });
});

describe('kalender-postfaecher domain', () => {
    /** @type {Map<string, string>} */
    let store;
    const prevWindow = globalThis.window;
    const prevLs = globalThis.localStorage;

    beforeEach(() => {
        store = new Map();
        globalThis.localStorage = {
            getItem: (k) => (store.has(k) ? store.get(k) : null),
            setItem: (k, v) => {
                store.set(String(k), String(v));
            },
            removeItem: (k) => {
                store.delete(String(k));
            }
        };
        globalThis.window = Object.assign(globalThis.window || globalThis, {
            ms365GetSchoolDomainNoAt: undefined,
            ms365AppDataV2: undefined
        });
        delete window.ms365GetSchoolDomainNoAt;
        delete window.ms365AppDataV2;
    });

    afterEach(() => {
        if (prevLs === undefined) delete globalThis.localStorage;
        else globalThis.localStorage = prevLs;
        if (prevWindow === undefined) delete globalThis.window;
        else globalThis.window = prevWindow;
    });

    it('liest Domain aus school-domain Helper', () => {
        window.ms365GetSchoolDomainNoAt = () => 'meine-schule.at';
        expect(resolveSchoolDomain()).toBe('meine-schule.at');
    });

    it('fällt auf Tenant-Stammdaten zurück', () => {
        localStorage.setItem('ms365-tenant-settings-v1', JSON.stringify({ domain: 'dsrom.de' }));
        expect(resolveSchoolDomain()).toBe('dsrom.de');
    });

    it('merged Picker-Mails ohne Duplikate', () => {
        expect(
            mergePrincipalMails(['a@schule.at'], [
                { mail: 'A@schule.at' },
                { mail: 'b@schule.at' },
                { mail: 'kein-mail' }
            ])
        ).toEqual(['a@schule.at', 'b@schule.at']);
    });
});
