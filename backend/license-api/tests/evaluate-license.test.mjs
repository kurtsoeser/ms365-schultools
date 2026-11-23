import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const {
    evaluateLicense,
    findFieldsForTenant,
    findLicenseItemForTenant,
    parseDomainList,
    canReportToolSiteUrl,
    normalizeReportedToolSiteUrl
} = require('../src/lib/evaluate-license.js');
const { issuerMatchesTenant } = require('../src/lib/validate-token.js');

describe('evaluateLicense', () => {
    const allowed = ['trial', 'active'];
    const now = new Date('2026-06-15T12:00:00Z');

    it('lehnt fehlende Listenzeile ab', () => {
        const r = evaluateLicense({
            tenantId: 'aaa',
            fields: null,
            allowedStatuses: allowed,
            now
        });
        expect(r.allowed).toBe(false);
        expect(r.reason).toBe('not_registered');
    });

    it('erlaubt active ohne Ablaufdatum', () => {
        const r = evaluateLicense({
            tenantId: 'aaa',
            fields: { Title: 'Testschule', TenantId: 'aaa', Status: 'active' },
            allowedStatuses: allowed,
            now
        });
        expect(r.allowed).toBe(true);
        expect(r.schoolName).toBe('Testschule');
        expect(r.reason).toBe('ok');
    });

    it('lehnt abgelaufenes ValidUntil ab', () => {
        const r = evaluateLicense({
            tenantId: 'aaa',
            fields: {
                Title: 'Alt',
                Status: 'active',
                ValidUntil: '2026-01-01'
            },
            allowedStatuses: allowed,
            now
        });
        expect(r.allowed).toBe(false);
        expect(r.reason).toBe('expired');
    });

    it('erlaubt ValidUntil am gleichen Tag', () => {
        const r = evaluateLicense({
            tenantId: 'aaa',
            fields: { Title: 'Ok', Status: 'trial', ValidUntil: '2026-06-15' },
            allowedStatuses: allowed,
            now
        });
        expect(r.allowed).toBe(true);
    });

    it('lehnt blocked ab', () => {
        const r = evaluateLicense({
            tenantId: 'aaa',
            fields: { Status: 'blocked', Title: 'X' },
            allowedStatuses: allowed,
            now
        });
        expect(r.allowed).toBe(false);
        expect(r.reason).toBe('blocked');
    });
});

describe('findFieldsForTenant', () => {
    it('findet TenantId case-insensitive', () => {
        const fields = findFieldsForTenant(
            [
                { fields: { TenantId: 'AAA-BBB', Title: 'A' } },
                { fields: { TenantId: 'ccc', Title: 'B' } }
            ],
            'aaa-bbb'
        );
        expect(fields.Title).toBe('A');
    });
});

describe('issuerMatchesTenant', () => {
    it('akzeptiert v2 issuer', () => {
        expect(
            issuerMatchesTenant(
                'https://login.microsoftonline.com/1fd37d8a-2972-44d2-afcb-81ae47a5bc98/v2.0',
                '1fd37d8a-2972-44d2-afcb-81ae47a5bc98'
            )
        ).toBe(true);
    });

    it('lehnt falschen Tenant ab', () => {
        expect(
            issuerMatchesTenant(
                'https://login.microsoftonline.com/aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa/v2.0',
                '1fd37d8a-2972-44d2-afcb-81ae47a5bc98'
            )
        ).toBe(false);
    });
});

describe('parseDomainList', () => {
    it('sammelt Primary + Additional und normalisiert URLs', () => {
        expect(
            parseDomainList(
                'kurtsoeser.at',
                'https://contoso.sharepoint.com/sites/x\nmodeebensee.at; hak-steyr.at'
            )
        ).toEqual(['kurtsoeser.at', 'contoso.sharepoint.com', 'modeebensee.at', 'hak-steyr.at']);
    });
});

describe('parseContactEmails', () => {
    const { parseContactEmails, upnMatchesDesignatedItContact } = require('../src/lib/evaluate-license.js');

    it('parst mehrere IT-Kontakte', () => {
        expect(parseContactEmails('a@schule.at; B@schule.at\nb@schule.at', null)).toEqual([
            'a@schule.at',
            'b@schule.at'
        ]);
    });

    it('erkennt designierten IT-Kontakt per UPN', () => {
        const emails = parseContactEmails('it@hak-steyr.at, backup@hak-steyr.at', '');
        expect(upnMatchesDesignatedItContact('IT@hak-steyr.at', emails)).toBe(true);
        expect(upnMatchesDesignatedItContact('lehrer@hak-steyr.at', emails)).toBe(false);
    });
});

describe('Tool-Site melden (Schul-IT → Betreiber)', () => {
    it('findLicenseItemForTenant liefert id + fields', () => {
        const hit = findLicenseItemForTenant(
            [
                {
                    id: '42',
                    fields: { TenantId: 'aaa-bbb', Title: 'Schule', ToolSiteUrl: 'https://x.sharepoint.com/sites/T' }
                }
            ],
            'AAA-BBB'
        );
        expect(hit).toEqual({
            id: '42',
            fields: { TenantId: 'aaa-bbb', Title: 'Schule', ToolSiteUrl: 'https://x.sharepoint.com/sites/T' }
        });
        expect(findFieldsForTenant([{ id: '1', fields: { TenantId: 'other' } }], 'aaa-bbb')).toBeNull();
    });

    it('canReportToolSiteUrl: IT-Kontakt oder Bootstrap ohne Kontakte', () => {
        expect(
            canReportToolSiteUrl({
                allowed: true,
                contactEmails: ['it@schule.at'],
                upn: 'it@schule.at'
            }).ok
        ).toBe(true);
        expect(
            canReportToolSiteUrl({
                allowed: true,
                contactEmails: ['it@schule.at'],
                upn: 'lehrer@schule.at'
            }).ok
        ).toBe(false);
        expect(
            canReportToolSiteUrl({
                allowed: true,
                contactEmails: [],
                upn: 'admin@schule.at'
            }).reason
        ).toBe('no_contacts_bootstrap');
        expect(canReportToolSiteUrl({ allowed: false, contactEmails: [], upn: 'x@y.at' }).ok).toBe(false);
    });

    it('normalizeReportedToolSiteUrl akzeptiert SharePoint-Sites', () => {
        expect(
            normalizeReportedToolSiteUrl('https://schule.sharepoint.com/sites/MS365-Schultools/')
        ).toEqual({
            ok: true,
            url: 'https://schule.sharepoint.com/sites/MS365-Schultools'
        });
        expect(normalizeReportedToolSiteUrl('http://schule.sharepoint.com/sites/X').ok).toBe(false);
        expect(normalizeReportedToolSiteUrl('https://example.com/foo').ok).toBe(false);
    });
});
