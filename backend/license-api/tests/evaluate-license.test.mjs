import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { evaluateLicense, findFieldsForTenant, parseDomainList } = require('../src/lib/evaluate-license.js');
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
