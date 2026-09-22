import { createRequire } from 'node:module';
import { describe, expect, it, beforeEach, afterEach } from 'vitest';

const require = createRequire(import.meta.url);

process.env.AZURE_TENANT_ID = process.env.AZURE_TENANT_ID || '11111111-1111-1111-1111-111111111111';
process.env.AZURE_CLIENT_ID = process.env.AZURE_CLIENT_ID || '22222222-2222-2222-2222-222222222222';
process.env.AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || 'test-secret';
process.env.LICENSE_API_BASE_URL = 'https://license.test/api/license';

const {
    licenseTokenFromRequest,
    assertTenantLicenseAllowed
} = require('../src/lib/require-tenant-license.js');

const TID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';

function mockRequest(headers) {
    const map = new Map(
        Object.entries(headers || {}).map(([k, v]) => [String(k).toLowerCase(), v])
    );
    return {
        headers: {
            get(name) {
                return map.get(String(name).toLowerCase()) || null;
            }
        }
    };
}

describe('licenseTokenFromRequest', () => {
    it('liest den Bearer aus X-MS365-License-Authorization', () => {
        const req = mockRequest({
            'X-MS365-License-Authorization': 'Bearer lic-token-1'
        });
        expect(licenseTokenFromRequest(req)).toBe('lic-token-1');
    });

    it('liefert leer ohne Header', () => {
        expect(licenseTokenFromRequest(mockRequest({}))).toBe('');
    });
});

describe('assertTenantLicenseAllowed', () => {
    const originalFetch = global.fetch;

    beforeEach(() => {
        global.fetch = async () => {
            throw new Error('fetch nicht gemockt');
        };
    });

    afterEach(() => {
        global.fetch = originalFetch;
    });

    it('lehnt fehlendes Lizenz-Token ab', async () => {
        await expect(assertTenantLicenseAllowed({ tid: TID }, '')).rejects.toMatchObject({
            status: 401
        });
    });

    it('erlaubt freigeschalteten Mandanten mit passender tid', async () => {
        global.fetch = async (url, opts) => {
            expect(String(url)).toBe('https://license.test/api/license/me');
            expect(opts.headers.Authorization).toBe('Bearer good');
            return {
                ok: true,
                status: 200,
                async text() {
                    return JSON.stringify({ allowed: true, tenantId: TID, message: 'ok' });
                }
            };
        };
        await expect(assertTenantLicenseAllowed({ tid: TID }, 'good')).resolves.toBeUndefined();
    });

    it('lehnt ab, wenn allowed false', async () => {
        global.fetch = async () => ({
            ok: true,
            status: 200,
            async text() {
                return JSON.stringify({
                    allowed: false,
                    tenantId: TID,
                    message: 'Nicht freigeschaltet.'
                });
            }
        });
        await expect(assertTenantLicenseAllowed({ tid: TID }, 'tok')).rejects.toMatchObject({
            status: 403,
            message: 'Nicht freigeschaltet.'
        });
    });

    it('lehnt tid-Mismatch ab', async () => {
        global.fetch = async () => ({
            ok: true,
            status: 200,
            async text() {
                return JSON.stringify({
                    allowed: true,
                    tenantId: '00000000-0000-0000-0000-000000000000'
                });
            }
        });
        await expect(assertTenantLicenseAllowed({ tid: TID }, 'tok')).rejects.toMatchObject({
            status: 403
        });
    });
});
