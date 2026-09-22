import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { tokenHasLicenseScope, issuerMatchesTenant } = require('../src/lib/validate-token.js');
const { isOperatorCaller } = require('../src/lib/require-operator.js');
const { tokenAudiences } = require('../src/lib/config.js');

const TID = '1fd37d8a-2972-44d2-afcb-81ae47a5bc98';
const OID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';

describe('License-Token', () => {
    it('erkennt License.Access', () => {
        expect(tokenHasLicenseScope({ scp: 'License.Access' })).toBe(true);
        expect(tokenHasLicenseScope({ scp: 'User.Read' })).toBe(false);
    });

    it('bindet den Issuer an die Tenant-ID', () => {
        expect(
            issuerMatchesTenant('https://login.microsoftonline.com/' + TID + '/v2.0', TID)
        ).toBe(true);
        expect(issuerMatchesTenant('https://login.microsoftonline.com/common/v2.0', TID)).toBe(
            false
        );
    });
});

describe('tokenAudiences', () => {
    it('filtert Graph-Audiences und fällt auf die App-ID zurück', () => {
        const prev = process.env.LICENSE_TOKEN_AUDIENCES;
        process.env.LICENSE_TOKEN_AUDIENCES =
            '00000003-0000-0000-c000-000000000000,https://graph.microsoft.com';
        try {
            expect(tokenAudiences('12e0cfe2-8337-4b35-93e8-542faf658eb3')).toEqual([
                '12e0cfe2-8337-4b35-93e8-542faf658eb3',
                'api://12e0cfe2-8337-4b35-93e8-542faf658eb3'
            ]);
        } finally {
            if (prev === undefined) delete process.env.LICENSE_TOKEN_AUDIENCES;
            else process.env.LICENSE_TOKEN_AUDIENCES = prev;
        }
    });
});

describe('isOperatorCaller', () => {
    it('erlaubt UPN oder OID', () => {
        const cfg = {
            operatorUpns: ['kurt@kurtsoeser.at'],
            operatorOids: [OID]
        };
        expect(isOperatorCaller({ upn: 'kurt@kurtsoeser.at', oid: 'x' }, cfg)).toBe(true);
        expect(isOperatorCaller({ upn: 'other@x.at', oid: OID }, cfg)).toBe(true);
        expect(isOperatorCaller({ upn: 'other@x.at', oid: 'bbbbbbbb-bbbb-cccc-dddd-eeeeeeeeeeee' }, cfg)).toBe(
            false
        );
    });
});
