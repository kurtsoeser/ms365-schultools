import { createRequire } from 'node:module';
import { describe, expect, it } from 'vitest';

const require = createRequire(import.meta.url);
const { validateTeamsPayload } = require('../src/lib/http-utils.js');
const { issuerMatchesTenant, tokenHasKursteamsScope } = require('../src/lib/validate-token.js');
const { hasOperatorRole, jobVisibleToCaller } = require('../src/lib/require-kursteam-caller.js');
const { DEFAULT_OPERATOR_ROLE_TEMPLATE_IDS } = require('../src/lib/config.js');

const TID = '1fd37d8a-2972-44d2-afcb-81ae47a5bc98';
const OID = 'aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee';

describe('validateTeamsPayload', () => {
    it('verlangt teams und ignoriert tenantId aus dem Body', () => {
        const result = validateTeamsPayload({
            tenantId: 'fremd-mandant',
            mailDomain: '@schule.example',
            teams: [{ teamName: '1A', gruppenmail: '1a', besitzer: 'a@schule.example' }]
        });
        expect(result.error).toBeUndefined();
        expect(result.tenantId).toBeUndefined();
        expect(result.mailDomain).toBe('schule.example');
        expect(result.teams).toHaveLength(1);
    });
});

describe('Kursteams-Token', () => {
    it('erkennt den Scope Kursteams.Create', () => {
        expect(tokenHasKursteamsScope({ scp: 'Kursteams.Create' })).toBe(true);
        expect(tokenHasKursteamsScope({ scp: 'User.Read' })).toBe(false);
    });

    it('bindet den Issuer an die Tenant-ID', () => {
        expect(
            issuerMatchesTenant('https://login.microsoftonline.com/' + TID + '/v2.0', TID)
        ).toBe(true);
        expect(
            issuerMatchesTenant('https://login.microsoftonline.com/common/v2.0', TID)
        ).toBe(false);
    });
});

describe('Wer anlegen darf', () => {
    it('erlaubt Global Admin und Teams-Administrator', () => {
        expect(
            hasOperatorRole(
                [{ roleTemplateId: DEFAULT_OPERATOR_ROLE_TEMPLATE_IDS[1] }],
                DEFAULT_OPERATOR_ROLE_TEMPLATE_IDS
            )
        ).toBe(true);
        expect(
            hasOperatorRole([{ roleTemplateId: '00000000-0000-0000-0000-000000000000' }], DEFAULT_OPERATOR_ROLE_TEMPLATE_IDS)
        ).toBe(false);
    });

    it('zeigt einen Job nur dem Konto, das ihn angelegt hat', () => {
        const job = { tenantId: TID, createdByOid: OID };
        expect(jobVisibleToCaller(job, { tid: TID, oid: OID })).toBe(true);
        expect(jobVisibleToCaller(job, { tid: TID, oid: 'bbbbbbbb-bbbb-cccc-dddd-eeeeeeeeeeee' })).toBe(
            false
        );
        expect(jobVisibleToCaller({ tenantId: TID }, { tid: TID, oid: OID })).toBe(false);
    });
});
