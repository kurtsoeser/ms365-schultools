import { describe, it, expect } from 'vitest';
import {
    computeSetupGlance,
    effectiveFreistellungFlowAccount
} from '../src/tools/power-automate/freistellung-setup-status.js';

describe('freistellung-setup-status', () => {
    it('computeSetupGlance reflects list, emails and flow flag', () => {
        const g = computeSetupGlance(
            {
                siteUrl: 'https://schule.sharepoint.com/sites/admin',
                listId: 'abc',
                emailDirektion: 'dir@schule.at',
                emailSonder: 'son@schule.at',
                emailMailbox: 'auto@schule.at'
            },
            { flowImported: true, onboardingDone: 2, onboardingTotal: 6 }
        );
        expect(g.listOk).toBe(true);
        expect(g.emailsOk).toBe(true);
        expect(g.flowOk).toBe(true);
        expect(g.prepOk).toBe(false);
    });

    it('computeSetupGlance prepOk when onboarding complete', () => {
        const g = computeSetupGlance({}, { onboardingDone: 6, onboardingTotal: 6 });
        expect(g.prepOk).toBe(true);
    });

    it('effectiveFreistellungFlowAccount prefers flowServiceAccount', () => {
        expect(
            effectiveFreistellungFlowAccount({
                flowServiceAccount: 'auto@schule.at',
                emailMailbox: 'mail@schule.at'
            })
        ).toBe('auto@schule.at');
        expect(effectiveFreistellungFlowAccount({ emailMailbox: 'mail@schule.at' })).toBe('mail@schule.at');
    });
});
