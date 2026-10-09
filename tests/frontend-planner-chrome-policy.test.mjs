import { describe, it, expect } from 'vitest';
import { isFrontendPlannerPage, showItAppChrome } from '../src/shared/frontend-planner-chrome-policy.js';

describe('frontend-planner-chrome-policy', () => {
    it('isFrontendPlannerPage erkennt die vier Planer', () => {
        expect(isFrontendPlannerPage('/tools/freistellung-planer.html')).toBe(true);
        expect(isFrontendPlannerPage('/x/tools/schularbeiten-planer.html')).toBe(true);
        expect(isFrontendPlannerPage('/tools/projektwochen.html')).toBe(true);
        expect(isFrontendPlannerPage('/tools/lehrer-freistellung-planer.html')).toBe(true);
        expect(isFrontendPlannerPage('/tools/freistellung-setup.html')).toBe(false);
    });

    it('showItAppChrome nur für IT/Global/Designated', () => {
        expect(showItAppChrome(null)).toBe(false);
        expect(showItAppChrome({ loggedIn: true, isIt: false, globalAdmin: false })).toBe(false);
        expect(showItAppChrome({ loggedIn: true, isIt: true })).toBe(true);
        expect(showItAppChrome({ loggedIn: true, globalAdmin: true })).toBe(true);
        expect(showItAppChrome({ loggedIn: true, designatedSchoolIt: true })).toBe(true);
        expect(showItAppChrome({ loggedIn: true, isSchueler: true, isIt: false })).toBe(false);
    });
});
