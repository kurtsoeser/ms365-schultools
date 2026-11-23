import { describe, expect, it } from 'vitest';
import {
    evaluateItPortalReleaseGate,
    evaluateLehrerPortalWithReleaseGate,
    isPortalReleased,
    readPortalReleasedAt
} from '../src/shared/portal-release-gate.js';

describe('portal-release-gate (b6)', () => {
    it('liest portalReleasedAt aus Setup', () => {
        expect(readPortalReleasedAt({ portalReleasedAt: '2026-11-21T10:00:00.000Z' })).toMatch(
            /^2026-11-21/
        );
        expect(isPortalReleased({ portalReleasedAt: '' })).toBe(false);
        expect(isPortalReleased({ portalReleasedAt: '2026-11-21T10:00:00.000Z' })).toBe(true);
    });

    it('Lehrer sieht setup_incomplete ohne IT-Freigabe', () => {
        const r = evaluateLehrerPortalWithReleaseGate({
            allowed: true,
            toolSiteUrl: 'https://schule.sharepoint.com/sites/x',
            teacherCount: 3,
            classCount: 2,
            itLibraryReady: true,
            portalReleased: false
        });
        expect(r.status).toBe('setup_incomplete');
        expect(r.missing).toContain('portalRelease');
    });

    it('Lehrer ready mit Freigabe', () => {
        const r = evaluateLehrerPortalWithReleaseGate({
            allowed: true,
            toolSiteUrl: 'https://schule.sharepoint.com/sites/x',
            teacherCount: 3,
            classCount: 2,
            portalReleased: true
        });
        expect(r.status).toBe('ready');
    });

    it('IT-Gate canRelease nur bei Voraussetzungen', () => {
        expect(
            evaluateItPortalReleaseGate({
                allowed: true,
                toolSiteUrl: '',
                teacherCount: 0,
                classCount: 0
            }).canRelease
        ).toBe(false);
        expect(
            evaluateItPortalReleaseGate({
                allowed: true,
                toolSiteUrl: 'https://schule.sharepoint.com/sites/x',
                teacherCount: 2,
                classCount: 2,
                itLibraryReady: true
            }).canRelease
        ).toBe(true);
    });
});
