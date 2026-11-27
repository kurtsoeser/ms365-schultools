import { describe, expect, it } from 'vitest';
import {
    evaluateLehrerPortalReadiness,
    isLehrerPortalReady,
    missingLabels,
    schoolItAdminMailto
} from '../src/shared/lehrer-portal-readiness.js';

describe('lehrer-portal-readiness', () => {
    it('baut mailto an die Schul-IT', () => {
        expect(schoolItAdminMailto({ itM365AdminEmail: 'it@schule.at', itM365Admin: 'Ada' })).toMatch(
            /^mailto:it@schule\.at\?/
        );
        expect(schoolItAdminMailto({ itM365AdminEmail: '' })).toBe('');
        expect(schoolItAdminMailto({ itM365AdminEmail: 'keine-mail' })).toBe('');
    });

    it('denied wenn Lizenz nicht erlaubt', () => {
        const r = evaluateLehrerPortalReadiness({ allowed: false });
        expect(r.status).toBe('denied');
        expect(r.missing).toContain('license');
        expect(r.title).toMatch(/freigeschaltet/i);
        expect(r.message).toMatch(/IT-Admin|Anfrage/i);
        expect(isLehrerPortalReady(r)).toBe(false);
    });

    it('ohne Lizenzentscheidung kein Setup-Banner (ready)', () => {
        expect(evaluateLehrerPortalReadiness({}).status).toBe('ready');
        expect(evaluateLehrerPortalReadiness({ allowed: null }).status).toBe('ready');
        expect(
            evaluateLehrerPortalReadiness({
                allowed: undefined,
                teacherCount: 0,
                classCount: 0
            }).status
        ).toBe('ready');
    });

    it('setup_incomplete bei fehlender Tool-Site und leeren Listen', () => {
        const r = evaluateLehrerPortalReadiness({
            allowed: true,
            toolSiteUrl: '',
            teacherCount: 0,
            classCount: 0,
            itLibraryReady: false
        });
        expect(r.status).toBe('setup_incomplete');
        expect(r.missing).toEqual(['toolSiteUrl', 'teachers', 'classes', 'itLibrary']);
        expect(r.message).toMatch(/Tool-Seite|Listen|IT-Sicherung/i);
        expect(isLehrerPortalReady(r)).toBe(false);
    });

    it('ready wenn Site und Stammdaten da sind', () => {
        const r = evaluateLehrerPortalReadiness({
            allowed: true,
            toolSiteUrl: 'https://schule.sharepoint.com/sites/MS365',
            teacherCount: 12,
            classCount: 8,
            itLibraryReady: true
        });
        expect(r.status).toBe('ready');
        expect(r.missing).toEqual([]);
        expect(isLehrerPortalReady(r)).toBe(true);
    });

    it('IT-Sicherung nur bei leeren Stammdaten mitzählen', () => {
        const withData = evaluateLehrerPortalReadiness({
            allowed: true,
            toolSiteUrl: 'https://schule.sharepoint.com/sites/x',
            teacherCount: 3,
            classCount: 2,
            itLibraryReady: false
        });
        expect(withData.status).toBe('ready');

        const empty = evaluateLehrerPortalReadiness({
            allowed: true,
            toolSiteUrl: 'https://schule.sharepoint.com/sites/x',
            teacherCount: 0,
            classCount: 0,
            itLibraryReady: false
        });
        expect(empty.status).toBe('setup_incomplete');
        expect(empty.missing).toEqual(['teachers', 'classes', 'itLibrary']);
    });

    it('missingLabels übersetzt IDs', () => {
        expect(missingLabels(['toolSiteUrl', 'teachers'])).toEqual([
            'Tool-Site-URL (SharePoint der Schul-Apps)',
            'Lehrerliste auf der Tool-Site'
        ]);
    });
});
