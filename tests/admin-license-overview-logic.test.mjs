import { describe, expect, it } from 'vitest';
import {
    adoptionCounts,
    collectContacts,
    hasContact,
    isDueWithin,
    isExpired,
    licenseCounts,
    prepareMailHref
} from '../src/shared/admin-license-overview-logic.js';

const today = new Date(2026, 10, 22);

describe('admin license overview', () => {
    it('zählt Ablauf, Status und fehlende Kontakte', () => {
        const schools = [
            { status: 'active', validUntil: '2026-12-01', contactEmails: ['it@a.at'] },
            { status: 'active', validUntil: '2026-11-01', contactEmail: 'alt@b.at' },
            { status: 'expired', validUntil: '2027-01-01', contactEmails: ['x@c.at'] },
            { status: 'trial', validUntil: '2026-12-10' },
            { status: 'blocked', validUntil: '2026-12-05', contactEmails: ['b@d.at'] },
            { status: 'active', pendingTenant: true, contactEmails: ['p@e.at'] }
        ];
        const counts = licenseCounts(schools, today, {
            isPending: (school) => !!school.pendingTenant
        });
        expect(counts.active).toBe(3);
        expect(counts.trial).toBe(1);
        expect(counts.due30).toBe(2);
        expect(counts.expired).toBe(2);
        expect(counts.pending).toBe(1);
        expect(counts.nocontact).toBe(1);
        expect(isDueWithin(schools[3], today, 30)).toBe(true);
        expect(isExpired(schools[1], today)).toBe(true);
        expect(hasContact(schools[3])).toBe(false);
    });

    it('liest Mails aus Lizenz und Stammdaten', () => {
        const rows = collectContacts({
            contactEmails: ['IT@Schule.at', 'IT@schule.at'],
            extra: {
                Direktion: 'Direktion: leitung@schule.at',
                DirektionEmail: 'dir@schule.at',
                ItM365Admin: 'Max',
                ItM365AdminEmail: 'it@schule.at',
                WeitereKontakte: 'sek@schule.at und noch text'
            }
        });
        expect(rows.map((row) => row.role + ' ' + row.email)).toEqual([
            'Lizenz it@schule.at',
            'IT it@schule.at',
            'Direktion dir@schule.at',
            'Kontakt sek@schule.at'
        ]);
    });

    it('zählt Pakete und Verkaufsbündel', () => {
        const packages = [
            { id: 'core', label: 'Basis' },
            { id: 'people', label: 'Personen' }
        ];
        const result = adoptionCounts(
            [
                { entitlements: { packages: ['core', 'people'] } },
                { entitlementsDoc: { packages: ['core'] } },
                { entitlements: { packages: [] } }
            ],
            packages,
            (ids) =>
                ids.includes('core') && ids.includes('people')
                    ? { bundle: { id: 'plus', label: 'Plus' }, exact: true }
                    : { bundle: null, exact: false }
        );
        expect(result.withPackages).toBe(2);
        expect(result.byPackage.core).toBe(2);
        expect(result.byPackage.people).toBe(1);
        expect(result.byBundle.plus.count).toBe(1);
        expect(result.custom).toBe(1);
    });

    it('baut einen mailto-Link', () => {
        const href = prepareMailHref(['it@schule.at'], 'Ablauf', 'Guten Tag');
        expect(href.startsWith('mailto:it@schule.at?subject=')).toBe(true);
        expect(decodeURIComponent(href)).toContain('Guten Tag');
        expect(prepareMailHref([], 'x', 'y')).toBe('');
    });
});
