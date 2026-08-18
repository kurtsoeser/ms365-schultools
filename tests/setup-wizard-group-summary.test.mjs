import { describe, expect, it } from 'vitest';
import {
    SW_MATCH_MEMBER_PREVIEW,
    buildLinkedGroupSummaryHtml,
    entraGroupOverviewUrl,
    formatGroupDateTime,
    personSummaryLabel,
    teamsGroupConversationsUrl,
    visibilityLabel
} from '../src/shared/setup-wizard-group-summary.js';

describe('setup-wizard-group-summary', () => {
    it('personSummaryLabel kombiniert Name und UPN', () => {
        expect(personSummaryLabel({ displayName: 'Anna', userPrincipalName: 'anna@schule.at' })).toBe(
            'Anna (anna@schule.at)'
        );
        expect(personSummaryLabel({ displayName: 'x@y.at', mail: 'x@y.at' })).toBe('x@y.at');
        expect(personSummaryLabel(null)).toBe('');
    });

    it('visibilityLabel übersetzt Graph-Werte', () => {
        expect(visibilityLabel('Private')).toBe('Privat');
        expect(visibilityLabel('Public')).toBe('Öffentlich');
        expect(visibilityLabel('')).toBe('');
    });

    it('formatGroupDateTime formatiert ISO auf de-AT', () => {
        const label = formatGroupDateTime('2026-08-18T10:31:00.000Z');
        expect(label).toMatch(/18\.08\.2026/);
        expect(formatGroupDateTime('')).toBe('');
    });

    it('Entra- und Teams-URLs enthalten die Gruppen-ID', () => {
        const id = '6ab2481b-81d2-44ea-b395-34faacaea7b5';
        expect(entraGroupOverviewUrl(id)).toContain(id);
        expect(teamsGroupConversationsUrl(id)).toContain(id);
        expect(entraGroupOverviewUrl('')).toBe('');
    });

    it('ohne Gruppen-ID erscheint der Leerhinweis', () => {
        const html = buildLinkedGroupSummaryHtml({});
        expect(html).toContain('Noch keine Gruppe gewählt.');
    });

    it('zeigt Name, Alias, ID, Mail, Beschreibung, Owner und Mitglieder', () => {
        const html = buildLinkedGroupSummaryHtml({
            groupId: 'gid-1',
            status: 'ready',
            artLabel: 'Microsoft 365‑Gruppe',
            hasTeam: true,
            group: {
                id: 'gid-1',
                displayName: 'Lehrer:innen',
                mailNickname: 'lehrer',
                mail: 'lehrer@schule.at',
                description: 'Alle Lehrkräfte',
                visibility: 'Private',
                createdDateTime: '2026-01-15T08:00:00.000Z'
            },
            owners: [{ displayName: 'Direktion', mail: 'dir@schule.at' }],
            members: [
                { displayName: 'Max', mail: 'max@schule.at' },
                { displayName: 'Lisa', mail: 'lisa@schule.at' }
            ],
            memberCount: 2
        });
        expect(html).toContain('Lehrer:innen');
        expect(html).toContain('lehrer');
        expect(html).toContain('gid-1');
        expect(html).toContain('lehrer@schule.at');
        expect(html).toContain('Alle Lehrkräfte');
        expect(html).toContain('Privat');
        expect(html).toContain('Besitzer (1)');
        expect(html).toContain('Direktion (dir@schule.at)');
        expect(html).toContain('Mitglieder (2)');
        expect(html).toContain('Max (max@schule.at)');
        expect(html).toContain('Lisa (lisa@schule.at)');
        expect(html).toContain('Team öffnen');
        expect(html).toContain('Entra öffnen');
        expect(html).not.toContain('Angezeigt:');
    });

    it('escaped HTML in Namen', () => {
        const html = buildLinkedGroupSummaryHtml({
            groupId: 'x',
            status: 'partial',
            group: { displayName: '<script>alert(1)</script>' }
        });
        expect(html).not.toContain('<script>');
        expect(html).toContain('&lt;script&gt;');
    });

    it('weist auf gekürzte Mitgliederliste hin', () => {
        expect(SW_MATCH_MEMBER_PREVIEW).toBe(50);
        const html = buildLinkedGroupSummaryHtml({
            groupId: 'x',
            status: 'ready',
            group: { displayName: 'Schüler' },
            members: [{ displayName: 'A', mail: 'a@x.at' }],
            memberCount: 120,
            membersTruncated: true
        });
        expect(html).toContain('Angezeigt: 1 von 120');
    });

    it('zeigt Login-Hinweis ohne Besitzerliste', () => {
        const html = buildLinkedGroupSummaryHtml({
            groupId: 'only-id',
            status: 'needsLogin'
        });
        expect(html).toContain('only-id');
        expect(html).toContain('Melden Sie sich an');
        expect(html).toContain('Details aus Microsoft 365 laden');
        expect(html).not.toContain('Besitzer (');
    });
});
