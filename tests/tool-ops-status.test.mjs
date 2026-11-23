import { describe, it, expect } from 'vitest';
import { resolveToolOpsKind, renderToolOpsStatusCard } from '../src/shared/tool-ops-status.js';

describe('tool-ops-status', () => {
    it('Priorität: need_login vor Site-Hinweis', () => {
        expect(
            resolveToolOpsKind({
                requireLogin: true,
                loggedIn: false,
                sync: { source: 'none', ok: false, siteUrl: '' }
            })
        ).toBe('need_login');
    });

    it('site_missing ohne Login-Pflicht', () => {
        expect(
            resolveToolOpsKind({
                requireLogin: false,
                loggedIn: false,
                sync: { source: 'none', ok: false, siteUrl: '' }
            })
        ).toBe('site_missing');
    });

    it('spo_local bei demo/local', () => {
        expect(
            resolveToolOpsKind({
                requireLogin: true,
                loggedIn: true,
                sync: { source: 'local', ok: false, siteUrl: 'https://contoso.sharepoint.com/sites/x' }
            })
        ).toBe('spo_local');
    });

    it('scope wenn live+ok und eingeloggt', () => {
        expect(
            resolveToolOpsKind({
                requireLogin: true,
                loggedIn: true,
                sync: { source: 'live', ok: true, siteUrl: 'https://x' }
            })
        ).toBe('scope');
    });

    it('leer bei Krankmeldung live+ok', () => {
        expect(
            resolveToolOpsKind({
                requireLogin: false,
                loggedIn: true,
                sync: { source: 'live', ok: true, siteUrl: 'https://x' }
            })
        ).toBe('');
    });

    it('render: eine Login-Karte, kein zweiter SharePoint-Text', () => {
        const html = renderToolOpsStatusCard({
            nounPlural: 'Elternbriefe',
            requireLogin: true,
            loggedIn: false,
            sync: { source: 'none', ok: false, siteUrl: '' },
            setupHref: 'sharepoint-liste-elternbrief.html'
        });
        expect(html).toContain('ms365-tool-ops--login');
        expect(html).toContain('anmelden');
        expect(html).not.toContain('Noch keine Site');
        expect(html.match(/role="status"/g) || []).toHaveLength(1);
    });

    it('render: Admin bekommt Setup-Link bei fehlender Site', () => {
        const html = renderToolOpsStatusCard({
            nounPlural: 'Einsammler',
            listWord: 'Listen',
            requireLogin: true,
            loggedIn: true,
            isAdmin: true,
            accountLabel: 'Ada',
            scopeToggleId: 'einShowAll',
            scopeToggleLabel: 'Alle Einsammler (Admin)',
            sync: { source: 'none', ok: false, siteUrl: '' },
            setupHref: 'sharepoint-liste-einsammler.html'
        });
        expect(html).toContain('Noch keine Site');
        expect(html).toContain('sharepoint-liste-einsammler.html');
        expect(html).toContain('einShowAll');
        expect(html).toContain('Angemeldet');
    });
});
