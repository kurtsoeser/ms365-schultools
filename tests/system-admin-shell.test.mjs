import { describe, expect, it } from 'vitest';
import {
    normalizeSystemTab,
    renderSystemAdminPage,
    renderSystemTabBar,
    systemPanelHiddenAttr
} from '../src/shared/system-admin-shell.js';

describe('system-admin-shell', () => {
    it('normalizes unknown tabs to site', () => {
        expect(normalizeSystemTab('')).toBe('site');
        expect(normalizeSystemTab('foo')).toBe('site');
        expect(normalizeSystemTab('perms')).toBe('perms');
        expect(normalizeSystemTab('Danger')).toBe('danger');
    });

    it('hides inactive panels', () => {
        expect(systemPanelHiddenAttr('site', 'site')).toBe('');
        expect(systemPanelHiddenAttr('site', 'perms')).toBe(' hidden');
    });

    it('renders numbered tabs plus danger zone', () => {
        const html = renderSystemTabBar({
            prefix: 'demo',
            activeTab: 'wartung',
            dataAttr: 'data-demo-sys-tab'
        });
        expect(html).toContain('1 · Datenquelle');
        expect(html).toContain('2 · Berechtigungen');
        expect(html).toContain('3 · Kalender');
        expect(html).toContain('4 · Wartung');
        expect(html).toContain('Danger Zone');
        expect(html).toContain('data-demo-sys-tab="wartung"');
        expect(html).toContain('is-active');
        expect(html).toContain('ms-sys-tabs__btn--danger');
    });

    it('renders one visible panel and hides the rest', () => {
        const html = renderSystemAdminPage({
            prefix: 'demo',
            activeTab: 'perms',
            dataAttr: 'data-demo-sys-tab',
            ariaLabel: 'Systemadministration Demo',
            panels: [
                { id: 'site', step: '1', title: 'Datenquelle', body: '<p id="site-body">site</p>' },
                { id: 'perms', step: '2', title: 'Berechtigungen', body: '<p id="perms-body">perms</p>' },
                { id: 'danger', title: 'Danger Zone', danger: true, body: '<button type="button">Reset</button>' }
            ]
        });
        expect(html).toContain('id="demoSysPanel-perms"');
        expect(html).toContain('id="demoSysPanel-site"');
        expect(html).toMatch(/id="demoSysPanel-site"[^>]*hidden/);
        expect(html).not.toMatch(/id="demoSysPanel-perms"[^>]*hidden/);
        expect(html).toContain('ms-sys-card--danger');
        expect(html).toContain('data-demo-sys-tab="perms"');
        expect(html).toContain('Systemadministration Demo');
    });
});
