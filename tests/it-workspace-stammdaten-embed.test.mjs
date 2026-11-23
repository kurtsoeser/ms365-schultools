import { beforeEach, describe, expect, it, vi } from 'vitest';
import { collectTenantRegisterNavItems } from '../src/shared/it-workspace-stammdaten-embed.js';

function makeTabBtn({ id, label, icon, group, hidden = false }) {
    return {
        id,
        hidden,
        className: 'tab-btn',
        textContent: label,
        getAttribute(name) {
            if (name === 'aria-hidden') return null;
            if (name === 'data-rail-group') return group || '';
            return null;
        },
        querySelector(sel) {
            if (sel === 'i.bi') return { className: icon || 'bi bi-dot' };
            return null;
        }
    };
}

describe('it-workspace-stammdaten-embed', () => {
    beforeEach(() => {
        const tabs = {
            querySelectorAll(sel) {
                if (sel !== '.tab-btn[role="tab"]') return [];
                return [
                    makeTabBtn({
                        id: 'tabMainStammdaten',
                        label: 'Schulprofil',
                        icon: 'bi bi-building',
                        group: 'schule'
                    }),
                    makeTabBtn({
                        id: 'tabMainSubjects',
                        label: 'Fächer',
                        icon: 'bi bi-journal-text',
                        group: 'unterricht'
                    }),
                    makeTabBtn({
                        id: 'tabMainOrganigramm',
                        label: 'Organigramm',
                        icon: 'bi bi-diagram-3',
                        group: 'uebersicht',
                        hidden: true
                    })
                ];
            }
        };
        vi.stubGlobal('document', {
            getElementById(id) {
                return id === 'tenantMainTabs' ? tabs : null;
            }
        });
        vi.stubGlobal('location', { pathname: '/tenant.html' });
        vi.stubGlobal('window', { location: globalThis.location });
    });

    it('liest sichtbare Stammdaten-Tabs für die IT-Leiste', () => {
        const items = collectTenantRegisterNavItems();
        expect(items.map((i) => i.id)).toEqual(['tabMainStammdaten', 'tabMainSubjects']);
        expect(items[0].label).toMatch(/Schulprofil/);
        expect(items[0].group).toBe('schule');
        expect(items[0].icon).toBe('bi bi-building');
    });
});
