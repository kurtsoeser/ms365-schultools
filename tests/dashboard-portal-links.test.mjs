import { describe, it, expect } from 'vitest';
import {
    normalizePortalQuickLinks,
    activePortalQuickLinks,
    savePortalQuickLinks,
    loadPortalQuickLinks,
    PORTAL_QUICK_LINKS_STORAGE_KEY
} from '../src/shared/dashboard-portal-links.js';

describe('dashboard-portal-links', () => {
    it('füllt Defaults und übernimmt Schul-URLs', () => {
        const links = normalizePortalQuickLinks([
            { id: 'webuntis', url: 'https://example.webuntis.com/', label: 'Untis' }
        ]);
        const untis = links.find((l) => l.id === 'webuntis');
        expect(untis.url).toBe('https://example.webuntis.com/');
        expect(untis.label).toBe('Untis');
        expect(links.find((l) => l.id === 'outlook').url).toContain('outlook.office.com');
    });

    it('activePortalQuickLinks filtert leere URLs', () => {
        const active = activePortalQuickLinks(
            normalizePortalQuickLinks([{ id: 'intranet', url: 'https://intranet.schule.at/' }])
        );
        expect(active.some((l) => l.id === 'outlook')).toBe(true);
        expect(active.some((l) => l.id === 'intranet')).toBe(true);
        expect(active.some((l) => l.id === 'sokrates')).toBe(false);
    });

    it('speichert tenant-scoped in Storage', () => {
        const mem = new Map();
        const storage = {
            getItem: (k) => (mem.has(k) ? mem.get(k) : null),
            setItem: (k, v) => mem.set(k, String(v)),
            removeItem: (k) => mem.delete(k)
        };
        savePortalQuickLinks(
            normalizePortalQuickLinks([{ id: 'sokrates', url: 'https://sokrates.example/' }]),
            storage
        );
        const raw = JSON.parse(storage.getItem(PORTAL_QUICK_LINKS_STORAGE_KEY));
        expect(raw.links || raw).toBeTruthy();
        const loaded = loadPortalQuickLinks(storage);
        expect(loaded.find((l) => l.id === 'sokrates').url).toBe('https://sokrates.example/');
    });
});
