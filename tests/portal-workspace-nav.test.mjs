import { describe, expect, it } from 'vitest';
import {
    buildDefaultPortalTree,
    normalizeIconClass,
    portalTreeFromSections,
    resolvePortalLocation,
    resolvePortalWorkspaceView
} from '../src/shared/portal-workspace-nav.js';

describe('portal-workspace-nav', () => {
    it('normalisiert Bootstrap-Icons für die Leiste', () => {
        expect(normalizeIconClass('bi-building')).toBe('bi bi-building');
        expect(normalizeIconClass('bi bi-geo-alt')).toBe('bi bi-geo-alt');
        expect(normalizeIconClass('')).toBe('bi bi-circle');
    });

    it('baut aus den Portal-Gruppen einen Baum mit echten Werkzeuglinks', () => {
        const tree = portalTreeFromSections([
            {
                id: 'schulalltag',
                label: 'Schulalltag',
                icon: 'bi-building',
                toolIds: ['termine-hub', 'schulaktivitaeten-planer']
            },
            { id: 'leer', label: 'Leer', icon: 'bi-grid', toolIds: [] }
        ]);
        expect(tree).toHaveLength(1);
        expect(tree[0].label).toBe('Schulalltag');
        expect(tree[0].tools.map((t) => t.label)).toEqual(['Kalender und Termine', 'Schulaktivitäten']);
        expect(tree[0].tools[1].href).toBe('tools/schulaktivitaeten-planer.html');
    });

    it('findet das Werkzeug in seiner Gruppe, auch mit Query', () => {
        const tree = buildDefaultPortalTree('lehrer');
        const alltag = resolvePortalLocation(tree, 'tools/schulaktivitaeten-planer.html');
        expect(alltag.section.id).toBe('schulalltag');
        expect(alltag.tool.label).toBe('Schulaktivitäten');

        const freistellung = resolvePortalLocation(tree, 'tools/lehrer-freistellung-planer.html');
        expect(freistellung.section.id).toBe('unterricht');
        expect(freistellung.tool.label).toMatch(/Freistellungen/);

        const einzeln = resolvePortalLocation(tree, 'tools/kursteams.html?mode=single');
        expect(einzeln.tool.id).toBe('kursteam-einzeln');
        const katalog = resolvePortalLocation(tree, 'tools/kursteams.html');
        expect(katalog.tool.id).toBe('kursteams');
    });

    it('Schüler-Baum enthält nicht die Lehrkräfte-Freistellung', () => {
        const tree = buildDefaultPortalTree('schueler');
        const ids = tree.flatMap((sec) => sec.tools.map((t) => t.id));
        expect(ids).toContain('schulsprecherwahl');
        expect(ids).not.toContain('lehrer-freistellung-planer');
        expect(ids).not.toContain('elternbrief');
    });

    it('entscheidet die Portal-Ansicht nach Rolle und IT-Bereich', () => {
        const store = new Map();
        globalThis.localStorage = {
            getItem: (key) => (store.has(key) ? store.get(key) : null),
            setItem: (key, value) => store.set(key, String(value)),
            removeItem: (key) => store.delete(key)
        };
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'full');
        expect(resolvePortalWorkspaceView({ loggedIn: true, isIt: true })).toBe('');
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'schueler');
        expect(resolvePortalWorkspaceView({ loggedIn: true, isIt: true })).toBe('schueler');
        localStorage.setItem('ms365-dashboard-it-preview-v1', 'portal');
        expect(resolvePortalWorkspaceView({ loggedIn: true, isIt: true })).toBe('lehrer');
        expect(resolvePortalWorkspaceView({ loggedIn: true, isLehrer: true })).toBe('lehrer');
        expect(resolvePortalWorkspaceView({ loggedIn: true, isSchueler: true })).toBe('schueler');
        expect(resolvePortalWorkspaceView({ loggedIn: false, isLehrer: true })).toBe('');
        expect(resolvePortalWorkspaceView(null)).toBe('');
    });
});
