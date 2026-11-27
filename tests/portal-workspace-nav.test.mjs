import { describe, expect, it } from 'vitest';
import {
    buildDefaultPortalTree,
    filterPortalTreeToVisible,
    normalizeIconClass,
    portalShellViewOrder,
    portalToolHref,
    portalTreeFromSections,
    reconcilePortalTreeWithCatalog,
    resolvePortalLocation,
    resolvePortalWorkspaceView,
    sanitizePortalNavHref
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
        expect(freistellung.section.id).toBe('antraege');
        expect(freistellung.section.label).toBe('Anträge und Formulare');
        expect(freistellung.tool.label).toMatch(/Freistellungen/);
        const schuelerFrei = resolvePortalLocation(tree, 'tools/freistellung-planer.html');
        expect(schuelerFrei.section.id).toBe('antraege');
        const matura = resolvePortalLocation(tree, 'tools/srdp-anmeldung.html');
        expect(matura.section.id).toBe('antraege');
        const wahlfach = resolvePortalLocation(tree, 'tools/wahlfach-zuteilung.html');
        expect(wahlfach.section.id).toBe('antraege');
        expect(wahlfach.section.label).toBe('Anträge und Formulare');
        const schuelerTree = buildDefaultPortalTree('schueler');
        const schuelerIds = schuelerTree.flatMap((sec) => sec.tools.map((t) => t.id));
        expect(schuelerIds).toContain('wahlfach-anmeldung');
        expect(schuelerIds).not.toContain('wahlfach-zuteilung');

        const ids = tree.flatMap((sec) => sec.tools.map((t) => t.id));
        expect(ids).not.toContain('kursteam-einzeln');
        expect(ids).not.toContain('unterrichtsteams-katalog');
        expect(ids).not.toContain('diplomarbeiten');
        const katalog = resolvePortalLocation(tree, 'tools/kursteams.html');
        expect(katalog.section.id).toBe('unterricht');
        expect(katalog.tool.id).toBe('kursteams');

        const both = portalTreeFromSections([
            {
                id: 'unterricht',
                label: 'Unterricht',
                icon: 'bi-mortarboard',
                toolIds: ['kursteams', 'kursteam-einzeln']
            }
        ]);
        expect(resolvePortalLocation(both, 'tools/kursteams.html?mode=single').tool.id).toBe(
            'kursteam-einzeln'
        );
        expect(resolvePortalLocation(both, 'tools/kursteams.html').tool.id).toBe('kursteams');
    });

    it('öffnet das Intranet für Lehrkräfte auf der hinterlegten SharePoint-Seite', () => {
        const tree = portalTreeFromSections(
            [
                {
                    id: 'intranet',
                    label: 'Intranet',
                    icon: 'bi-house-door',
                    toolIds: ['sharepoint-intranet-hub', 'termine-hub']
                }
            ],
            { intranetUrl: 'https://kurtrocks.sharepoint.com/sites/intranet' }
        );
        expect(tree[0].tools[0].href).toBe('https://kurtrocks.sharepoint.com/sites/intranet');
        expect(tree[0].tools[1].href).toBe('tools/termine-hub.html');
        expect(portalToolHref('sharepoint-intranet-hub', { intranetUrl: '' })).toBe('');
        expect(portalToolHref('termine-hub')).toBe('tools/termine-hub.html');
        expect(sanitizePortalNavHref('https://kurtrocks.sharepoint.com/sites/intranet')).toBe(
            'https://kurtrocks.sharepoint.com/sites/intranet'
        );
        expect(sanitizePortalNavHref('javascript:alert(1)')).toBe('');
    });

    it('Lehrer-Portal enthält KlassenChats nicht – auch nicht aus einem alten Sitzungsbaum', () => {
        const tree = buildDefaultPortalTree('lehrer');
        const ids = tree.flatMap((sec) => sec.tools.map((t) => t.id));
        expect(ids).toContain('kollegium-umfragen');
        expect(ids).toContain('elternsprechtag');
        expect(ids).not.toContain('klassenchats');
        expect(ids).not.toContain('schulsprecherwahl');
        expect(ids).not.toContain('elternsprechtag-bookings');
        expect(ids).not.toContain('kursteam-einzeln');
        expect(ids).not.toContain('unterrichtsteams-katalog');
        expect(ids).not.toContain('diplomarbeiten');
        const stale = filterPortalTreeToVisible(
            [
                {
                    id: 'kollegium',
                    label: 'Kollegium',
                    tools: [
                        { id: 'kollegium-umfragen', href: 'tools/kollegium-umfragen.html', label: 'Umfragen' },
                        { id: 'klassenchats', href: 'tools/klassenchats.html', label: 'KlassenChats' },
                        {
                            id: 'kursteam-einzeln',
                            href: 'tools/kursteams.html?mode=single',
                            label: 'Einzelne Unterrichtsteams'
                        },
                        {
                            id: 'unterrichtsteams-katalog',
                            href: 'tools/unterrichtsteams-katalog.html',
                            label: 'Unterrichtsteams-Katalog'
                        },
                        {
                            id: 'diplomarbeiten',
                            href: 'tools/diplomarbeiten.html',
                            label: 'Diplomarbeiten / VWA'
                        }
                    ]
                }
            ],
            'lehrer'
        );
        expect(stale.flatMap((sec) => sec.tools.map((t) => t.id))).toEqual(['kollegium-umfragen']);
    });

    it('Schüler-Baum enthält nicht die Lehrkräfte-Freistellung', () => {
        const tree = buildDefaultPortalTree('schueler');
        const ids = tree.flatMap((sec) => sec.tools.map((t) => t.id));
        expect(ids).not.toContain('schulsprecherwahl');
        expect(ids).not.toContain('lehrer-freistellung-planer');
        expect(ids).not.toContain('elternbrief');
        expect(ids).toContain('wahlfach-anmeldung');
        expect(ids).not.toContain('wahlfach-zuteilung');
    });

    it('legt ein Werkzeug aus einer alten Sitzung in die Katalog-Gruppe', () => {
        const moved = reconcilePortalTreeWithCatalog(
            [
                {
                    id: 'unterricht',
                    label: 'Unterricht',
                    tools: [
                        {
                            id: 'wahlfach-zuteilung',
                            href: 'tools/wahlfach-zuteilung.html',
                            label: 'Wahlfach-Zuteilung'
                        },
                        {
                            id: 'schularbeiten-planer',
                            href: 'tools/schularbeiten-planer.html',
                            label: 'Schularbeiten'
                        }
                    ]
                }
            ],
            'lehrer'
        );
        const wahl = moved.find((sec) => sec.tools.some((t) => t.id === 'wahlfach-zuteilung'));
        expect(wahl && wahl.id).toBe('antraege');
        expect(moved.find((sec) => sec.id === 'unterricht').tools.map((t) => t.id)).toEqual([
            'schularbeiten-planer'
        ]);
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

    it('hängt das Schüler-Formular an die Schüler-Leiste, auch aus der Lehrkraft-Vorschau', () => {
        expect(portalShellViewOrder('lehrer')).toEqual(['lehrer', 'schueler']);
        expect(portalShellViewOrder('schueler')[0]).toBe('schueler');
        expect(resolvePortalLocation(buildDefaultPortalTree('lehrer'), 'tools/wahlfach-anmeldung.html')).toBeNull();
        const loc = resolvePortalLocation(
            buildDefaultPortalTree('schueler'),
            'tools/wahlfach-anmeldung.html'
        );
        expect(loc && loc.tool && loc.tool.id).toBe('wahlfach-anmeldung');
    });
});
