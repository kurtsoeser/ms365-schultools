import { describe, expect, it } from 'vitest';
import {
    catalogCardMatchesQuery,
    entryFromCatalogCard,
    groupQuickSearchHits,
    rankQuickSearch,
    scoreQuickSearchEntry
} from '../src/shared/dashboard-quick-search-logic.js';

function entry(partial) {
    return {
        id: partial.id || 'x',
        toolId: partial.toolId || partial.id || 'x',
        title: partial.title || 'X',
        keywords: partial.keywords || '',
        blurb: partial.blurb || '',
        kind: partial.kind || 'werkzeug',
        group: partial.group || '',
        href: partial.href || 'tools/x.html',
        icon: 'bi bi-box'
    };
}

describe('dashboard-quick-search-logic', () => {
    it('rankt sRDP vor allgemeinen Intranet-Tools', () => {
        const hits = rankQuickSearch(
            [
                entry({
                    id: 'sharepoint-intranet-hub',
                    title: 'Schul-Intranet',
                    keywords: 'intranet hub sharepoint website',
                    blurb: 'Schul-Website und Hub'
                }),
                entry({
                    id: 'srdp-anmeldung',
                    toolId: 'srdp-anmeldung',
                    title: 'Maturaanmeldung (sRDP)',
                    keywords: 'srdp anmeldung matura sharepoint',
                    blurb: 'Matura-Anmeldung'
                }),
                entry({
                    id: 'srdp-mail-setup',
                    toolId: 'srdp-mail-setup',
                    title: 'sRDP Bestätigungsmail',
                    keywords: 'srdp matura mail',
                    blurb: 'Flow-Paket'
                }),
                entry({
                    id: 'lehrerliste',
                    title: 'Lehrerliste',
                    keywords: 'lehrer liste sharepoint',
                    blurb: 'Intranet-Liste'
                })
            ],
            'srdp'
        );

        expect(hits.length).toBeGreaterThanOrEqual(2);
        expect(hits.every((h) => /srdp/i.test(h.entry.id + h.entry.title))).toBe(true);
        expect(hits.some((h) => h.entry.id === 'sharepoint-intranet-hub')).toBe(false);
        expect(hits[0].entry.id).toMatch(/srdp/);
    });

    it('gruppiert Treffer nach Kategorie', () => {
        const groups = groupQuickSearchHits([
            { entry: entry({ id: 'a', kind: 'werkzeug', title: 'A' }), score: 90, match: 'exact' },
            { entry: entry({ id: 'b', kind: 'playbook', title: 'B' }), score: 80, match: 'prefix' },
            { entry: entry({ id: 'c', kind: 'einstellung', title: 'C' }), score: 70, match: 'token' }
        ]);
        expect(groups.map((g) => g.kind)).toEqual(['werkzeug', 'playbook', 'einstellung']);
    });

    it('gruppierte Anzeige-Reihenfolge weicht vom Ranking ab (Klick-Index)', () => {
        const ranked = [
            { entry: entry({ id: 'portal-prep', kind: 'einstellung', title: 'Portal vorbereiten' }), score: 100, match: 'exact' },
            { entry: entry({ id: 'webuntis', kind: 'werkzeug', title: 'Daten importieren' }), score: 50, match: 'keyword' }
        ];
        const groups = groupQuickSearchHits(ranked);
        const displayOrder = groups.flatMap((g) => g.hits.map((h) => h.entry.id));
        expect(displayOrder).toEqual(['webuntis', 'portal-prep']);
        expect(displayOrder[0]).not.toBe(ranked[0].entry.id);
    });

    it('nutzt Aliase für kurze Codes wie SLG', () => {
        const hit = scoreQuickSearchEntry(
            entry({
                id: 'slg-schueler',
                toolId: 'slg-schueler',
                title: 'Schüler:innen-Sammelgruppe',
                keywords: 'sammelgruppe schüler'
            }),
            'slg'
        );
        expect(hit).toBeTruthy();
        expect(hit.score).toBeGreaterThanOrEqual(90);
    });

    it('findet Portal vorbereiten über Dashboard und Portal', () => {
        const portalPrep = entry({
            id: 'einstellung-portal-vorbereiten',
            toolId: 'einstellung-portal-vorbereiten',
            title: 'Portal vorbereiten',
            keywords: 'portal vorbereiten dashboard gestalten freigeben',
            kind: 'einstellung',
            href: 'dashboard-werkzeug-zugriff.html'
        });
        const portalHome = entry({
            id: 'einstellung-portal',
            toolId: 'einstellung-portal',
            title: 'Portal-Ansicht',
            keywords: 'portal lehrer start',
            kind: 'einstellung',
            href: 'index.html'
        });
        const dashHits = rankQuickSearch([portalPrep, portalHome], 'dashboard');
        expect(dashHits.some((h) => h.entry.id === 'einstellung-portal-vorbereiten')).toBe(true);
        expect(dashHits[0].entry.id).toBe('einstellung-portal-vorbereiten');

        const portalHits = rankQuickSearch([portalPrep, portalHome], 'portal vorbereiten');
        expect(portalHits[0].entry.id).toBe('einstellung-portal-vorbereiten');
    });

    it('ignoriert Status-Text auf Katalog-Karten', () => {
        const card = {
            getAttribute: (k) => {
                if (k === 'data-tool-id') return 'sharepoint-intranet-hub';
                if (k === 'data-search') return 'intranet hub';
                if (k === 'data-cluster') return 'intranet';
                if (k === 'aria-label') return 'Schul-Intranet';
                return '';
            },
            querySelector: (sel) => {
                if (sel === 'h2') return { textContent: 'Schul-Intranet', querySelector: () => null };
                if (sel === 'p') return { textContent: 'Hub aufsetzen' };
                if (sel === 'a.btn[href], a[href]') return { getAttribute: () => 'tools/sharepoint-intranet-hub.html' };
                return null;
            }
        };
        // Simulierter „Status“-Text darf nicht matchen – entryFromCatalogCard liest ihn nicht
        const entryBuilt = entryFromCatalogCard(card);
        expect(entryBuilt.title).toBe('Schul-Intranet');
        expect(catalogCardMatchesQuery(card, 'srdp')).toBe(false);
        expect(catalogCardMatchesQuery(card, 'intranet')).toBe(true);
    });
});
