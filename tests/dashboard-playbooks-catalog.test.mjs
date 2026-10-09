import { describe, expect, it } from 'vitest';
import {
    DASHBOARD_PLAYBOOKS,
    PLAYBOOK_CATEGORIES,
    computePlaybookProgress,
    formatPlaybookMetaLine,
    formatPlaybookProgressLabel,
    playbookBiClass,
    playbooksInCategory,
    playbookProgressBarSegments
} from '../src/shared/dashboard-playbooks-catalog.js';

describe('dashboard-playbooks-catalog', () => {
    it('enthält die Dashboard-Playbooks inkl. Werkzeug-Playbooks', () => {
        expect(DASHBOARD_PLAYBOOKS.length).toBe(9);
        expect(DASHBOARD_PLAYBOOKS.some((p) => p.id === 'schuljahresstart')).toBe(true);
        expect(DASHBOARD_PLAYBOOKS.some((p) => p.id === 'elternsprechtag')).toBe(true);
        expect(DASHBOARD_PLAYBOOKS.some((p) => p.id === 'freistellungen')).toBe(true);
        expect(PLAYBOOK_CATEGORIES.length).toBe(4);
        DASHBOARD_PLAYBOOKS.forEach((p) => {
            expect(PLAYBOOK_CATEGORIES.some((c) => c.id === p.categoryId)).toBe(true);
        });
        expect(playbooksInCategory('grundlagen').length).toBe(3);
    });

    it('zählt erledigte Schritte', () => {
        const p = computePlaybookProgress({ hub: true, lehrer: true }, ['hub', 'lehrer', 'termine']);
        expect(p.done).toBe(2);
        expect(p.total).toBe(3);
        expect(p.status).toBe('in-progress');
    });

    it('erkennt Abschluss und Start', () => {
        expect(
            computePlaybookProgress({ a: true, b: true }, ['a', 'b']).status
        ).toBe('complete');
        expect(computePlaybookProgress({}, ['a', 'b']).status).toBe('not-started');
    });

    it('formatiert Status-Texte', () => {
        expect(formatPlaybookProgressLabel({ done: 0, total: 6, status: 'not-started' })).toBe(
            'Noch nicht gestartet'
        );
        expect(formatPlaybookProgressLabel({ done: 2, total: 6, status: 'in-progress' })).toBe(
            '2 von 6 Schritte erledigt'
        );
        expect(formatPlaybookProgressLabel({ done: 6, total: 6, status: 'complete' })).toBe(
            'Abgeschlossen'
        );
    });

    it('rendert Balken-Segmente', () => {
        const half = playbookProgressBarSegments(0.5, 10);
        expect(half.length).toBe(10);
        expect(half.includes('█')).toBe(true);
        expect(half.includes('░')).toBe(true);
    });

    it('normalisiert Icon-Klassen und Meta-Zeile', () => {
        expect(playbookBiClass('bi-rocket-takeoff')).toBe('bi bi-rocket-takeoff');
        expect(playbookBiClass('bi bi-flag')).toBe('bi bi-flag');
        const def = DASHBOARD_PLAYBOOKS[0];
        expect(formatPlaybookMetaLine(def)).toContain('Schritte');
        expect(formatPlaybookMetaLine(def)).toContain(def.detail || '');
    });
});
