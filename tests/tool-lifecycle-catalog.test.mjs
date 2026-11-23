import { describe, expect, it } from 'vitest';
import { DASHBOARD_TOOL_LABELS } from '../src/shared/dashboard-audience-catalog.js';
import {
    lifecycleCounts,
    lifecycleForTool,
    listToolLifecycles
} from '../src/shared/tool-lifecycle-catalog.js';

describe('tool-lifecycle-catalog (a11)', () => {
    it('klassifiziert jedes Dashboard-Werkzeug', () => {
        const ids = Object.keys(DASHBOARD_TOOL_LABELS);
        expect(ids.length).toBeGreaterThan(40);
        for (const id of ids) {
            const life = lifecycleForTool(id);
            expect(['kern', 'pflege', 'archiv']).toContain(life);
        }
        expect(listToolLifecycles()).toHaveLength(ids.length);
    });

    it('markiert Kern-Planer als kern und liefert Zähler', () => {
        expect(lifecycleForTool('lehrer-freistellung-planer')).toBe('kern');
        expect(lifecycleForTool('srdp-anmeldung')).toBe('kern');
        expect(lifecycleForTool('spielwiesen')).toBe('archiv');
        const c = lifecycleCounts();
        expect(c.total).toBe(Object.keys(DASHBOARD_TOOL_LABELS).length);
        expect(c.kern + c.pflege + c.archiv).toBe(c.total);
        expect(c.kern).toBeGreaterThan(5);
    });
});
