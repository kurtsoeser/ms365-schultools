import { describe, it, expect, afterEach } from 'vitest';
import {
    primaryHrefForBlock,
    linkSyncHealth,
    actionsForBlock,
    setDatenlandkarteHrefContext,
    resolveDatenlandkarteHref
} from '../src/tools/datenlandkarte/datenlandkarte-register-bridge.js';

describe('datenlandkarte-register-bridge', () => {
    afterEach(() => {
        setDatenlandkarteHrefContext('tool');
    });
    it('verlinkt Stammdaten-Blöcke ins Schulregister', () => {
        expect(primaryHrefForBlock({ id: 'stamm-classes', layer: 'stamm' })).toBe('../tenant.html#klassen');
        expect(primaryHrefForBlock({ id: 'import-webuntis', layer: 'schuljahr' })).toBe('../tenant.html#schueler');
        expect(primaryHrefForBlock({ id: 'spo-sp-klassen', layer: 'sharepoint' })).toBe('../tenant.html#stammdaten');
    });

    it('bewertet Sync-Links anhand der Zähler', () => {
        const metrics = {
            classes: { value: 12 },
            spoListKlassen: { value: 12 },
            students: { value: 100 },
            spoListSchuelerinnen: { value: 90 }
        };
        expect(linkSyncHealth('spo-sp-klassen-stamm', metrics)).toBe('ok');
        expect(linkSyncHealth('spo-sp-schueler-stamm', metrics)).toBe('warn');
    });

    it('liefert Aktionen für Schüler-Block', () => {
        const acts = actionsForBlock({ id: 'year-students', layer: 'schuljahr', href: '../tenant.html#schueler' });
        expect(acts.some((a) => a.href.includes('#schueler'))).toBe(true);
    });

    it('nutzt Register-Hashes im eingebetteten Modus', () => {
        setDatenlandkarteHrefContext('register');
        expect(primaryHrefForBlock({ id: 'stamm-classes', layer: 'stamm' })).toBe('#klassen');
        expect(resolveDatenlandkarteHref('../tools/kursteams.html')).toBe('tools/kursteams.html');
    });
});
