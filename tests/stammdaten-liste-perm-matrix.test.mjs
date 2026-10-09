import { describe, it, expect } from 'vitest';
import {
    normalizeListPermProfiles,
    patchListPermProfiles,
    migrateLegacyPermConfig,
    grantsForListKey,
    normalizeGrantRows
} from '../src/tools/sharepoint/stammdaten-liste-perm-matrix.js';

describe('stammdaten-liste-perm-matrix', () => {
    it('normalizeListPermProfiles nutzt Defaults ohne Speicher', () => {
        const p = normalizeListPermProfiles(null);
        expect(p.schueler.schueler).toBeNull();
        expect(p.faecher.lehrer).toBe('read');
    });

    it('patchListPermProfiles spiegelt Fachgruppen mit Fächer', () => {
        const next = patchListPermProfiles(normalizeListPermProfiles(null), ['faecher', 'fachgruppen'], 'lehrer', 'contribute');
        expect(next.faecher.lehrer).toBe('contribute');
        expect(next.fachgruppen.lehrer).toBe('contribute');
    });

    it('Schüler-Spalte kann auf keinen Zugriff gesetzt werden', () => {
        const next = patchListPermProfiles(normalizeListPermProfiles(null), ['schueler'], 'schueler', null);
        expect(next.schueler.schueler).toBeNull();
    });

    it('migrateLegacyPermConfig erzeugt Zeilen aus drei Gruppen', () => {
        const rows = migrateLegacyPermConfig({
            groupAdminId: 'a1',
            groupAdmin: 'Admin',
            groupLehrerId: 'l1',
            groupLehrer: 'Lehrer',
            listProfiles: null
        });
        expect(rows.length).toBe(2);
        expect(rows[0].groupId).toBe('a1');
        expect(rows[0].cells.klassen).toBe('fullControl');
        expect(rows[1].cells.schueler).toBe('contribute');
    });

    it('grantsForListKey filtert leere Zellen', () => {
        const rows = normalizeGrantRows([
            {
                groupId: 'g1',
                groupLabel: 'G',
                cells: { klassen: 'read', schueler: null }
            }
        ]);
        const grants = grantsForListKey('klassen', rows);
        expect(grants).toEqual([{ groupId: 'g1', groupLabel: 'G', level: 'read' }]);
    });
});
