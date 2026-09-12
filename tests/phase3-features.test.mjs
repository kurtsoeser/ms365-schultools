import { describe, it, expect } from 'vitest';
import {
    buildDiplomPlan,
    isDiplomGroup,
    filterDiplomGroups,
    buildDiplomMailNickname
} from '../src/tools/diplomarbeiten/diplomarbeiten-logic.js';
import {
    validateMigrationSelection,
    buildCopyBody,
    normalizeDriveItem,
    sortDriveItems
} from '../src/tools/datei-migration/datei-migration-logic.js';
import { buildSpielwiesenPlan, isSpielwiesenGroup } from '../src/tools/spielwiesen/spielwiesen-logic.js';

describe('diplomarbeiten-logic', () => {
    it('baut Standard-Namen', () => {
        const p = buildDiplomPlan({ year: '2027', topic: 'KI im Unterricht', mentor: 'Müller', student: 'Anna' });
        expect(p.ok).toBe(true);
        expect(p.displayName).toContain('Diplomarbeit 2027');
        expect(p.mailNickname).toMatch(/^dipl-2027-/);
        expect(buildDiplomMailNickname({ year: 2026, topic: 'Test Äpfel' })).toContain('dipl-2026-');
    });

    it('filtert Diplom-Gruppen', () => {
        const groups = [
            { displayName: 'Diplomarbeit 2026 – Solar', mailNickname: 'dipl-2026-solar' },
            { displayName: 'Klasse 4HM', mailNickname: 'jg20274hm' }
        ];
        expect(isDiplomGroup(groups[0])).toBe(true);
        expect(filterDiplomGroups(groups)).toHaveLength(1);
    });
});

describe('datei-migration-logic', () => {
    it('validiert Auswahl', () => {
        const bad = validateMigrationSelection({ sourceGroupId: 'a', destGroupId: 'a', itemIds: [] });
        expect(bad.ok).toBe(false);
        const ok = validateMigrationSelection({ sourceGroupId: 'a', destGroupId: 'b', itemIds: ['x'] });
        expect(ok.ok).toBe(true);
        const body = buildCopyBody({ destDriveId: 'd1', destFolderId: 'root', newName: 'Kopie' });
        expect(body.parentReference.driveId).toBe('d1');
        expect(body.name).toBe('Kopie');
    });

    it('sortiert Ordner vor Dateien', () => {
        const rows = sortDriveItems([
            normalizeDriveItem({ id: '1', name: 'z.txt', size: 1 }),
            normalizeDriveItem({ id: '2', name: 'a', folder: { childCount: 2 }, size: 0 })
        ]);
        expect(rows[0].isFolder).toBe(true);
    });
});

describe('spielwiesen-logic', () => {
    it('baut Spielwiesen-Plan mit Notebook-Checkliste', () => {
        const p = buildSpielwiesenPlan({ label: 'Teams Basics', year: '2026' });
        expect(p.ok).toBe(true);
        expect(p.mailNickname).toMatch(/^spiel-2026-/);
        expect(p.notebookChecklist.length).toBeGreaterThan(3);
        expect(isSpielwiesenGroup({ mailNickname: p.mailNickname, displayName: p.displayName })).toBe(true);
    });
});
