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
    sortDriveItems,
    pushBreadcrumb,
    sliceBreadcrumb
} from '../src/tools/datei-migration/datei-migration-logic.js';
import {
    buildSpielwiesenPlan,
    isSpielwiesenGroup,
    buildTeacherSpielPlan,
    buildBulkTeacherPlans,
    buildDemoStudentPlan,
    isDemoStudentUser,
    validateDemoPool,
    MAX_DEMO_STUDENTS
} from '../src/tools/spielwiesen/spielwiesen-logic.js';
import {
    sanitizeEducationClassCode,
    parseTeamsOperationPath,
    EDUCATION_OBJECT_TYPE_EXTENSION
} from '../src/shared/education-class-team.js';

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

    it('navigiert Breadcrumbs', () => {
        const root = [{ id: 'root', name: 'Stamm' }];
        const deeper = pushBreadcrumb(root, { id: 'f1', name: 'Material' });
        expect(deeper).toHaveLength(2);
        expect(sliceBreadcrumb(deeper, 0)).toEqual([{ id: 'root', name: 'Stamm' }]);
    });
});

describe('spielwiesen-logic', () => {
    it('baut Spielwiesen-Plan mit Notebook-Checkliste', () => {
        const p = buildSpielwiesenPlan({ label: 'Teams Basics', year: '2026' });
        expect(p.ok).toBe(true);
        expect(p.mailNickname).toMatch(/^spiel-2026-/);
        expect(p.educationClass).toBe(true);
        expect(p.notebookChecklist.length).toBeGreaterThan(3);
        expect(isSpielwiesenGroup({ mailNickname: p.mailNickname, displayName: p.displayName })).toBe(true);
    });

    it('plant Lehrer-Bulk und Demo-Schüler', () => {
        const bulk = buildBulkTeacherPlans({
            year: '2026',
            teachers: [
                { code: 'MU', name: 'Müller', email: 'mu@schule.at' },
                { code: 'XY', name: 'Ohne Mail', email: '' }
            ],
            selectedCodes: ['MU', 'XY']
        });
        expect(bulk.plans).toHaveLength(2);
        expect(bulk.ok).toBe(false);
        expect(buildTeacherSpielPlan({ code: 'MU', email: 'mu@schule.at', year: 2026 }).mailNickname).toBe(
            'spiel-2026-mu'
        );
        const stu = buildDemoStudentPlan({ index: 3, domain: 'schule.at' });
        expect(stu.ok).toBe(true);
        expect(stu.userPrincipalName).toBe('demo.schueler03@schule.at');
        expect(isDemoStudentUser({ displayName: stu.displayName, department: 'DEMO' })).toBe(true);
        expect(validateDemoPool([{ id: '1' }]).ok).toBe(true);
        expect(validateDemoPool(new Array(MAX_DEMO_STUDENTS + 1).fill({ id: 'x' })).ok).toBe(false);
    });
});

describe('education-class-team', () => {
    it('sanitized classCode und Operation-Pfad', () => {
        expect(sanitizeEducationClassCode('spiel-2026-mu!')).toBe('spiel2026mu');
        expect(
            parseTeamsOperationPath("https://graph.microsoft.com/v1.0/teams('tid')/operations('oid')")
        ).toBe('/teams/tid/operations/oid');
        expect(EDUCATION_OBJECT_TYPE_EXTENSION).toContain('Education_ObjectType');
    });
});
