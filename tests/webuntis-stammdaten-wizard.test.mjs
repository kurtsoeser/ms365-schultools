import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';
import {
    buildWizardPreview,
    compileWizardApplyPayload,
    applyRecommendedSubjectImport,
    mergeWebuntisImportWithExisting,
    summarizeWebuntisMergeResult
} from '../src/shared/webuntis-stammdaten-wizard-logic.js';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');
const fixtures = join(root, 'fixtures', 'webuntis');

function loadScript(rel, sandbox) {
    runInContext(readFileSync(join(projectRoot, rel), 'utf8'), sandbox, { filename: join(projectRoot, rel) });
}

function loadApis() {
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    loadScript('src/shared/person-email-from-name.js', sandbox);
    loadScript('src/shared/eltern-guardians.js', sandbox);
    loadScript('src/shared/tenant-settings-core.js', sandbox);
    loadScript('src/shared/webuntis-export-import.js', sandbox);
    loadScript('src/shared/school-sis-import.js', sandbox);
    return {
        webuntis: sandbox.ms365WebuntisExportImport,
        schoolSis: sandbox.ms365SchoolSisImport,
        parseTeachersLines: sandbox.ms365TenantSettingsParseTeachersLines,
        parseStudentsLines: sandbox.ms365TenantSettingsParseStudentsLines,
        parseSubjectsLines: sandbox.ms365TenantSettingsParseSubjectsLines,
        parseClassesLines: sandbox.ms365TenantSettingsParseClassesLines
    };
}

function loadJson(name) {
    return JSON.parse(readFileSync(join(fixtures, name), 'utf8'));
}

describe('webuntis-stammdaten-wizard-logic', () => {
    let deps;

    beforeEach(() => {
        deps = loadApis();
    });

    it('baut Vorschau aus Teacher- und Student-Export', () => {
        const preview = buildWizardPreview(
            {
                sheets: [
                    { name: 'Teacher_2026.csv', aoa: loadJson('teacher-stammdaten-aoa.json') },
                    { name: 'Student_aoa.json', aoa: loadJson('student-aoa.json') },
                    { name: 'LegalGuardian.json', aoa: loadJson('guardian-aoa.json') }
                ],
                classImports: [],
                emailOpts: { domain: 'schule.at', pattern: 'vorname.nachname', firstNameMode: 'first' }
            },
            deps
        );
        expect(preview.teachers.length).toBeGreaterThanOrEqual(2);
        expect(preview.students.length).toBe(3);
        expect(preview.students.some((s) => s.parentCount >= 1)).toBe(true);
        const payload = compileWizardApplyPayload(preview, deps);
        expect(payload.teachersLines).toContain('ALTH');
        expect(payload.studentsLines).toContain('1A');
    });

    it('verknüpft Klassen-PDF mit Teacher-Export (KV + Abschlussjahr)', () => {
        const classText = readFileSync(join(fixtures, 'class-pdf-text.txt'), 'utf8');
        const preview = buildWizardPreview(
            {
                sheets: [{ name: 'Teacher_2026.csv', aoa: loadJson('teacher-stammdaten-aoa.json') }],
                classPdfImports: [{ source: 'Class.pdf', text: classText }],
                schoolYearEnd: 2027,
                emailOpts: { domain: 'schule.at', pattern: 'vorname.nachname', firstNameMode: 'first' }
            },
            deps
        );
        expect(preview.classes.length).toBeGreaterThan(10);
        const hak = preview.classes.find((c) => c.code === '1AK');
        expect(hak).toBeTruthy();
        expect(hak.year).toBe('2031');
        expect(hak.headCode).toBe('HAGAU');
        const withMail = preview.classes.find((c) => c.code === '1AS');
        expect(withMail.headCode).toBe('RAUNE');
        expect(withMail.headName.length).toBeGreaterThan(1);
    });

    it('empfohlene Fächer-Auswahl lässt Basis, schließt Nummerierung und Verwaltung aus', () => {
        const rows = [
            { code: 'D', name: 'Deutsch', selected: true },
            { code: 'D1', name: 'Deutsch 1', selected: true },
            { code: 'ADM', name: 'Admin', selected: true },
            { code: 'M', name: 'Mathe', selected: true, admin: false }
        ];
        const stats = applyRecommendedSubjectImport(rows);
        expect(stats.selected).toBe(2);
        expect(rows.find((r) => r.code === 'D').selected).toBe(true);
        expect(rows.find((r) => r.code === 'M').selected).toBe(true);
        expect(rows.find((r) => r.code === 'D1').selected).toBe(false);
        expect(rows.find((r) => r.code === 'ADM').selected).toBe(false);
    });

    it('empfohlene Auswahl behält Hauptfach bei Ü- und Plus-Varianten', () => {
        const rows = [
            { code: 'RW', name: 'Religion', selected: true },
            { code: 'RWÜ', name: 'Religion Ü', selected: true },
            { code: 'M', name: 'Mathe', selected: true },
            { code: 'M+', name: 'Mathe Plus', selected: true }
        ];
        const stats = applyRecommendedSubjectImport(rows);
        expect(stats.selected).toBe(2);
        expect(rows.find((r) => r.code === 'RW').selected).toBe(true);
        expect(rows.find((r) => r.code === 'RWÜ').selected).toBe(false);
        expect(rows.find((r) => r.code === 'M').selected).toBe(true);
        expect(rows.find((r) => r.code === 'M+').selected).toBe(false);
    });

    it('mergeWebuntisImportWithExisting vermeidet Doppelzeilen und aktualisiert Treffer', () => {
        const apis = loadApis();
        const existing = {
            teachersLines: 'MU;Max Mustermann;max@schule.at',
            studentsLines: '1A;Anna Beispiel;anna@schule.at;#id:EXT1001',
            subjectsLines: 'D;Deutsch',
            classesLines: '1AK;2030;1A-Klasse'
        };
        const payload = {
            teachersLines: 'MU;Max Mustermann (WU);max@schule.at',
            studentsLines:
                '1B;Anna Beispiel;anna@schule.at;#id:EXT1001;Eltern;eltern@schule.at',
            subjectsLines: 'D;Deutsch (WU)\nM;Mathematik',
            classesLines: '1AK;2030;1A-Klasse;KV Name;kv@schule.at'
        };
        const merged = mergeWebuntisImportWithExisting(existing, payload, apis);
        expect(merged.lines.teachersLines.split('\n')).toHaveLength(1);
        expect(merged.lines.teachersLines).toContain('Max Mustermann (WU)');
        expect(merged.stats.teachers.updated).toBe(1);
        expect(merged.stats.subjects.added).toBe(1);
        expect(merged.stats.subjects.updated).toBe(1);
        expect(merged.studentDiff.counts.updated).toBeGreaterThanOrEqual(1);
        expect(merged.studentDiff.counts.added).toBe(0);
        expect(merged.lines.studentsLines).toContain('1B');
        expect(merged.lines.studentsLines).toContain('eltern@schule.at');
        const summary = summarizeWebuntisMergeResult(merged.stats, merged.studentDiff, apis.schoolSis);
        expect(summary).toMatch(/Schüler/);
    });
});
