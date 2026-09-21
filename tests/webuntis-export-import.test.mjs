import { describe, it, expect, beforeEach } from 'vitest';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createContext, runInContext } from 'node:vm';

const root = dirname(fileURLToPath(import.meta.url));
const projectRoot = join(root, '..');
const fixtures = join(root, 'fixtures', 'webuntis');

function loadScript(rel, sandbox) {
    const full = join(projectRoot, rel);
    runInContext(readFileSync(full, 'utf8'), sandbox, { filename: full });
}

function loadAll() {
    const sandbox = { console };
    sandbox.window = sandbox;
    createContext(sandbox);
    loadScript('src/shared/person-email-from-name.js', sandbox);
    loadScript('src/shared/webuntis-export-import.js', sandbox);
    loadScript('src/shared/school-sis-import.js', sandbox);
    return sandbox;
}

function loadJson(name) {
    return JSON.parse(readFileSync(join(fixtures, name), 'utf8'));
}

describe('person-email-from-name', () => {
    let api;

    beforeEach(() => {
        api = loadAll().ms365PersonEmailFromName;
    });

    it('erzeugt vorname.nachname mit Umlauten', () => {
        const r = api.suggestEmail(
            { foreName: 'Jürgen', longName: 'Müller' },
            { domain: 'schule.at', pattern: 'vorname.nachname' }
        );
        expect(r.email).toBe('juergen.mueller@schule.at');
        expect(r.generated).toBe(true);
    });

    it('nutzt standardmäßig nur den ersten Vornamen', () => {
        const r = api.suggestEmail(
            { foreName: 'Anna-Sophie', longName: 'Bindlehner' },
            { domain: 'hak-steyr.at', pattern: 'vorname.nachname' }
        );
        expect(r.email).toBe('anna.bindlehner@hak-steyr.at');
        expect(r.firstNameMode).toBe('first');
    });

    it('kann alle Vornamen für die Mail verwenden', () => {
        const r = api.suggestEmail(
            { foreName: 'Anna-Sophie', longName: 'Bindlehner' },
            { domain: 'hak-steyr.at', pattern: 'vorname.nachname', firstNameMode: 'all' }
        );
        expect(r.email).toBe('anna.sophie.bindlehner@hak-steyr.at');
        expect(r.firstNameMode).toBe('all');
    });

    it('liefert Doppelname-Varianten', () => {
        const locals = api.localPartCandidates('Devran Eren', 'Acikdilli', 'vorname.nachname', 'first');
        expect(locals[0]).toBe('devran.acikdilli');
        expect(locals.some((x) => x.includes('devran'))).toBe(true);
        expect(locals.some((x) => x.includes('acikdilli'))).toBe(true);
        expect(locals.length).toBeGreaterThan(2);
    });

    it('behält vorhandene Mail', () => {
        const r = api.suggestEmail(
            { foreName: 'Anna', longName: 'Beispiel', email: 'anna@schule.at' },
            { domain: 'schule.at' }
        );
        expect(r.email).toBe('anna@schule.at');
        expect(r.generated).toBe(false);
    });
});

describe('webuntis-export-import', () => {
    let ctx;
    let wu;
    let sis;

    beforeEach(() => {
        ctx = loadAll();
        wu = ctx.ms365WebuntisExportImport;
        sis = ctx.ms365SchoolSisImport;
    });

    it('erkennt Student-/Guardian-/Teacher-Header', () => {
        expect(wu.detectExportKindFromAoa(loadJson('student-aoa.json'))).toBe('student');
        expect(wu.detectExportKindFromAoa(loadJson('guardian-aoa.json'))).toBe('guardian');
        expect(wu.detectExportKindFromAoa(loadJson('teacher-aoa.json'))).toBe('teacher');
    });

    it('parst Schüler und filtert Abgänger', () => {
        const students = wu.parseStudentsAoa(loadJson('student-aoa.json'));
        expect(students.some((s) => s.name.includes('Tom'))).toBe(false);
        expect(students.length).toBe(3);
        const anna = students.find((s) => s.untisInternalId === '1001');
        expect(anna.externalId).toBe('EXT1001');
        expect(anna.klasse).toBe('1A');
    });

    it('ordnet Eltern über Untis-IDs zu', () => {
        const result = wu.importStudentsFromWebuntis({
            studentAoa: loadJson('student-aoa.json'),
            guardianAoa: loadJson('guardian-aoa.json'),
            domain: 'schule.at',
            pattern: 'vorname.nachname',
            applyEmails: true
        });
        expect(result.meta.studentCount).toBe(3);
        const anna = result.records.find((s) => s.untisInternalId === '1001');
        expect(anna.parentPairs).toHaveLength(2);
        expect(anna.parentPairs.map((p) => p.email).sort()).toEqual([
            'maria.beispiel@mail.com',
            'thomas.beispiel@mail.com'
        ]);
        expect(result.matchCounts.byInternalId).toBeGreaterThanOrEqual(3);
        expect(result.unmatchedGuardians.length).toBe(1);
        expect(anna.email).toMatch(/@schule\.at$/);
    });

    it('parst Lehrer nur als Primärzeilen', () => {
        const teachers = wu.parseTeachersAoa(loadJson('teacher-aoa.json'));
        expect(teachers.map((t) => t.code)).toEqual(['ALTH', 'ANDES']);
        expect(teachers[0].name).toBe('Lukas Althuber');
    });

    it('SIS erkennt WebUntis-Student-Header', () => {
        expect(sis.detectSourceFromAoa(loadJson('student-aoa.json'))).toBe('webuntis');
    });

    it('Diff matcht über externalId und behält Mail beim Merge', () => {
        const incoming = wu.importStudentsFromWebuntis({
            studentAoa: loadJson('student-aoa.json'),
            guardianAoa: loadJson('guardian-aoa.json'),
            applyEmails: false
        }).records;
        const existing = [
            {
                klasse: '1A',
                name: 'Anna Beispiel',
                email: 'anna.beispiel@schule.at',
                externalId: 'EXT1001',
                parentPairs: []
            }
        ];
        const withMail = sis.preferExistingStudentEmails(incoming, existing);
        const anna = withMail.find((s) => s.externalId === 'EXT1001');
        expect(anna.email).toBe('anna.beispiel@schule.at');
        const diff = sis.diffSisImport(existing, withMail);
        expect(diff.counts.added).toBe(2);
        expect(diff.updated.length + diff.unchanged.length).toBeGreaterThanOrEqual(1);
        const merged = sis.applySisImport(existing, withMail, { mode: 'merge' });
        const kept = merged.find((s) => s.externalId === 'EXT1001');
        expect(kept.email).toBe('anna.beispiel@schule.at');
        expect(kept.parentPairs.length).toBe(2);
    });

    it('recordsToSemicolonLines rundet externalId mit #id:', () => {
        const lines = sis.recordsToSemicolonLines([
            {
                klasse: '1A',
                name: 'Anna',
                email: 'a@schule.at',
                externalId: 'EXT1',
                parentPairs: [{ name: 'M', email: 'm@mail.com' }]
            }
        ]);
        expect(lines).toContain('#id:EXT1');
        expect(lines).toContain('m@mail.com');
    });

    it('parst WebUntis-Klassen-PDF-Text mit KV-Kürzel und Abschlussjahr', () => {
        const text = readFileSync(join(fixtures, 'class-pdf-text.txt'), 'utf8');
        const parsed = wu.parseClassesFromPdfText(text, { skipWithoutTeacher: true });
        expect(parsed.schoolYear.label).toBe('2026/2027');
        expect(parsed.meta.withTeacher).toBeGreaterThanOrEqual(30);
        const hak = parsed.classes.find((c) => c.code === '1AK');
        expect(hak.headCode).toBe('HAGAU');
        expect(hak.deptText).toBe('.hak');
        expect(hak.year).toBe('2031'); // 5-jährig, Stufe 1, Endjahr 2027
        const has = parsed.classes.find((c) => c.code === '1AS');
        expect(has.headCode).toBe('RAUNE');
        expect(has.year).toBe('2029'); // 3-jährig HAS
        expect(parsed.classes.some((c) => c.code === '1A')).toBe(false);
    });

    it('inferGraduationYear: 5. HAK-Klassen → Schuljahres-Endjahr', () => {
        expect(wu.inferGraduationYear('5AK', '.hak', 2027)).toBe('2027');
        expect(wu.inferGraduationYear('5AK', '', 2027)).toBe('2027');
        expect(wu.inferGraduationYear('5AK', '.has', 2027)).toBe('2027'); // K schlägt fälschliches HAS
        expect(wu.inferGraduationYear('5AS', '.has', 2027)).toBe(''); // HAS nur 3 Jahre
        expect(wu.inferGraduationYear('1AK', '.hak', 2027)).toBe('2031');
        expect(wu.inferGraduationYear('3AS', '', 2027)).toBe('2027');
    });

    it('reichert Klassen mit Lehrerliste an', () => {
        const text = readFileSync(join(fixtures, 'class-pdf-text.txt'), 'utf8');
        const result = wu.importClassesFromWebuntisPdf({
            text,
            teachers: [{ code: 'HAGAU', name: 'Hagara Ursula', email: 'ursula.hagara@schule.at' }],
            skipWithoutTeacher: true
        });
        const hak = result.classes.find((c) => c.code === '1AK');
        expect(hak.headName).toBe('Hagara Ursula');
        expect(hak.headEmail).toBe('ursula.hagara@schule.at');
        expect(result.lines).toContain('1AK;2031;');
        expect(result.meta.matched).toBeGreaterThanOrEqual(1);
    });

    it('parst Klassen-PDF-Wörter positionsbasiert', () => {
        const words = JSON.parse(readFileSync(join(fixtures, 'class-pdf-words.json'), 'utf8'));
        const parsed = wu.parseClassesFromPdfWords(words, { skipWithoutTeacher: true });
        expect(parsed.classes.find((c) => c.code === '2BS').headCode).toBe('RAKOW');
        expect(parsed.classes.find((c) => c.code === 'FS_BAFEP')).toBeFalsy();
    });
});
