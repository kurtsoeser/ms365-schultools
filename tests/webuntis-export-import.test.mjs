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
        expect(wu.detectExportKindFromAoa(loadJson('teacher-stammdaten-aoa.json'))).toBe('teacher');
    });

    it('parst Student_*.csv (WebUntis Stammdaten, tab)', () => {
        const students = wu.parseStudentsAoa(loadJson('student-stammdaten-aoa.json'));
        expect(students.length).toBe(4);
        const amir = students.find((s) => s.untisInternalId === '8613');
        expect(amir.klasse).toBe('2BS');
        expect(amir.name).toBe('Amirhan Abubakarov');
        expect(amir.phone).toContain('677');
        const emin = students.find((s) => s.untisInternalId === '4937');
        expect(emin.email).toBe('emin.acikyuerek@hak-steyr.at');
        const noClass = students.find((s) => s.untisInternalId === '6894');
        expect(noClass.klasse).toBe('');
        expect(wu.detectExportKindFromAoa(loadJson('student-stammdaten-aoa.json'))).toBe('student');
    });

    it('parst Schüler und filtert Abgänger', () => {
        const students = wu.parseStudentsAoa(loadJson('student-aoa.json'));
        expect(students.some((s) => s.name.includes('Tom'))).toBe(false);
        expect(students.length).toBe(3);
        const anna = students.find((s) => s.untisInternalId === '1001');
        expect(anna.externalId).toBe('EXT1001');
        expect(anna.klasse).toBe('1A');
    });

    it('parst LegalGuardian_*.csv (WebUntis Stammdaten)', () => {
        const rows = wu.parseGuardiansAoa(loadJson('legal-guardian-stammdaten-aoa.json'));
        expect(rows.length).toBe(3);
        const both = rows.find((g) => g.email === 'spanring@reload.co.at');
        expect(both.studentInternalId).toBe('6165');
        expect(both.name).toBe('Mathias Spanring');
        const singleName = rows.find((g) => g.email === 'dario_glavas91@gmx.at');
        expect(singleName.name).toContain('Glavas');
        const orphan = rows.find((g) => g.email === 'sonja.hinterleitner@gmail.com');
        expect(orphan.studentInternalId).toBe('');
        expect(orphan.phone).toMatch(/0699|4369910758267/);
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

    it('parst WebUntis Subject_*.pdf (Langname+Kurzname zusammengeklebt)', () => {
        const text = readFileSync(join(fixtures, 'subject-pdf-text.txt'), 'utf8');
        const parsed = wu.parseSubjectsFromPdfText(text);
        expect(parsed.meta.subjectCount).toBeGreaterThan(35);
        const deutsch = parsed.subjects.find((s) => s.code === 'D');
        expect(deutsch).toBeTruthy();
        expect(deutsch.name).toMatch(/DEUTSCH/i);
        expect(parsed.subjects.some((s) => s.admin)).toBe(true);
        const lines = wu.subjectsToSemicolonLines(parsed.subjects.slice(0, 2));
        expect(lines).toContain(';');
    });

    it('parst Subject-PDF Zeilenpaar (Langname / Kurzname getrennt)', () => {
        const text = [
            '50 ADMINISTRATOR',
            'ADM',
            '03 DEUTSCH',
            'D',
            '01 ETHIK',
            'ETH',
            '17 CHEMIE',
            'CH'
        ].join('\n');
        const parsed = wu.parseSubjectsFromPdfText(text);
        expect(parsed.subjects.find((s) => s.code === 'D').name).toMatch(/DEUTSCH/);
        expect(parsed.subjects.find((s) => s.code === 'ETH').name).toBe('ETHIK');
        expect(parsed.subjects.find((s) => s.code === 'ETH').curriculumNo).toBe('01');
        expect(parsed.subjects.find((s) => s.code === 'ADM').name).toMatch(/ADMINISTRATOR/);
        expect(parsed.subjects.find((s) => s.code === 'K')).toBeFalsy();
    });

    it('parst Subject-PDF mit Kurzname-Spalte (positionsbasiert)', () => {
        const words = [
            { str: 'ADM', x: 30, y: 138, page: 1 },
            { str: '50', x: 90, y: 138, page: 1 },
            { str: 'ADMINISTRATOR', x: 102, y: 138, page: 1 },
            { str: 'D', x: 30, y: 632, page: 1 },
            { str: '03', x: 90, y: 632, page: 1 },
            { str: 'DEUTSCH', x: 102, y: 632, page: 1 },
            { str: 'ETH', x: 30, y: 287, page: 1 },
            { str: '01', x: 90, y: 287, page: 1 },
            { str: 'ETHIK', x: 102, y: 287, page: 1 }
        ];
        const parsed = wu.parseSubjectsFromPdfWords(words);
        expect(parsed.subjects.find((s) => s.code === 'D').name).toBe('DEUTSCH');
        expect(parsed.subjects.find((s) => s.code === 'ETH').name).toBe('ETHIK');
        expect(parsed.subjects.find((s) => s.code === 'D').curriculumNo).toBe('03');
        const imp = wu.importSubjectsFromWebuntisPdf({ words, text: '' });
        expect(imp.lines).toContain('D;DEUTSCH');
    });

    it('Subject-PDF-Wörter: gleiche Y auf verschiedenen Seiten nicht zusammenführen', () => {
        const words = [
            { str: 'D1', x: 30, y: 647, page: 1 },
            { str: '03', x: 90, y: 647, page: 1 },
            { str: 'DEUTSCH', x: 102, y: 647, page: 1 },
            { str: 'IMEDIA', x: 30, y: 647, page: 2 },
            { str: 'Interaktive', x: 90, y: 647, page: 2 },
            { str: 'Medien', x: 150, y: 647, page: 2 }
        ];
        const parsed = wu.parseSubjectsFromPdfWords(words);
        expect(parsed.subjects.map((s) => s.code).sort()).toEqual(['D1', 'IMEDIA']);
    });

    it('parst WebUntis Teacher_*.csv (Stammdaten name/longName/foreName)', () => {
        const teachers = wu.parseTeachersAoa(loadJson('teacher-stammdaten-aoa.json'));
        expect(teachers.map((t) => t.code)).toEqual(['ALTH', 'SOESER']);
        expect(teachers[0].email).toBe('lukas.althuber@hak-steyr.at');
        expect(teachers[1].name).toBe('Kurt Söser');
    });

    it('SIS erkennt WebUntis-Student-Header', () => {
        expect(sis.detectSourceFromAoa(loadJson('student-aoa.json'))).toBe('webuntis');
    });

    it('ignoriert private Schüler-Mail aus WebUntis und erzeugt Schuladresse', () => {
        const wu = loadAll().ms365WebuntisExportImport;
        const aoa = loadJson('student-aoa.json');
        aoa[1][14] = 'anna.privat@gmail.com';
        const imp = wu.importStudentsFromWebuntis({
            studentAoa: aoa,
            guardianAoa: [],
            domain: 'schule.at',
            pattern: 'vorname.nachname',
            firstNameMode: 'first',
            applyEmails: true
        });
        expect(imp.meta.privateStudentEmailsIgnored).toBe(1);
        const anna = imp.records.find((s) => s.externalId === 'EXT1001');
        expect(anna.email).toBe('anna.beispiel@schule.at');
        expect(anna.email).not.toContain('gmail');
    });

    it('behält Schüler-Mail nur auf der konfigurierten Schuldomain', () => {
        const wu = loadAll().ms365WebuntisExportImport;
        expect(wu.sanitizeStudentEmailFromWebuntis('max@gmx.at', 'hak-steyr.at')).toBe('');
        expect(wu.sanitizeStudentEmailFromWebuntis('max.mustermann@hak-steyr.at', 'hak-steyr.at')).toBe(
            'max.mustermann@hak-steyr.at'
        );
        expect(wu.sanitizeStudentEmailFromWebuntis('privat@web.de', '')).toBe('');
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

    it('nutzt Schuljahr-Fallback für Abschlussjahr wenn PDF kein Schuljahr hat', () => {
        const text = readFileSync(join(fixtures, 'class-pdf-text.txt'), 'utf8');
        const ohneSj = text.replace(/Schuljahr\s*:?\s*2026\/2027/i, '');
        const result = wu.importClassesFromWebuntisPdf({
            text: ohneSj,
            teachers: [],
            skipWithoutTeacher: true,
            schoolYearEnd: 2027
        });
        const hak = result.classes.find((c) => c.code === '1AK');
        expect(hak).toBeTruthy();
        expect(hak.year).toBe('2031');
        expect(result.schoolYear.endYear).toBe(2027);
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
