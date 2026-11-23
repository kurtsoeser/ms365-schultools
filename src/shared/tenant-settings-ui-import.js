/**
 * window.ms365* list-import / file-IO bridges for tenant-settings-ui.
 * Extracted move-only (Analyse 02 Phase C).
 */
import {
    CLASS_LIST_SHEET_NAME,
    CLASS_LIST_XLSX_FILENAME,
    classListTemplateAoa,
    classesRowsToAoa,
    downloadClassesCsv
} from './tenant-settings-ui-classes.js';
import {
    SUBJECT_LIST_SHEET_NAME,
    SUBJECT_LIST_XLSX_FILENAME,
    subjectListTemplateAoa,
    subjectsFromSpreadsheetJsonRows
} from './tenant-settings-ui-subjects.js';
import {
    ensureXlsxReady,
    downloadXlsxTemplate,
    downloadXlsxMultiSheet,
    readFilesToAoaSheets,
    readFilesToAllSheets,
    importStudentSheetsResult,
    teacherJsonRowsToSemicolonLines,
    teachersRowsToAoa,
    downloadTeachersCsv,
    sortTeacherRows,
    schoolDomainForEmail,
    tenantSchoolEmailPatternId,
    tenantSchoolEmailFirstNameMode,
    studentJsonRowsToSemicolonLines,
    studentsRowsToAoa,
    downloadStudentsCsv,
    sortStudentRows,
    adminJsonRowsToSemicolonLines,
    subjectsJsonRowsToSemicolonLines,
    argesJsonRowsToSemicolonLines,
    classesJsonRowsToSemicolonLines,
    findWorksheetName,
    importFileToWorkbook,
    csvTextToWorkbook,
    normStr
} from './tenant-settings-ui-utils.js';

export function installTenantSettingsImportApis() {
    window.ms365SubjectListImport = {
        isXlsxReady: ensureXlsxReady,
        templateAoa: subjectListTemplateAoa,
        rowsFromJson: subjectsFromSpreadsheetJsonRows,
        downloadTemplate() {
            return downloadXlsxTemplate(SUBJECT_LIST_XLSX_FILENAME, subjectListTemplateAoa(), SUBJECT_LIST_SHEET_NAME);
        }
    };
    
    window.ms365ClassListImport = {
        isXlsxReady: ensureXlsxReady,
        templateAoa: classListTemplateAoa,
        downloadTemplate() {
            return downloadXlsxTemplate(CLASS_LIST_XLSX_FILENAME, classListTemplateAoa(), CLASS_LIST_SHEET_NAME);
        },
        exportCsv(rows, filename) {
            return downloadClassesCsv(filename || 'Klassenliste.csv', rows);
        },
        exportXlsx(rows, filename) {
            return downloadXlsxTemplate(filename || 'Klassenliste.xlsx', classesRowsToAoa(rows), CLASS_LIST_SHEET_NAME);
        }
    };
    
    window.ms365TeacherListImport = {
        isXlsxReady: ensureXlsxReady,
        sortRows: sortTeacherRows,
        downloadTemplate() {
            return downloadXlsxTemplate(
                'Lehrerliste-Vorlage.xlsx',
                [
                    ['Kürzel', 'Name', 'E-Mail'],
                    ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                    ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                ],
                'Lehrer'
            );
        },
        downloadCsvTemplate() {
            return downloadTeachersCsv('Lehrerliste-Vorlage.csv', [
                { code: 'MU', name: 'Max Mustermann', email: 'max.mustermann@schule.de' }
            ]);
        },
        exportCsv(rows, filename) {
            return downloadTeachersCsv(filename || 'Lehrerliste.csv', rows);
        },
        exportXlsx(rows, filename) {
            return downloadXlsxTemplate(filename || 'Lehrerliste.xlsx', teachersRowsToAoa(rows), 'Lehrer');
        },
        importFile(file, onLines, onError) {
            if (!file) return;
            readFilesToAoaSheets([file])
                .then(function (sheets) {
                    const wu = window.ms365WebuntisExportImport;
                    const aoa = sheets[0] && sheets[0].aoa;
                    const kind =
                        wu && aoa && typeof wu.detectExportKindFromAoa === 'function' ? wu.detectExportKindFromAoa(aoa) : '';
                    if (kind === 'teacher' && wu.importTeachersFromWebuntis) {
                        const domain = schoolDomainForEmail();
                        const pattern = tenantSchoolEmailPatternId();
                        const firstNameMode = tenantSchoolEmailFirstNameMode();
                        const result = wu.importTeachersFromWebuntis({
                            teacherAoa: aoa,
                            domain: domain,
                            pattern: pattern,
                            firstNameMode: firstNameMode,
                            applyEmails: !!domain
                        });
                        if (onLines) onLines(result.lines, result);
                        return;
                    }
                    const rows = sheets[0]
                        ? XLSX.utils.sheet_to_json(
                              XLSX.utils.aoa_to_sheet(sheets[0].aoa),
                              { defval: '' }
                          )
                        : [];
                    // TeacherSalary-ähnliche Header ohne Erkennung: Fallback JSON
                    if (onLines) onLines(teacherJsonRowsToSemicolonLines(rows));
                })
                .catch(function (err) {
                    if (onError) onError('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)));
                });
        }
    };
    
    /** Schüler-Zeilen: Klasse;Name;E-Mail[;Eltern1;Eltern1Mail;…] */
    
    window.ms365SpreadsheetCsv = {
        csvTextToWorkbook: csvTextToWorkbook
    };
    
    window.ms365StammdatenFileIO = {
        readFilesToAoaSheets: readFilesToAoaSheets,
        readFilesToAllSheets: readFilesToAllSheets
    };
    
    window.ms365StudentListImport = {
        isXlsxReady: ensureXlsxReady,
        sortRows: sortStudentRows,
        downloadTemplate() {
            if (window.ms365SchoolSisImport && typeof window.ms365SchoolSisImport.downloadXlsxTemplates === 'function') {
                if (window.ms365SchoolSisImport.downloadXlsxTemplates()) return true;
            }
            return downloadXlsxTemplate(
                'Schuelerliste-Vorlage.xlsx',
                [
                    ['Klasse', 'Name', 'E-Mail', 'Eltern1', 'Eltern1Mail', 'Eltern2', 'Eltern2Mail'],
                    ['1AK', 'Lisa Beispiel', 'lisa.beispiel@schule.de', 'Maria Beispiel', 'maria@mail.com', '', ''],
                    ['1AK', 'Max Muster', 'max.muster@schule.de', 'Eva Muster', 'eva@mail.com', 'Tom Muster', 'tom@mail.com']
                ],
                'Schueler'
            );
        },
        downloadCsvTemplate() {
            if (window.ms365SchoolSisImport && typeof window.ms365SchoolSisImport.downloadCsv === 'function') {
                return window.ms365SchoolSisImport.downloadCsv(
                    'Schueler-Eltern-Vorlage.csv',
                    window.ms365SchoolSisImport.ms365TemplateAoa()
                );
            }
            return downloadStudentsCsv('Schuelerliste-Vorlage.csv', [
                {
                    klasse: '1AK',
                    name: 'Lisa Beispiel',
                    email: 'lisa.beispiel@schule.de',
                    parentPairs: [{ name: 'Maria Beispiel', email: 'maria@mail.com' }]
                }
            ]);
        },
        exportCsv(rows, filename) {
            return downloadStudentsCsv(filename || 'Schuelerliste.csv', rows);
        },
        exportXlsx(rows, filename) {
            return downloadXlsxTemplate(filename || 'Schuelerliste.xlsx', studentsRowsToAoa(rows), 'Schueler');
        },
        importFile(file, onLines, onError, sourceHint, existingStudents) {
            const files = file && file.length != null ? file : file ? [file] : [];
            if (!files.length) return;
            if (!ensureXlsxReady()) {
                if (onError) onError('Import: Excel-Bibliothek nicht geladen – Seite neu laden.');
                return;
            }
            const pattern = tenantSchoolEmailPatternId();
            const firstNameMode = tenantSchoolEmailFirstNameMode();
            readFilesToAoaSheets(files)
                .then(function (sheets) {
                    const result = importStudentSheetsResult(
                        sheets,
                        sourceHint || 'auto',
                        existingStudents || [],
                        pattern,
                        firstNameMode
                    );
                    if (onLines) onLines(result.lines, result);
                })
                .catch(function (err) {
                    if (onError) onError('Import fehlgeschlagen: ' + (err && err.message ? err.message : String(err)));
                });
        },
        importFiles(files, onLines, onError, sourceHint, existingStudents) {
            return this.importFile(files, onLines, onError, sourceHint, existingStudents);
        }
    };
    
    window.ms365SchuldatenMasterImport = {
        isXlsxReady: ensureXlsxReady,
        downloadTemplate() {
            return downloadXlsxMultiSheet('MS365-Schuldaten-Vorlage.xlsx', [
                {
                    name: 'Verwaltung',
                    aoa: [
                        ['Rolle', 'Name', 'E-Mail'],
                        ['Direktion', 'Direktorin Beispiel', 'direktion@schule.de'],
                        ['Sekretariat', 'Sekretariat', 'sekretariat@schule.de']
                    ]
                },
                {
                    name: 'Lehrer',
                    aoa: [
                        ['Kürzel', 'Name', 'E-Mail'],
                        ['MU', 'Max Mustermann', 'max.mustermann@schule.de'],
                        ['BME', 'Anna Beispiel', 'anna.beispiel@schule.de']
                    ]
                },
                {
                    name: 'Schueler',
                    aoa: [
                        ['Klasse', 'Name', 'E-Mail', 'Eltern1', 'Eltern1Mail', 'Eltern2', 'Eltern2Mail'],
                        ['1AK', 'Lisa Beispiel', 'lisa.beispiel@schule.de', 'Maria Beispiel', 'maria@mail.com', '', ''],
                        ['1AK', 'Max Muster', 'max.muster@schule.de', 'Eva Muster', 'eva@mail.com', 'Tom Muster', 'tom@mail.com']
                    ]
                },
                {
                    name: 'Faecher',
                    aoa: subjectListTemplateAoa()
                },
                {
                    name: 'ARGE',
                    aoa: [
                        ['Kürzel', 'Name', 'Fächer'],
                        ['SPRACHEN', 'Sprachen', 'D,E'],
                        ['NAWI', 'Naturwissenschaften', 'BIO,CH,PH']
                    ]
                },
                {
                    name: 'Klassen',
                    aoa: classListTemplateAoa()
                }
            ]);
        },
        importFile(file, onPayload, onError) {
            const name = String(file.name || '').toLowerCase();
            if (name.endsWith('.csv')) {
                if (onError) onError('Gesamt-Import: Bitte die XLSX-Vorlage verwenden (mehrere Arbeitsblätter).');
                return;
            }
            importFileToWorkbook(
                file,
                (wb) => {
                    try {
                        const out = {};
                        const snV = findWorksheetName(wb, ['verwaltung', 'administration']);
                        const snL = findWorksheetName(wb, ['lehrer', 'lehrerinnen', 'teachers']);
                        const snS = findWorksheetName(wb, ['schueler', 'schuler', 'schüler', 'students', 'schuelerinnen']);
                        const snF = findWorksheetName(wb, ['faecher', 'fächer', 'subjects']);
                        const snA = findWorksheetName(wb, ['arge', 'arbeitsgruppen', 'arbeitsgemeinschaften']);
                        const snK = findWorksheetName(wb, ['klassen', 'classes']);
                        function sheetLines(sn, conv) {
                            if (!sn || !wb.Sheets[sn]) return null;
                            const rows = XLSX.utils.sheet_to_json(wb.Sheets[sn], { defval: '' });
                            if (!rows || !rows.length) return null;
                            const lines = conv(rows);
                            return normStr(lines).length ? lines : null;
                        }
                        const v = sheetLines(snV, adminJsonRowsToSemicolonLines);
                        if (v != null) out.verwaltungLines = v;
                        const l = sheetLines(snL, teacherJsonRowsToSemicolonLines);
                        if (l != null) out.lehrerLines = l;
                        const s = sheetLines(snS, studentJsonRowsToSemicolonLines);
                        if (s != null) out.schuelerLines = s;
                        const f = sheetLines(snF, subjectsJsonRowsToSemicolonLines);
                        if (f != null) out.faecherLines = f;
                        const a = sheetLines(snA, argesJsonRowsToSemicolonLines);
                        if (a != null) out.argeLines = a;
                        const k = sheetLines(snK, classesJsonRowsToSemicolonLines);
                        if (k != null) out.klassenLines = k;
                        if (onPayload) onPayload(out);
                    } catch (err) {
                        if (onError) onError(String(err?.message || err));
                    }
                },
                onError
            );
        }
    };
    
}
