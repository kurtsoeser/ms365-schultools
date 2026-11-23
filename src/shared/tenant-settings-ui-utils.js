/**
 * Shared helpers for tenant-settings-ui (dialogs, norm*, spreadsheet IO).
 * Barrel re-export (R20/c6) — implementations live in tenant-settings-ui-utils-*.js.
 */
export {
    normStr,
    dlgAlert,
    dlgConfirm,
    dlgPrompt,
    normCode,
    escapeHtml,
    safeJsonParse,
    normHeaderKey,
    getField
} from './tenant-settings-ui-utils-dialog.js';

export {
    ensureXlsxReady,
    sheetToJsonRows,
    workbookFirstSheetAoa,
    aoaMaxColumnCount,
    csvTextToWorkbook,
    parseCsvTextToJsonRows,
    downloadXlsxTemplate,
    downloadXlsxMultiSheet,
    importSpreadsheetFileToJsonRows,
    findWorksheetName,
    importFileToWorkbook,
    workbookToSheetList,
    readFilesToAllSheets
} from './tenant-settings-ui-utils-spreadsheet.js';

export {
    exposeSubjectCatalogGuardOnWindow,
    confirmSubjectCatalogIfShrinks,
    teacherJsonRowsToSemicolonLines,
    teachersRowsToAoa,
    csvEscapeField,
    downloadTeachersCsv,
    compareTeachersDe,
    sortTeacherRows
} from './tenant-settings-ui-utils-teachers-lines.js';

export {
    schoolDomainForEmail,
    emailPatternFromSelect,
    emailFirstNameModeFromSelect,
    emailPatternFromLocalStorageFallback,
    emailFirstNameModeFromLocalStorageFallback,
    tenantSchoolEmailPatternId,
    tenantSchoolEmailFirstNameMode
} from './tenant-settings-ui-utils-email-pattern.js';

export {
    readFilesToAoaSheets,
    importStudentSheetsResult,
    studentJsonRowsToSemicolonLines,
    studentsRowsToAoa,
    downloadStudentsCsv,
    sortStudentRows
} from './tenant-settings-ui-utils-students-lines.js';

export {
    adminJsonRowsToSemicolonLines,
    subjectsJsonRowsToSemicolonLines,
    argesJsonRowsToSemicolonLines,
    classesJsonRowsToSemicolonLines
} from './tenant-settings-ui-utils-list-lines.js';
