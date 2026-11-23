/**
 * Split tenant-settings-ui-utils.js into themed modules + thin barrel (R20/c6).
 */
import fs from 'fs';
import path from 'path';
import { spawnSync } from 'child_process';

const ROOT = path.resolve('e:/MS365schule/src/shared');
const SRC = path.join(ROOT, 'tenant-settings-ui-utils.js');
const lines = fs.readFileSync(SRC, 'utf8').split(/\r?\n/);

function writeCheck(file, text) {
    const p = path.join(ROOT, file);
    fs.writeFileSync(p, text.endsWith('\n') ? text : text + '\n');
    const check = spawnSync(process.execPath, ['--check', p], { encoding: 'utf8' });
    console.log(file, text.split(/\r?\n/).length, check.status === 0 ? 'OK' : 'FAIL');
    if (check.status !== 0) {
        console.error(check.stderr || check.stdout);
        process.exit(1);
    }
}

function slice1(a, b) {
    return lines.slice(a - 1, b).join('\n');
}

// Original line ranges (1-based inclusive) from current utils.js
// Header imports: 1-22
const HEADER_IMPORTS = `import {
    subjectCatalogShrinkReport,
    formatSubjectCatalogShrinkConfirmDe,
    subjectsFromImportPayload
} from './subject-list-enrich.js';
`;

// dialog: functions lines 24-108
const dialogBody = slice1(24, 108);
writeCheck(
    'tenant-settings-ui-utils-dialog.js',
    `/**
 * Dialog / norm / field helpers for tenant-settings-ui.
 */
${dialogBody}

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
};
`
);

// spreadsheet: 110-231 + 738-833 (workbook helpers after list-lines)
const spreadsheetEarly = slice1(110, 231);
const spreadsheetLate = slice1(738, 833);
writeCheck(
    'tenant-settings-ui-utils-spreadsheet.js',
    `/**
 * Spreadsheet / XLSX / CSV IO helpers for tenant-settings-ui.
 */
${spreadsheetEarly}

${spreadsheetLate}

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
};
`
);

// teachers-lines + subject catalog: 233-351
const teachersBody = slice1(233, 351);
writeCheck(
    'tenant-settings-ui-utils-teachers-lines.js',
    `/**
 * Teacher list / subject-catalog helpers for tenant-settings-ui.
 */
${HEADER_IMPORTS}
import {
    dlgConfirm,
    getField,
    normCode,
    normStr
} from './tenant-settings-ui-utils-dialog.js';

${teachersBody}

export {
    exposeSubjectCatalogGuardOnWindow,
    confirmSubjectCatalogIfShrinks,
    teacherJsonRowsToSemicolonLines,
    teachersRowsToAoa,
    csvEscapeField,
    downloadTeachersCsv,
    compareTeachersDe,
    sortTeacherRows
};
`
);

// email-pattern: 353-452
const emailBody = slice1(353, 452);
writeCheck(
    'tenant-settings-ui-utils-email-pattern.js',
    `/**
 * School e-mail pattern helpers for tenant-settings-ui.
 */
import { normStr } from './tenant-settings-ui-utils-dialog.js';

${emailBody}

export {
    schoolDomainForEmail,
    emailPatternFromSelect,
    emailFirstNameModeFromSelect,
    emailPatternFromLocalStorageFallback,
    emailFirstNameModeFromLocalStorageFallback,
    tenantSchoolEmailPatternId,
    tenantSchoolEmailFirstNameMode
};
`
);

// students: 454-685
const studentsBody = slice1(454, 685);
writeCheck(
    'tenant-settings-ui-utils-students-lines.js',
    `/**
 * Student spreadsheet / lines helpers for tenant-settings-ui.
 */
import { getField, normCode, normStr } from './tenant-settings-ui-utils-dialog.js';
import {
    ensureXlsxReady,
    workbookFirstSheetAoa,
    csvTextToWorkbook
} from './tenant-settings-ui-utils-spreadsheet.js';
import { csvEscapeField, compareTeachersDe } from './tenant-settings-ui-utils-teachers-lines.js';
import { schoolDomainForEmail } from './tenant-settings-ui-utils-email-pattern.js';

${studentsBody}

export {
    readFilesToAoaSheets,
    importStudentSheetsResult,
    studentJsonRowsToSemicolonLines,
    studentsRowsToAoa,
    downloadStudentsCsv,
    sortStudentRows
};
`
);

// list-lines: 687-736
const listBody = slice1(687, 736);
writeCheck(
    'tenant-settings-ui-utils-list-lines.js',
    `/**
 * Admin / subjects / ARGEs / classes JSON→lines helpers.
 */
import { getField, normCode, normStr } from './tenant-settings-ui-utils-dialog.js';

${listBody}

export {
    adminJsonRowsToSemicolonLines,
    subjectsJsonRowsToSemicolonLines,
    argesJsonRowsToSemicolonLines,
    classesJsonRowsToSemicolonLines
};
`
);

// barrel
writeCheck(
    'tenant-settings-ui-utils.js',
    `/**
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
`
);

console.log('utils split done');
