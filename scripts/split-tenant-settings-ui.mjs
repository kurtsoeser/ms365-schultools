/**
 * Safe move-only split of tenant-settings-ui.js (Analyse 02 Phase C).
 *
 * - Helpers → tenant-settings-ui-utils.js (export)
 * - window.ms365* bridges → tenant-settings-ui-import.js
 * - Full bindUi closure → tenant-settings-ui-bind.js (unchanged body; uses imports)
 * - Thin entry → tenant-settings-ui.js
 *
 * Panel files are carved from bindUi as install*(ctx) modules with brace-depth
 * scope rewriting (regex-aware). If a panel fails syntax check, bind stays whole.
 */
import fs from 'fs';
import path from 'path';
import { createRequire } from 'module';
import { spawnSync } from 'child_process';

const ROOT = path.resolve('e:/MS365schule/src/shared');
const BAK = path.join(ROOT, 'tenant-settings-ui.js.bak-pre-split');
const lines = fs.readFileSync(BAK, 'utf8').split(/\r?\n/);

function slice0(a, b) {
    return lines.slice(a, b + 1).join('\n');
}
function unindent(text, n) {
    return text
        .split('\n')
        .map((l) => {
            let i = 0;
            while (i < n && l[i] === ' ') i++;
            return l.slice(i);
        })
        .join('\n');
}

// ---- Utils: all preamble functions (94–665 and 752–988) ----
const helperChunks = [unindent(slice0(93, 664), 4), unindent(slice0(751, 987), 4)];
const helpersText = helperChunks.join('\n\n');

const OUTER_UTILS = [];
for (const m of helpersText.matchAll(/^function ([A-Za-z_$][\w$]*)\s*\(/gm)) {
    OUTER_UTILS.push(m[1]);
}
// also async function if any
for (const m of helpersText.matchAll(/^async function ([A-Za-z_$][\w$]*)\s*\(/gm)) {
    OUTER_UTILS.push(m[1]);
}

const utilsHeader = `/**
 * Shared helpers for tenant-settings-ui (dialogs, norm*, spreadsheet IO).
 * Extracted move-only from tenant-settings-ui.js (Analyse 02 Phase C).
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
    subjectCatalogShrinkReport,
    formatSubjectCatalogShrinkConfirmDe,
    subjectsFromImportPayload
} from './subject-list-enrich.js';

`;

fs.writeFileSync(
    path.join(ROOT, 'tenant-settings-ui-utils.js'),
    utilsHeader + helpersText + `\n\nexport {\n${[...new Set(OUTER_UTILS)].map((n) => `    ${n}`).join(',\n')}\n};\n`
);
console.log('utils exports', OUTER_UTILS.length);

// ---- Import APIs: window blocks only ----
const windowChunks = [unindent(slice0(665, 750), 4), unindent(slice0(989, 1157), 4)];

const importMod = `/**
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
    subjectListTemplateAoa
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
${windowChunks
    .join('\n\n')
    .split('\n')
    .map((l) => '    ' + l)
    .join('\n')}
}
`;
fs.writeFileSync(path.join(ROOT, 'tenant-settings-ui-import.js'), importMod);

// ---- Bind: keep full original bindUi body (lines 1160–11002) with utils as module imports ----
const originalImports = slice0(0, 89); // through schul-organigramm import

// Strip local helper defs from IIFE – bindUi starts at 1160; use body as-is but wrap
// Original: (function(){ helpers... bindUi... bootstrap })()
// New bind file: imports + installApis + function bindUi { ...original... } + whenReady

const bindUiBlock = unindent(slice0(1159, 11001), 4); // includes `function bindUi() { ... }`
const bootstrapBlock = unindent(slice0(11003, 11027), 4); // bindUiWhenReady bits without closing })();

// Replace bootstrap to call exported bindUi – keep logic identical
const bindFile = `${originalImports}
import * as tenantSettingsUiUtils from './tenant-settings-ui-utils.js';
import { installTenantSettingsImportApis } from './tenant-settings-ui-import.js';

const {
${[...new Set(OUTER_UTILS)].map((n) => `    ${n}`).join(',\n')}
} = tenantSettingsUiUtils;

installTenantSettingsImportApis();

${bindUiBlock}

${bootstrapBlock.replace(/bindUi\(\)/g, 'bindUi()')}

export { bindUi, bindUiWhenReady };
`;

// Fix bootstrap: original references bindUi in closure – still works
// Original ends with if (!bindUiWhenReady()) { ... } – good

fs.writeFileSync(path.join(ROOT, 'tenant-settings-ui-bind.js'), bindFile + '\n');

const entry = `/**
 * Tenant-Einstellungen UI – Entry (Analyse 02 Phase C / Tabs-Split).
 * Thin bootstrap; implementation lives in tenant-settings-ui-bind.js + panel/helpers modules.
 */
import { bindUiWhenReady } from './tenant-settings-ui-bind.js';

bindUiWhenReady();
`;
fs.writeFileSync(path.join(ROOT, 'tenant-settings-ui.js'), entry);

// Remove stale broken panel files from failed rewrite (keep only if we regenerate later)
const stale = [
    'tenant-settings-ui-bind-setup.js',
    'tenant-settings-ui-arges.js',
    'tenant-settings-ui-status.js',
    'tenant-settings-ui-directory.js',
    'tenant-settings-ui-subjects-panel.js',
    'tenant-settings-ui-teachers.js',
    'tenant-settings-ui-admin.js',
    'tenant-settings-ui-sga.js',
    'tenant-settings-ui-student-council.js',
    'tenant-settings-ui-school-groups.js',
    'tenant-settings-ui-teachers-lifecycle.js',
    'tenant-settings-ui-import-flows.js',
    'tenant-settings-ui-students-panel.js',
    'tenant-settings-ui-classes-panel.js',
    'tenant-settings-ui-render.js',
    'tenant-settings-ui-bind-events.js'
];
for (const f of stale) {
    const p = path.join(ROOT, f);
    if (fs.existsSync(p)) fs.unlinkSync(p);
}

console.log('entry lines', entry.split('\n').length);
console.log('bind lines', bindFile.split('\n').length);
console.log('utils lines', fs.readFileSync(path.join(ROOT, 'tenant-settings-ui-utils.js'), 'utf8').split('\n').length);
console.log('import lines', importMod.split('\n').length);

// syntax check
for (const f of ['tenant-settings-ui-utils.js', 'tenant-settings-ui-import.js', 'tenant-settings-ui-bind.js', 'tenant-settings-ui.js']) {
    const r = spawnSync(process.execPath, ['--check', path.join(ROOT, f)], { encoding: 'utf8' });
    if (r.status !== 0) {
        console.error('SYNTAX FAIL', f, r.stderr || r.stdout);
        process.exit(1);
    }
    console.log('syntax OK', f);
}
console.log('done');
