/**
 * Split tenant-settings-ui-bind.js bindUi() into install*(ctx) panel modules.
 * Utils/imports stay module-scoped (not rewritten onto ctx).
 */
import fs from 'fs';
import path from 'path';
import { spawnSync } from 'child_process';

const ROOT = path.resolve('e:/MS365schule/src/shared');
const BIND = path.join(ROOT, 'tenant-settings-ui-bind.js');
const src = fs.readFileSync(BIND, 'utf8');
const lines = src.split(/\r?\n/);

// Find bindUi
const bindUiLine = lines.findIndex((l) => /^function bindUi\(\)/.test(l));
if (bindUiLine < 0) throw new Error('bindUi not found');

/** Scan code with regex/string/comment awareness; invoke onBrace(ch, depthAfter) */
function scanBindUiBody(onLineAtDepth1, onDone) {
    const text = lines.slice(bindUiLine).join('\n');
    const braceOpen = text.indexOf('{');
    let d = 0;
    let state = 'code';
    let templateExprDepth = 0;
    for (let p = braceOpen; p < text.length; p++) {
        const ch = text[p];
        const next = text[p + 1];
        if (state === 'linecomment') {
            // Newline ends the comment; fall through so depth-1 decl scan sees the next line.
            if (ch === '\n') state = 'code';
            else continue;
        }
        if (state === 'blockcomment') {
            if (ch === '*' && next === '/') {
                state = 'code';
                p++;
            }
            continue;
        }
        if (state === 'squote') {
            if (ch === '\\') {
                p++;
                continue;
            }
            if (ch === "'") state = 'code';
            continue;
        }
        if (state === 'dquote') {
            if (ch === '\\') {
                p++;
                continue;
            }
            if (ch === '"') state = 'code';
            continue;
        }
        if (state === 'template') {
            if (ch === '\\') {
                p++;
                continue;
            }
            if (ch === '`') {
                state = 'code';
                continue;
            }
            if (ch === '$' && next === '{') {
                templateExprDepth = 1;
                state = 'code';
                p++;
                continue;
            }
            continue;
        }
        if (ch === '/' && next === '/') {
            state = 'linecomment';
            p++;
            continue;
        }
        if (ch === '/' && next === '*') {
            state = 'blockcomment';
            p++;
            continue;
        }
        if (ch === "'") {
            state = 'squote';
            continue;
        }
        if (ch === '"') {
            state = 'dquote';
            continue;
        }
        if (ch === '`') {
            state = 'template';
            continue;
        }
        if (templateExprDepth > 0) {
            if (ch === '{') templateExprDepth++;
            if (ch === '}') {
                templateExprDepth--;
                if (templateExprDepth === 0) state = 'template';
            }
            continue;
        }
        // regex
        if (ch === '/') {
            let k = p - 1;
            while (k >= 0 && /[ \t]/.test(text[k])) k--;
            const prev = k >= 0 ? text[k] : '';
            const isDivision = /[\w$)\]]/.test(prev);
            if (!isDivision) {
                p++;
                while (p < text.length) {
                    const c = text[p];
                    if (c === '\\') {
                        p += 2;
                        continue;
                    }
                    if (c === '[') {
                        p++;
                        while (p < text.length && text[p] !== ']') {
                            if (text[p] === '\\') p++;
                            p++;
                        }
                        p++;
                        continue;
                    }
                    if (c === '/') {
                        p++;
                        while (p < text.length && /[a-z]/i.test(text[p])) p++;
                        p--;
                        break;
                    }
                    p++;
                }
                continue;
            }
        }
        if (ch === '{') d++;
        else if (ch === '}') {
            d--;
            if (d === 0) {
                onDone(p);
                return;
            }
        }
        if (ch === '\n' && d === 1) {
            const lineStart = p + 1;
            const lineEnd = text.indexOf('\n', lineStart);
            const line = text.slice(lineStart, lineEnd < 0 ? text.length : lineEnd);
            onLineAtDepth1(line);
        }
    }
}

const scopeNames = new Set();
const functionNames = new Set();
scanBindUiBody((line) => {
    let m = line.match(/^(\s*)(async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/);
    if (m) {
        scopeNames.add(m[3]);
        functionNames.add(m[3]);
        return;
    }
    m = line.match(/^(\s*)(const|let)\s+([A-Za-z_$][\w$]*)\s*=/);
    if (m) scopeNames.add(m[3]);
}, () => {});

// Module-level helpers/imports used as free vars inside bindUi → must live on ctx
const MODULE_BINDINGS = [
    'normStr',
    'dlgAlert',
    'dlgConfirm',
    'dlgPrompt',
    'normCode',
    'escapeHtml',
    'safeJsonParse',
    'normHeaderKey',
    'getField',
    'ensureXlsxReady',
    'sheetToJsonRows',
    'workbookFirstSheetAoa',
    'aoaMaxColumnCount',
    'csvTextToWorkbook',
    'parseCsvTextToJsonRows',
    'downloadXlsxTemplate',
    'downloadXlsxMultiSheet',
    'importSpreadsheetFileToJsonRows',
    'exposeSubjectCatalogGuardOnWindow',
    'confirmSubjectCatalogIfShrinks',
    'teacherJsonRowsToSemicolonLines',
    'teachersRowsToAoa',
    'csvEscapeField',
    'downloadTeachersCsv',
    'compareTeachersDe',
    'sortTeacherRows',
    'schoolDomainForEmail',
    'emailPatternFromSelect',
    'emailFirstNameModeFromSelect',
    'emailPatternFromLocalStorageFallback',
    'emailFirstNameModeFromLocalStorageFallback',
    'tenantSchoolEmailPatternId',
    'tenantSchoolEmailFirstNameMode',
    'readFilesToAoaSheets',
    'importStudentSheetsResult',
    'studentJsonRowsToSemicolonLines',
    'studentsRowsToAoa',
    'downloadStudentsCsv',
    'sortStudentRows',
    'adminJsonRowsToSemicolonLines',
    'subjectsJsonRowsToSemicolonLines',
    'argesJsonRowsToSemicolonLines',
    'classesJsonRowsToSemicolonLines',
    'findWorksheetName',
    'importFileToWorkbook',
    'workbookToSheetList',
    'readFilesToAllSheets',
    'CLASS_LIST_SHEET_NAME',
    'CLASS_LIST_XLSX_FILENAME',
    'classListTemplateAoa',
    'classesRowsToAoa',
    'classesToLinesPure',
    'downloadClassesCsv',
    'studentsToLinesPure',
    'enrichStudentLinesWithDemoParents',
    'resolveStudentRepSyncSource',
    'sgaStudentsFromStudentCouncil',
    'studentCouncilFromSgaStudents',
    'SUBJECT_LIST_SHEET_NAME',
    'SUBJECT_LIST_XLSX_FILENAME',
    'subjectListTemplateAoa',
    'subjectsFromSpreadsheetJsonRows',
    'subjectsToLinesPure',
    'pdfImportTargetFromFilename',
    'resolveRegisterImportTarget',
    'spreadsheetImportTargetFromAoa',
    'REGISTER_IMPORT_TAB_BTN',
    'REGISTER_IMPORT_TARGET_LABEL',
    'openSubjectPdfImportReview',
    'ADMIN_TIER_SCHULLEITUNG',
    'ADMIN_TIER_VERWALTUNG',
    'inferAdminTierForRole',
    'normalizeAdminTier',
    'BUILTIN_AUDIENCE_SCHULLEITUNG',
    'ensureAdminAudienceOnSettings',
    'normalizeAdminAudienceMemberships',
    'normalizeVerwaltungAudienceGroups',
    'mountTenantAdminAudienceBoard',
    'mountVerwaltungStammdatenTree',
    'mountTenantSchoolEmailPatternUi',
    'mountTenantIntranetQuickSync',
    'mountTenantSchoolProfileUi',
    'reloadTenantSchoolProfileForm',
    'mountTenantStammdatenSetupNav',
    'mountTenantPlanerSitesUi',
    'writeSchoolData',
    'isStammdatenLocalFirst',
    'localFirstStatusMessage',
    'enrichClassesFromLinkedGroups',
    'enrichClassesKvFromGraphOwners',
    'collectPriorClassFieldsByCode',
    'deriveClassEntryFromGroup',
    'fillRegisterRowsHeadFromGroupOwners',
    'enrichSubjectsFromCatalogSources',
    'subjectCatalogShrinkReport',
    'formatSubjectCatalogShrinkConfirmDe',
    'subjectsFromImportPayload',
    'deriveSubjectEntryFromGroup',
    'pickEntraGroupsMulti',
    'pickEntraUser',
    'studentBelongsToClassRow',
    'summarizeWebuntisMergeResult',
    'diffTeachersImport',
    'summarizeTeachersDiff',
    'mergeTeachersImportLists',
    'replaceTeachersImportList',
    'mountTenantImportPanel',
    'mergeSisImportViaAdapter',
    'notifyTenantStudentsChanged',
    'applyTabRegisterStatus',
    'personMatchTabStatus',
    'classMatchTabStatus',
    'catalogLinkTabStatus',
    'updateRegisterLastSavedMeta',
    'formatRegisterLastSavedSuffix',
    'renderSchoolYearSelect',
    'bindSchoolYearControls',
    'summarizeClassTeaching',
    'mountKlassenlehrerPanel'
];
MODULE_BINDINGS.forEach((n) => scopeNames.add(n));

console.log('scope', scopeNames.size, 'fns', functionNames.size);
console.log(
    'short names',
    [...scopeNames].filter((n) => n.length <= 2).join(',') || '(none)'
);

function rewriteToCtx(src) {
    const out = [];
    let i = 0;
    let state = 'code';
    let templateExprDepth = 0;
    let inFnParams = false;
    let fnParamsParen = 0;
    /** @type {Array<'{'|'('|'['>} */
    const bracketStack = [];
    let lastEmitted = '';

    while (i < src.length) {
        const ch = src[i];
        const next = src[i + 1];
        if (state === 'linecomment') {
            out.push(ch);
            if (ch === '\n') state = 'code';
            i++;
            continue;
        }
        if (state === 'blockcomment') {
            out.push(ch);
            if (ch === '*' && next === '/') {
                out.push(next);
                i += 2;
                state = 'code';
                continue;
            }
            i++;
            continue;
        }
        if (state === 'squote') {
            out.push(ch);
            if (ch === '\\') {
                out.push(src[i + 1] || '');
                i += 2;
                continue;
            }
            if (ch === "'") state = 'code';
            i++;
            continue;
        }
        if (state === 'dquote') {
            out.push(ch);
            if (ch === '\\') {
                out.push(src[i + 1] || '');
                i += 2;
                continue;
            }
            if (ch === '"') state = 'code';
            i++;
            continue;
        }
        if (state === 'template') {
            out.push(ch);
            if (ch === '\\') {
                out.push(src[i + 1] || '');
                i += 2;
                continue;
            }
            if (ch === '`') {
                state = 'code';
                i++;
                continue;
            }
            if (ch === '$' && next === '{') {
                out.push('{');
                i += 2;
                templateExprDepth = 1;
                state = 'code';
                continue;
            }
            i++;
            continue;
        }
        if (ch === '/' && next === '/') {
            out.push(ch, next);
            i += 2;
            state = 'linecomment';
            continue;
        }
        if (ch === '/' && next === '*') {
            out.push(ch, next);
            i += 2;
            state = 'blockcomment';
            continue;
        }
        if (ch === "'") {
            out.push(ch);
            state = 'squote';
            i++;
            continue;
        }
        if (ch === '"') {
            out.push(ch);
            state = 'dquote';
            i++;
            continue;
        }
        if (ch === '`') {
            out.push(ch);
            state = 'template';
            i++;
            continue;
        }
        if (templateExprDepth > 0) {
            if (ch === '{') templateExprDepth++;
            if (ch === '}') {
                templateExprDepth--;
                out.push(ch);
                i++;
                if (templateExprDepth === 0) state = 'template';
                continue;
            }
        }
        // regex
        if (ch === '/') {
            let k = i - 1;
            while (k >= 0 && /[ \t]/.test(src[k])) k--;
            const prev = k >= 0 ? src[k] : '';
            const isDivision = /[\w$)\]]/.test(prev);
            if (!isDivision) {
                out.push('/');
                i++;
                while (i < src.length) {
                    const c = src[i];
                    out.push(c);
                    if (c === '\\') {
                        out.push(src[i + 1] || '');
                        i += 2;
                        continue;
                    }
                    if (c === '[') {
                        i++;
                        while (i < src.length && src[i] !== ']') {
                            if (src[i] === '\\') {
                                out.push(src[i], src[i + 1] || '');
                                i += 2;
                                continue;
                            }
                            out.push(src[i]);
                            i++;
                        }
                        if (i < src.length) {
                            out.push(src[i]);
                            i++;
                        }
                        continue;
                    }
                    if (c === '/') {
                        i++;
                        while (i < src.length && /[a-z]/i.test(src[i])) {
                            out.push(src[i]);
                            i++;
                        }
                        break;
                    }
                    i++;
                }
                continue;
            }
        }

        if (/[A-Za-z_$]/.test(ch)) {
            let j = i + 1;
            while (j < src.length && /[\w$]/.test(src[j])) j++;
            const ident = src.slice(i, j);
            let k = i - 1;
            while (k >= 0 && /[ \t]/.test(src[k])) k--;
            const prev = k >= 0 ? src[k] : '';
            const precededByDot = prev === '.';
            let n = j;
            while (n < src.length && /[ \t]/.test(src[n])) n++;
            const followedByColon = src[n] === ':' && src[n + 1] !== ':';
            let p = k;
            while (p >= 0 && /[ \t\n\r]/.test(src[p])) p--;
            const prevNonSpace = p >= 0 ? src[p] : '';
            const likelyObjectKey =
                followedByColon && (prevNonSpace === '{' || prevNonSpace === ',' || prevNonSpace === '(');
            // Object shorthand `{ scheduleAutoSave, setSummary }` → `scheduleAutoSave: ctx.scheduleAutoSave`
            let n2 = n;
            while (n2 < src.length && /[ \t\n\r]/.test(src[n2])) n2++;
            const afterIdent = src[n2];
            const inObject =
                bracketStack.length > 0 && bracketStack[bracketStack.length - 1] === '{';
            const likelyObjectShorthand =
                inObject &&
                !followedByColon &&
                (prevNonSpace === '{' || prevNonSpace === ',') &&
                (afterIdent === ',' || afterIdent === '}');
            const isFnDeclName = /\b(?:async\s+)?function\s*$/.test(lastEmitted);

            if (isFnDeclName || inFnParams || precededByDot || likelyObjectKey || !scopeNames.has(ident)) {
                out.push(ident);
                lastEmitted += ident;
            } else if (likelyObjectShorthand) {
                const exp = ident + ': ctx.' + ident;
                out.push(exp);
                lastEmitted += exp;
            } else {
                out.push('ctx.' + ident);
                lastEmitted += 'ctx.' + ident;
            }
            if (lastEmitted.length > 120) lastEmitted = lastEmitted.slice(-120);
            i = j;
            continue;
        }

        if (ch === '(') {
            if (/\bfunction\s+[A-Za-z_$][\w$]*\s*$/.test(lastEmitted) || /\bfunction\s*$/.test(lastEmitted)) {
                inFnParams = true;
                fnParamsParen = 1;
            } else if (inFnParams) fnParamsParen++;
            bracketStack.push('(');
            out.push(ch);
            lastEmitted += ch;
            i++;
            continue;
        }
        if (ch === ')') {
            if (inFnParams) {
                fnParamsParen--;
                if (fnParamsParen <= 0) {
                    inFnParams = false;
                    fnParamsParen = 0;
                }
            }
            if (bracketStack.length && bracketStack[bracketStack.length - 1] === '(') bracketStack.pop();
            out.push(ch);
            lastEmitted += ch;
            i++;
            continue;
        }
        if (ch === '{') {
            bracketStack.push('{');
            out.push(ch);
            lastEmitted += ch;
            i++;
            continue;
        }
        if (ch === '}') {
            if (bracketStack.length && bracketStack[bracketStack.length - 1] === '{') bracketStack.pop();
            out.push(ch);
            lastEmitted += ch;
            i++;
            continue;
        }
        if (ch === '[') {
            bracketStack.push('[');
            out.push(ch);
            lastEmitted += ch;
            i++;
            continue;
        }
        if (ch === ']') {
            if (bracketStack.length && bracketStack[bracketStack.length - 1] === '[') bracketStack.pop();
            out.push(ch);
            lastEmitted += ch;
            i++;
            continue;
        }
        out.push(ch);
        lastEmitted += ch;
        if (lastEmitted.length > 120) lastEmitted = lastEmitted.slice(-120);
        i++;
    }
    return out.join('');
}

function convertDecls(src) {
    return src
        .split('\n')
        .map((line) => {
            let m = line.match(/^(\s*)(async\s+)?function\s+([A-Za-z_$][\w$]*)\s*\(/);
            if (m && functionNames.has(m[3])) {
                return `${m[1]}ctx.${m[3]} = ${m[2] || ''}function ${m[3]}(` + line.slice(m[0].length);
            }
            m = line.match(/^(\s*)(const|let)\s+([A-Za-z_$][\w$]*)(\s*=)/);
            if (m && scopeNames.has(m[3])) {
                return `${m[1]}ctx.${m[3]}${m[4]}` + line.slice(m[0].length);
            }
            return line;
        })
        .join('\n');
}

// Map original bak line numbers → bind.js line numbers.
// In bind.js, bindUi starts at bindUiLine (0-based). Original bindUi was at line 1160 (1-based).
// Content mapping: original line L (1-based) → bind line (bindUiLine + (L - 1160))
function bindLineFromOriginal(orig1) {
    return bindUiLine + (orig1 - 1160);
}

const PANELS = [
    { file: 'tenant-settings-ui-bind-setup.js', exportName: 'installTenantBindSetup', oStart: 1183, oEnd: 1469 },
    { file: 'tenant-settings-ui-arges.js', exportName: 'installTenantArgesPanel', oStart: 1470, oEnd: 2026 },
    { file: 'tenant-settings-ui-status.js', exportName: 'installTenantStatusPanel', oStart: 2027, oEnd: 2413 },
    { file: 'tenant-settings-ui-directory.js', exportName: 'installTenantDirectoryHelpers', oStart: 2414, oEnd: 3200 },
    { file: 'tenant-settings-ui-subjects-panel.js', exportName: 'installTenantSubjectsPanel', oStart: 3201, oEnd: 3298 },
    { file: 'tenant-settings-ui-teachers.js', exportName: 'installTenantTeachersPanel', oStart: 3299, oEnd: 4057 },
    { file: 'tenant-settings-ui-admin.js', exportName: 'installTenantAdminPanel', oStart: 4058, oEnd: 4259 },
    { file: 'tenant-settings-ui-sga.js', exportName: 'installTenantSgaPanel', oStart: 4260, oEnd: 4689 },
    {
        file: 'tenant-settings-ui-student-council.js',
        exportName: 'installTenantStudentCouncilPanel',
        oStart: 4690,
        oEnd: 4980
    },
    {
        file: 'tenant-settings-ui-school-groups.js',
        exportName: 'installTenantSchoolGroupsPanel',
        oStart: 4981,
        oEnd: 5836
    },
    {
        file: 'tenant-settings-ui-teachers-lifecycle.js',
        exportName: 'installTenantTeachersLifecycle',
        oStart: 5837,
        oEnd: 6417
    },
    { file: 'tenant-settings-ui-import-flows.js', exportName: 'installTenantImportFlows', oStart: 6418, oEnd: 6789 },
    { file: 'tenant-settings-ui-students-panel.js', exportName: 'installTenantStudentsPanel', oStart: 6790, oEnd: 7367 },
    { file: 'tenant-settings-ui-classes-panel.js', exportName: 'installTenantClassesPanel', oStart: 7368, oEnd: 7802 },
    { file: 'tenant-settings-ui-render.js', exportName: 'installTenantRenderStorage', oStart: 7803, oEnd: 8566 },
    { file: 'tenant-settings-ui-bind-events.js', exportName: 'installTenantBindEvents', oStart: 8567, oEnd: 11001 }
];

function wrap(exportName, body, comment) {
    let b = convertDecls(body);
    b = rewriteToCtx(b);
    return `/**
 * ${comment}
 * @param {Record<string, any>} ctx
 */
export function ${exportName}(ctx) {
${b
    .split('\n')
    .map((l) => (l.length ? '    ' + l : l))
    .join('\n')}
}
`;
}

for (const panel of PANELS) {
    const a = bindLineFromOriginal(panel.oStart);
    const b = bindLineFromOriginal(panel.oEnd);
    const chunk = lines.slice(a, b + 1).join('\n');
    const text = wrap(panel.exportName, chunk, panel.file);
    fs.writeFileSync(path.join(ROOT, panel.file), text);
    const check = spawnSync(process.execPath, ['--check', path.join(ROOT, panel.file)], { encoding: 'utf8' });
    console.log(panel.file, text.split('\n').length, check.status === 0 ? 'OK' : 'FAIL');
    if (check.status !== 0) {
        console.error(check.stderr || check.stdout);
        process.exit(1);
    }
}

// Spot-check arges params
const arges = fs.readFileSync(path.join(ROOT, 'tenant-settings-ui-arges.js'), 'utf8');
if (arges.includes('function argesToLines(ctx.')) {
    console.error('param rewrite regression');
    process.exit(1);
}
console.log('all panels syntax OK');
