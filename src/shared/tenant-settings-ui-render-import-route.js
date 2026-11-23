/**
 * tenant-settings-ui-render-import-route.js
 * @param {Record<string, any>} ctx
 */
export function installTenantRenderImportRoute(ctx) {
    ctx.importSubjectsFromWebuntisPdfFile = async function importSubjectsFromWebuntisPdfFile(file) {
        const wu = window.ms365WebuntisExportImport;
        const pdfApi = window.ms365PdfText;
        if (!wu || typeof wu.importSubjectsFromWebuntisPdf !== 'function') {
            ctx.setSummary('Fächer-PDF: WebUntis-Modul fehlt – Seite neu laden.', 'warn');
            return;
        }
        if (!pdfApi || typeof pdfApi.extractPdfFile !== 'function') {
            ctx.setSummary('Fächer-PDF: PDF-Modul fehlt – Seite neu laden.', 'warn');
            return;
        }
        ctx.setSummary('Fächer-PDF wird gelesen …', 'ok');
        try {
            const content = await pdfApi.extractPdfFile(file);
            const result = wu.importSubjectsFromWebuntisPdf({
                text: content.text,
                words: content.words
            });
            const subjects = Array.isArray(result.subjects) ? result.subjects : [];
            if (!subjects.length) {
                ctx.setSummary('Keine Fächer im PDF erkannt (WebUntis Fächer-Export prüfen).', 'warn');
                return;
            }
            const picked = await ctx.openSubjectPdfImportReview({
                subjects: subjects,
                sourceFileName: file && file.name,
                meta: result.meta
            });
            if (!picked || !picked.subjects || !picked.subjects.length) {
                ctx.setSummary('Fächer-Import abgebrochen – keine Auswahl übernommen.', 'warn');
                return;
            }
            const out = picked.subjects
                .map(function (s) {
                    return { code: ctx.normCode(s && s.code), name: ctx.normStr(s && s.name || '') };
                })
                .filter(function (s) {
                    return !!s.code;
                });
            const ok = await ctx.confirmSubjectCatalogIfShrinks(out, 'Fächer-Import (WebUntis-PDF)');
            if (!ok) {
                ctx.setSummary('Fächer-Import abgebrochen – bestehende Liste unverändert.', 'warn');
                return;
            }
            if (ctx.taSubjects) {
                ctx.taSubjects.value = out.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
            }
            ctx.renderSubjectsTableFromTextarea();
            const meta = result.meta || {};
            const bits = [`${out.length} Fachgruppen übernommen`];
            if (subjects.length > out.length) {
                bits.push(subjects.length - out.length + ' im PDF nicht übernommen');
            }
            if (meta.adminCount) bits.push(`${meta.adminCount} Verwaltungsfächer im PDF`);
            ctx.setSummary('Fächer importiert: ' + bits.join(' · '), 'ok');
            ctx.scheduleAutoSave();
        } catch (e) {
            ctx.setSummary('Fächer-PDF: ' + (e && e.message ? e.message : String(e)), 'warn');
        }
    }

    ctx.importSubjectsRows = async function importSubjectsRows(jsonRows) {
        const out = ctx.subjectsFromSpreadsheetJsonRows(jsonRows);
        const ok = await ctx.confirmSubjectCatalogIfShrinks(out, 'Fächer-Import (CSV/XLSX)');
        if (!ok) {
            ctx.setSummary('Fächer-Import abgebrochen – bestehende Liste unverändert.', 'warn');
            return;
        }
        if (ctx.taSubjects) ctx.taSubjects.value = out.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
        ctx.renderSubjectsTableFromTextarea();
        ctx.setSummary(`Fächer importiert: ${out.length} (schulweit)`, 'ok');
        ctx.scheduleAutoSave();
    }

    ctx.focusRegisterTabForImportTarget = function focusRegisterTabForImportTarget(target) {
        const btnId = ctx.REGISTER_IMPORT_TAB_BTN[target];
        if (!btnId) return;
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-register-tab-request', { detail: { tabBtnId: btnId } })
            );
        } catch {
            /* ignore */
        }
    }

    ctx.importRoutedSuffix = function importRoutedSuffix(originTab, target) {
        if (!target || target === originTab || target === 'generic') return '';
        const label = ctx.REGISTER_IMPORT_TARGET_LABEL[target];
        return label ? ' → ' + label : '';
    }

    ctx.importSubjectsFromWebuntisSpreadsheet = async function importSubjectsFromWebuntisSpreadsheet(aoa, sourceLabel, routeSuffix) {
        const wu = window.ms365WebuntisExportImport;
        if (!wu || typeof wu.detectExportKindFromAoa !== 'function') return false;
        const kind = wu.detectExportKindFromAoa(aoa);
        let subjects = [];
        if (kind === 'subject' && typeof wu.parseSubjectsAoa === 'function') {
            subjects = wu.parseSubjectsAoa(aoa);
        } else if (kind === 'lessons' && typeof wu.uniqueSubjectsFromLessonsAoa === 'function') {
            subjects = wu.uniqueSubjectsFromLessonsAoa(aoa);
        } else {
            return false;
        }
        const out = (subjects || [])
            .map(function (s) {
                return { code: ctx.normCode(s && s.code), name: ctx.normStr(s && s.name || '') };
            })
            .filter(function (s) {
                return !!s.code;
            });
        if (!out.length) return false;
        const ok = await ctx.confirmSubjectCatalogIfShrinks(out, sourceLabel || 'Fächer-Import (WebUntis)');
        if (!ok) {
            ctx.setSummary('Fächer-Import abgebrochen – bestehende Liste unverändert.', 'warn');
            return true;
        }
        if (ctx.taSubjects) {
            ctx.taSubjects.value = out.map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
        }
        ctx.renderSubjectsTableFromTextarea();
        ctx.setSummary('Fächer importiert: ' + out.length + ' (WebUntis)' + (routeSuffix || ''), 'ok');
        ctx.scheduleAutoSave();
        return true;
    }

    ctx.importTeachersFromWebuntisSpreadsheet = function importTeachersFromWebuntisSpreadsheet(aoa, routeSuffix) {
        const wu = window.ms365WebuntisExportImport;
        if (!wu || typeof wu.importTeachersFromWebuntis !== 'function') return false;
        if (wu.detectExportKindFromAoa(aoa) !== 'teacher') return false;
        const domain = ctx.schoolDomainForEmail();
        const pattern = ctx.tenantSchoolEmailPatternId();
        const firstNameMode = ctx.tenantSchoolEmailFirstNameMode();
        const result = wu.importTeachersFromWebuntis({
            teacherAoa: aoa,
            domain: domain,
            pattern: pattern,
            firstNameMode: firstNameMode,
            applyEmails: !!domain
        });
        ctx.showTeachersWuDiffPreview(ctx.getTeachersFromTextarea(), result, routeSuffix);
        return true;
    }

    ctx.importStudentsViaModule = function importStudentsViaModule(file, sourceHint) {
        if (!window.ms365StudentListImport || typeof window.ms365StudentListImport.importFile !== 'function') {
            ctx.importFileToRows(file, (rows) => ctx.importStudentsRows(rows));
            return;
        }
        const sourceEl = document.getElementById('tenantStudentsImportSource');
        const hint = sourceHint || (sourceEl ? String(sourceEl.value || 'auto') : 'auto');
        window.ms365StudentListImport.importFile(
            file,
            (lines, result) => {
                const prev = ctx.getStudentsFromTextarea();
                ctx.showSisDiffPreview(prev, result || { lines: lines, records: [], source: hint });
                const meta = result && result.meta ? result.meta : {};
                const bits = [];
                if (meta.emailsGenerated) bits.push(meta.emailsGenerated + ' Mails vorgeschlagen');
                if (meta.unmatchedGuardians) bits.push(meta.unmatchedGuardians + ' Eltern ohne Zuordnung');
                if (meta.emailHint) bits.push(meta.emailHint);
                if (meta.classHint) bits.push(meta.classHint);
                if (bits.length) ctx.setSummary('WebUntis/SIS: ' + bits.join(' · '), 'ok');
            },
            (msg) => ctx.setSummary(msg, 'warn'),
            hint,
            ctx.getStudentsFromTextarea()
        );
    }

    /**
     * WebUntis- oder Excel/CSV-Import: erkennt Typ und schreibt in die richtige Register-Liste.
     * @param {File} file
     * @param {'subjects'|'teachers'|'students'|'classes'|'arges'} originTab
     */
    ctx.handleRegisterImportFile = async function handleRegisterImportFile(file, originTab) {
        if (!file) return;
        const wu = window.ms365WebuntisExportImport;
        const low = String(file.name || '').toLowerCase();
        const isPdf = low.endsWith('.pdf') || String(file.type || '').toLowerCase() === 'application/pdf';

        if (isPdf) {
            const pdfHint = ctx.pdfImportTargetFromFilename(file.name);
            const pdfTarget =
                pdfHint ||
                (originTab === 'classes' ? 'classes' : originTab === 'subjects' ? 'subjects' : 'subjects');
            if (pdfTarget === 'classes') {
                if (originTab !== 'classes') ctx.focusRegisterTabForImportTarget('classes');
                ctx.importClassesFromWebuntisPdfFile(file);
                return;
            }
            if (originTab !== 'subjects') ctx.focusRegisterTabForImportTarget('subjects');
            await ctx.importSubjectsFromWebuntisPdfFile(file);
            return;
        }

        let aoa = null;
        try {
            const sheets = await ctx.readFilesToAoaSheets([file]);
            aoa = sheets[0] && sheets[0].aoa;
        } catch (e) {
            ctx.setSummary('Import: ' + (e && e.message ? e.message : String(e)), 'warn');
            return;
        }

        const detected = ctx.spreadsheetImportTargetFromAoa(aoa, wu && wu.detectExportKindFromAoa);
        const target = ctx.resolveRegisterImportTarget(originTab, detected);
        if (target !== originTab && target !== 'generic') {
            ctx.focusRegisterTabForImportTarget(target);
        }
        const routeNote = ctx.importRoutedSuffix(originTab, target);

        if (target === 'teachers') {
            if (ctx.importTeachersFromWebuntisSpreadsheet(aoa, routeNote)) return;
            ctx.importFileToRows(file, (rows) => ctx.importTeachersRows(rows));
            return;
        }
        if (target === 'students') {
            ctx.importStudentsViaModule(file, 'webuntis');
            return;
        }
        if (target === 'subjects') {
            const usedWu = await ctx.importSubjectsFromWebuntisSpreadsheet(
                aoa,
                'Fächer-Import (WebUntis CSV)',
                routeNote
            );
            if (usedWu) return;
            ctx.importFileToRows(file, (rows) => void ctx.importSubjectsRows(rows));
            return;
        }
        if (target === 'classes') {
            ctx.importFileToRows(file, (rows) => ctx.importClassesRows(rows));
            return;
        }
        if (target === 'arges') {
            ctx.importFileToRows(file, (rows) => ctx.importArgesRows(rows));
            return;
        }

        if (originTab === 'teachers') {
            ctx.importFileToRows(file, (rows) => ctx.importTeachersRows(rows));
        } else if (originTab === 'students') {
            ctx.importStudentsViaModule(file, 'auto');
        } else if (originTab === 'classes') {
            ctx.importFileToRows(file, (rows) => ctx.importClassesRows(rows));
        } else if (originTab === 'arges') {
            ctx.importFileToRows(file, (rows) => ctx.importArgesRows(rows));
        } else {
            ctx.importFileToRows(file, (rows) => void ctx.importSubjectsRows(rows));
        }
    }

}
