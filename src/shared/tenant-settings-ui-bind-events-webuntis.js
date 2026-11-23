/**
 * tenant-settings-ui-bind-events-webuntis.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindEventsWebuntis(ctx) {
    ctx.webuntisMergeDeps = function webuntisMergeDeps() {
        return {
            webuntis: window.ms365WebuntisExportImport,
            schoolSis: window.ms365SchoolSisImport,
            parseTeachersLines:
                typeof window.ms365TenantSettingsParseTeachersLines === 'function'
                    ? window.ms365TenantSettingsParseTeachersLines
                    : function () {
                          return [];
                      },
            parseStudentsLines:
                typeof window.ms365TenantSettingsParseStudentsLines === 'function'
                    ? window.ms365TenantSettingsParseStudentsLines
                    : function () {
                          return [];
                      },
            parseSubjectsLines:
                typeof window.ms365TenantSettingsParseSubjectsLines === 'function'
                    ? window.ms365TenantSettingsParseSubjectsLines
                    : function () {
                          return [];
                      },
            parseClassesLines:
                typeof window.ms365TenantSettingsParseClassesLines === 'function'
                    ? window.ms365TenantSettingsParseClassesLines
                    : function () {
                          return [];
                      }
        };
    }

    ctx.currentTenantImportLines = function currentTenantImportLines() {
        return {
            teachersLines: ctx.taTeachers ? ctx.taTeachers.value : '',
            studentsLines: ctx.taStudents ? ctx.taStudents.value : '',
            subjectsLines: ctx.taSubjects ? ctx.taSubjects.value : '',
            classesLines: ctx.taClasses ? ctx.taClasses.value : ''
        };
    }

    ctx.applyWebuntisMerged = async function applyWebuntisMerged(merged) {
        if (!merged) return;
        const lines = merged.lines || {};
        if (lines.subjectsLines != null && typeof ctx.parseLinesToSubjects === 'function') {
            const nextSubjects = ctx.parseLinesToSubjects(lines.subjectsLines);
            const ok = await ctx.confirmSubjectCatalogIfShrinks(nextSubjects, 'WebUntis-Import');
            if (!ok) {
                if (ctx.summary) {
                    ctx.summary.textContent =
                        'WebUntis-Import abgebrochen – die Fächerliste würde Kürzel verlieren. Lehrer/Schüler wurden nicht übernommen.';
                    ctx.summary.className = 'summary warn';
                }
                return;
            }
        }
        const prevStudents = ctx.getStudentsFromTextarea();
        if (ctx.taTeachers && lines.teachersLines != null) {
            ctx.taTeachers.value = lines.teachersLines;
            ctx.teachersSortState.key = null;
            ctx.teachersSortState.dir = 1;
            ctx.renderTeachersTableFromTextarea();
        }
        if (ctx.taStudents && lines.studentsLines != null) {
            ctx.taStudents.value = lines.studentsLines;
            ctx.studentsSortState.key = null;
            ctx.studentsSortState.dir = 1;
            ctx.renderStudentsTableFromTextarea();
        }
        if (ctx.taSubjects && lines.subjectsLines != null) {
            ctx.taSubjects.value = lines.subjectsLines;
            ctx.renderSubjectsTableFromTextarea();
        }
        if (ctx.taClasses && lines.classesLines != null) {
            ctx.taClasses.value = lines.classesLines;
            ctx.renderClassesTableFromTextarea();
        }
        const sis = ctx.sisApi();
        if (merged.studentDiff && sis) {
            ctx.rememberSisHistory(merged.studentDiff, 'webuntis', 'merge');
            ctx.renderStudentLifecyclePanel(prevStudents, ctx.getStudentsFromTextarea());
            ctx.renderStudentSammelLifecyclePanel(prevStudents, ctx.getStudentsFromTextarea());
        }
        if (ctx.summary) {
            const detail = ctx.summarizeWebuntisMergeResult(merged.stats, merged.studentDiff, sis || {});
            ctx.summary.textContent =
                'WebUntis-Import mit bestehenden Stammdaten abgeglichen (kein Doppelimport). ' +
                detail +
                ' – bitte unten „Speichern“ klicken (aktualisiert auch die Schul-Landkarte).';
            ctx.summary.className =
                merged.studentDiff && merged.studentDiff.conflicts && merged.studentDiff.conflicts.length
                    ? 'summary warn'
                    : 'summary ok';
        }
        if (merged.studentDiff && merged.studentDiff.conflicts && merged.studentDiff.conflicts.length) {
            const msgs = merged.studentDiff.conflicts
                .slice(0, 5)
                .map(function (c) {
                    return c.summary || c.type;
                })
                .join('\n');
            ctx.dlgAlert(
                'Beim Abgleich wurden Konflikte gefunden (bitte in der Schülerliste prüfen):\n\n' + msgs,
                { title: 'WebUntis-Abgleich', okText: 'Verstanden' }
            );
        }
        ctx.scheduleAutoSave();
    }

    ctx.mountTenantImportPanel({
        getCurrentLines: ctx.currentTenantImportLines,
        getMergeDeps: ctx.webuntisMergeDeps,
        sisApi: ctx.sisApi,
        summary: ctx.summary,
        onApplyMerged: function (merged) {
            return ctx.applyWebuntisMerged(merged);
        }
    });

    window.ms365TenantKlassenHubApi = {
        getClasses: function () {
            return ctx.getClassesFromTextarea();
        },
        getStudents: function () {
            return ctx.getStudentsFromTextarea();
        },
        normClassCode: ctx.normCode,
        getClassM365: function (code) {
            return ctx.getClassGroupMatchByKey(code);
        },
        updateStudent: function (idx, patch) {
            const all = ctx.getStudentsFromTextarea();
            if (!all[idx] || !patch) return;
            Object.keys(patch).forEach(function (k) {
                all[idx][k] = patch[k];
            });
            ctx.setStudentsTextareaFromRows(all);
            ctx.renderStudentsTableFromTextarea();
            ctx.scheduleAutoSave();
            try {
                window.dispatchEvent(new CustomEvent('ms365-tenant-klassen-hub-refresh'));
            } catch {
                /* ignore */
            }
        },
        removeStudent: function (idx) {
            const all = ctx.getStudentsFromTextarea();
            if (!all[idx]) return;
            all.splice(idx, 1);
            ctx.setStudentsTextareaFromRows(all);
            ctx.renderStudentsTableFromTextarea();
            ctx.scheduleAutoSave();
            try {
                window.dispatchEvent(new CustomEvent('ms365-tenant-klassen-hub-refresh'));
            } catch {
                /* ignore */
            }
        },
        addStudent: function (row) {
            const all = ctx.getStudentsFromTextarea();
            all.push({
                klasse: ctx.normStr(row && row.klasse),
                name: ctx.normStr(row && row.name),
                email: ctx.normStr(row && row.email).toLowerCase(),
                parentPairs: []
            });
            ctx.setStudentsTextareaFromRows(all);
            ctx.renderStudentsTableFromTextarea();
            ctx.scheduleAutoSave();
            try {
                window.dispatchEvent(new CustomEvent('ms365-tenant-klassen-hub-refresh'));
            } catch {
                /* ignore */
            }
        },
        focusClassInListView: function (classCode) {
            const code = ctx.normCode(classCode);
            if (ctx.studentsClassFilter) {
                ctx.studentsClassFilter.value = code;
                ctx.studentsClassFilter.dispatchEvent(new Event('change', { bubbles: true }));
            }
            const tab = document.getElementById('tabMainSchueler');
            if (tab) tab.click();
        },
        addClass: function (code, name) {
            const c = ctx.normCode(code);
            if (!c) return false;
            const all = ctx.getClassesFromTextarea();
            if (all.some(function (row) {
                return ctx.normCode(row && row.code) === c;
            })) {
                return false;
            }
            all.push({
                code: c,
                name: ctx.normStr(name) || c,
                year: '',
                headName: '',
                headEmail: ''
            });
            ctx.setClassesTextareaFromRows(all);
            ctx.renderClassesTableFromTextarea();
            ctx.scheduleAutoSave();
            try {
                window.dispatchEvent(new CustomEvent('ms365-tenant-klassen-hub-refresh'));
            } catch {
                /* ignore */
            }
            return true;
        }
    };
    try {
        window.dispatchEvent(new CustomEvent('ms365-tenant-klassen-hub-ready'));
    } catch {
        /* ignore */
    }

    // Accordion: immer nur EIN Schritt offen (details.step)
    try {
        const steps = Array.from(document.querySelectorAll('details.step'));
        steps.forEach((d) => {
            d.addEventListener('toggle', () => {
                if (!d.open) return;
                steps.forEach((o) => {
                    if (o !== d) o.open = false;
                });
            });
        });
    } catch {
        // ignore
    }
}
