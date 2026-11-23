/**
 * tenant-settings-ui-bind-events-catalog.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindEventsCatalog(ctx) {
    if (ctx.taSubjects) {
        ctx.taSubjects.addEventListener('input', () => {
            ctx.taSubjects.dataset.ms365UserEdited = '1';
            ctx.subjectsSortState.key = null;
            ctx.subjectsSortState.dir = 1;
            ctx.renderSubjectsTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }

    if (ctx.taArges) {
        ctx.taArges.addEventListener('input', () => {
            ctx.taArges.dataset.ms365UserEdited = '1';
            ctx.argesSortState.key = null;
            ctx.argesSortState.dir = 1;
            ctx.renderArgesTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }

    if (ctx.btnAddSubjectRow) {
        ctx.btnAddSubjectRow.addEventListener('click', () => {
            const all = ctx.getSubjectsFromTextarea();
            all.push({ code: '', name: '' });
            ctx.setSubjectsTextareaFromRows(all);
            ctx.renderSubjectsTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.btnEnrichSubjects && !ctx.btnEnrichSubjects.dataset.tenantEnrichSubjectsBound) {
        ctx.btnEnrichSubjects.dataset.tenantEnrichSubjectsBound = '1';
        ctx.btnEnrichSubjects.addEventListener('click', () => {
            const rows = ctx.getSubjectsFromTextarea();
            const result = ctx.applySubjectEnrichFromSources(rows, {
                persist: true,
                summary: true,
                reason: 'subject-enrich-manual'
            });
            if (!result.changed) {
                ctx.setSummary(
                    'Keine zusätzlichen Fächer in ARGE, Schulstruktur oder Legacy-Speicher gefunden – Liste unverändert.',
                    'warn'
                );
            }
            ctx.renderStatusOverview();
        });
    }

    if (ctx.btnAddArgeRow) {
        ctx.btnAddArgeRow.addEventListener('click', () => {
            const all = ctx.getArgesFromTextarea();
            all.push({ code: '', name: '', subjects: [], headName: '', headEmail: '' });
            ctx.setArgesTextareaFromRows(all);
            ctx.argesTableEditFocus = { rowIdx: all.length - 1, field: 'code' };
            ctx.renderArgesTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.btnVerifyArgesGraph && !ctx.btnVerifyArgesGraph.dataset.tenantVerifyArgesBound) {
        ctx.btnVerifyArgesGraph.dataset.tenantVerifyArgesBound = '1';
        ctx.btnVerifyArgesGraph.addEventListener('click', async () => {
            const rows = ctx.getArgesFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Keine ARGEs zum Prüfen vorhanden.', 'warn');
                return;
            }
            const btn = ctx.btnVerifyArgesGraph;
            let found = 0;
            let missed = 0;
            let skipped = 0;
            const seen = new Set();
            try {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
                ctx.setSummary('ARGE-Gruppen-Abgleich läuft …', 'warn');
                let token = null;
                for (let i = 0; i < rows.length; i++) {
                    const key = ctx.normCode(rows[i].code);
                    if (!key) {
                        skipped++;
                        continue;
                    }
                    if (seen.has(key)) continue;
                    seen.add(key);
                    if (!token) {
                        try {
                            token = await ctx.graphApi().getGraphToken();
                        } catch {
                            token = null;
                        }
                    }
                    const res = await ctx.verifyArgeGroupForRow(rows[i], token ? { token: token } : {});
                    if (res && res.skipped) skipped++;
                    else if (res && res.found) found++;
                    else missed++;
                    if (typeof ctx.graphApi().sleep === 'function' && i < rows.length - 1) {
                        await ctx.graphApi().sleep(300);
                    }
                }
                ctx.renderArgesTableFromTextarea();
                ctx.renderStatusOverview();
                ctx.setSummary(
                    'ARGEs: ' + found + ' gefunden, ' + missed + ' nicht gefunden, ' + skipped + ' ohne Kürzel',
                    missed ? 'warn' : 'ok'
                );
            } catch (e) {
                ctx.setSummary('ARGE-Gruppen-Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        });
    }

    if (ctx.taTeachers) {
        const refreshTeachersTableFromTa = () => {
            ctx.teachersSortState.key = null;
            ctx.teachersSortState.dir = 1;
            ctx.renderTeachersTableFromTextarea();
        };
        ctx.taTeachers.addEventListener('input', refreshTeachersTableFromTa);
        ctx.taTeachers.addEventListener('input', () => {
            ctx.taTeachers.dataset.ms365UserEdited = '1';
            ctx.scheduleAutoSave();
        });
        // Paste löst in manchen Umgebungen verzögert `input` aus – Tabelle trotzdem sofort aktualisieren.
        ctx.taTeachers.addEventListener('paste', () => {
            requestAnimationFrame(refreshTeachersTableFromTa);
        });
    }
    if (ctx.teachersTable && !ctx.teachersTable.dataset.teachersSortBound) {
        ctx.teachersTable.dataset.teachersSortBound = '1';
        ctx.teachersTable.querySelectorAll('th.is-sortable').forEach((th) => {
            th.addEventListener('click', () => ctx.applyTeachersSort(th.dataset.sortKey));
            th.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    ctx.applyTeachersSort(th.dataset.sortKey);
                }
            });
        });
    }
    ctx.bindTenantTableSort(ctx.subjectsTable, ctx.subjectsSortState, ['code', 'name'], ctx.applySubjectsSort);
    ctx.bindTenantTableSort(ctx.argesTable, ctx.argesSortState, ['code', 'name', 'subjects', 'headName', 'headEmail'], ctx.applyArgesSort);
    ctx.bindTenantTableSort(ctx.classesTable, ctx.classesSortState, ['code', 'year', 'name', 'headName', 'headEmail'], ctx.applyClassesSort);
    ctx.bindTenantTableSort(ctx.sgaTable, ctx.sgaSortState, ['scope', 'name', 'email'], ctx.applySgaSort);
    ctx.bindTenantTableSort(
        ctx.studentCouncilTable,
        ctx.studentCouncilSortState,
        ['klasse', 'name', 'email'],
        ctx.applyStudentCouncilSort
    );
    if (ctx.btnAddTeacherRow) {
        ctx.btnAddTeacherRow.addEventListener('click', () => {
            const all = ctx.getTeachersFromTextarea();
            all.push({ code: '', name: '', email: '' });
            ctx.setTeachersTextareaFromRows(all);
            ctx.renderTeachersTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.btnTeachersExportCsv) {
        ctx.btnTeachersExportCsv.addEventListener('click', () => {
            const api = window.ms365TeacherListImport;
            const rows = ctx.getTeachersFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Lehrerliste ist leer – nichts zu exportieren.', 'warn');
                return;
            }
            if (!api || typeof api.exportCsv !== 'function' || !api.exportCsv(rows)) {
                ctx.setSummary('CSV-Export fehlgeschlagen.', 'warn');
                return;
            }
            ctx.setSummary('Lehrerliste als CSV exportiert (' + rows.length + ').', 'ok');
        });
    }
    if (ctx.btnTeachersExportXlsx) {
        ctx.btnTeachersExportXlsx.addEventListener('click', () => {
            const api = window.ms365TeacherListImport;
            const rows = ctx.getTeachersFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Lehrerliste ist leer – nichts zu exportieren.', 'warn');
                return;
            }
            if (!api || (typeof api.isXlsxReady === 'function' && !api.isXlsxReady())) {
                ctx.setSummary('Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
                return;
            }
            if (!api || typeof api.exportXlsx !== 'function' || !api.exportXlsx(rows)) {
                ctx.setSummary('XLSX-Export fehlgeschlagen.', 'warn');
                return;
            }
            ctx.setSummary('Lehrerliste als XLSX exportiert (' + rows.length + ').', 'ok');
        });
    }
    if (ctx.btnVerifyTeachersGraph && !ctx.btnVerifyTeachersGraph.dataset.tenantVerifyTeachersBound) {
        ctx.btnVerifyTeachersGraph.dataset.tenantVerifyTeachersBound = '1';
        ctx.btnVerifyTeachersGraph.addEventListener('click', async () => {
            await ctx.runVerifyGraphDirectoryRows(
                ctx.getTeachersFromTextarea(),
                function (r) {
                    return r.email;
                },
                'Lehrkräfte',
                ctx.btnVerifyTeachersGraph,
                function () {
                    ctx.renderTeachersTableFromTextarea();
                }
            );
        });
    }
}
