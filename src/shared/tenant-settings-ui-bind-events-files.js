/**
 * tenant-settings-ui-bind-events-files.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindEventsFiles(ctx) {
    if (ctx.fileSubjects) {
        ctx.fileSubjects.addEventListener('change', (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            void ctx.handleRegisterImportFile(f, 'subjects').finally(function () {
                ctx.closeListImportDrop(ctx.fileSubjects);
            });
            ctx.fileSubjects.value = '';
        });
    }
    if (ctx.fileArges) {
        ctx.fileArges.addEventListener('change', (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            void ctx.handleRegisterImportFile(f, 'arges').finally(function () {
                ctx.closeListImportDrop(ctx.fileArges);
            });
            ctx.fileArges.value = '';
        });
    }
    if (ctx.btnArgesImportM365) {
        ctx.btnArgesImportM365.addEventListener('click', () => {
            ctx.closeListImportDrop(ctx.fileArges);
            ctx.importArgesFromM365Groups();
        });
    }
    if (ctx.btnSubjectsImportM365) {
        ctx.btnSubjectsImportM365.addEventListener('click', () => {
            ctx.closeListImportDrop(ctx.fileSubjects);
            ctx.importSubjectsFromM365Groups();
        });
    }
    if (ctx.btnClassesImportM365) {
        ctx.btnClassesImportM365.addEventListener('click', () => {
            ctx.closeListImportDrop(ctx.fileClasses);
            ctx.importClassesFromM365Groups();
        });
    }
    if (ctx.fileTeachers) {
        ctx.fileTeachers.addEventListener('change', (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            void ctx.handleRegisterImportFile(f, 'teachers').finally(function () {
                ctx.closeListImportDrop(ctx.fileTeachers);
            });
            ctx.fileTeachers.value = '';
        });
    }
    if (ctx.fileStudents) {
        ctx.fileStudents.addEventListener('change', (e) => {
            const files = e.target.files;
            const list = files && files.length != null ? Array.from(files) : [];
            const done = function () {
                ctx.closeListImportDrop(ctx.fileStudents);
            };
            if (list.length === 1) {
                void ctx.handleRegisterImportFile(list[0], 'students').finally(done);
            } else if (list.length > 1) {
                const sourceEl = document.getElementById('tenantStudentsImportSource');
                const sourceHint = sourceEl ? String(sourceEl.value || 'auto') : 'auto';
                ctx.importStudentsViaModule(files, sourceHint);
                done();
            } else {
                done();
            }
            ctx.fileStudents.value = '';
        });
    }
    ctx.btnSisMerge = document.getElementById('tenantSisDiffMerge');
    ctx.btnSisReplace = document.getElementById('tenantSisDiffReplace');
    ctx.btnSisCancel = document.getElementById('tenantSisDiffCancel');
    if (ctx.btnSisMerge) ctx.btnSisMerge.addEventListener('click', () => ctx.applyPendingSis('merge'));
    if (ctx.btnSisReplace) ctx.btnSisReplace.addEventListener('click', () => ctx.applyPendingSis('replace'));
    if (ctx.btnSisCancel) ctx.btnSisCancel.addEventListener('click', ctx.hideSisDiff);
    ctx.btnTeachersWuMerge = document.getElementById('tenantTeachersWuDiffMerge');
    ctx.btnTeachersWuReplace = document.getElementById('tenantTeachersWuDiffReplace');
    ctx.btnTeachersWuCancel = document.getElementById('tenantTeachersWuDiffCancel');
    if (ctx.btnTeachersWuMerge) ctx.btnTeachersWuMerge.addEventListener('click', () => ctx.applyPendingTeachersWu('merge'));
    if (ctx.btnTeachersWuReplace) ctx.btnTeachersWuReplace.addEventListener('click', () => ctx.applyPendingTeachersWu('replace'));
    if (ctx.btnTeachersWuCancel) {
        ctx.btnTeachersWuCancel.addEventListener('click', () => ctx.hideTeachersWuDiff({ restore: true }));
    }

    ctx.renderActionLog = function renderActionLog() {
        const ul = document.getElementById('tenantActionLogList');
        if (!ul) return;
        ul.innerHTML = '';
        const rows =
            window.ms365ActionLog && typeof window.ms365ActionLog.list === 'function'
                ? window.ms365ActionLog.list(40)
                : [];
        if (!rows.length) {
            const li = document.createElement('li');
            li.textContent = 'Noch keine Einträge.';
            ul.appendChild(li);
            return;
        }
        rows.forEach(function (row) {
            const li = document.createElement('li');
            const when = row.at ? String(row.at).replace('T', ' ').slice(0, 19) : '';
            li.textContent =
                (when ? when + ' · ' : '') +
                (row.tool || 'app') +
                ': ' +
                (row.summary || row.action) +
                (row.result === 'error' ? ' (Fehler)' : '');
            ul.appendChild(li);
        });
    }
    ctx.btnLogRefresh = document.getElementById('tenantActionLogRefresh');
    ctx.btnLogExport = document.getElementById('tenantActionLogExport');
    ctx.btnLogClear = document.getElementById('tenantActionLogClear');
    if (ctx.btnLogRefresh) ctx.btnLogRefresh.addEventListener('click', ctx.renderActionLog);
    if (ctx.btnLogExport) {
        ctx.btnLogExport.addEventListener('click', function () {
            if (!window.ms365ActionLog || typeof window.ms365ActionLog.exportJson !== 'function') return;
            const blob = new Blob([window.ms365ActionLog.exportJson()], { type: 'application/json' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'ms365-aktionsprotokoll.json';
            document.body.appendChild(a);
            a.click();
            a.remove();
            setTimeout(function () {
                URL.revokeObjectURL(url);
            }, 250);
        });
    }
    if (ctx.btnLogClear) {
        ctx.btnLogClear.addEventListener('click', function () {
            if (window.ms365ActionLog) window.ms365ActionLog.clear();
            ctx.renderActionLog();
        });
    }
    ctx.renderActionLog();

    if (ctx.fileClasses) {
        ctx.fileClasses.addEventListener('change', (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            void ctx.handleRegisterImportFile(f, 'classes').finally(function () {
                ctx.closeListImportDrop(ctx.fileClasses);
            });
            ctx.fileClasses.value = '';
        });
    }

    if (ctx.btnSubjectsTpl) {
        ctx.btnSubjectsTpl.addEventListener('click', () => {
            const ok =
                window.ms365SubjectListImport && typeof window.ms365SubjectListImport.downloadTemplate === 'function'
                    ? window.ms365SubjectListImport.downloadTemplate()
                    : false;
            if (!ok) ctx.setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
        });
    }
    if (ctx.btnArgesTpl) {
        ctx.btnArgesTpl.addEventListener('click', () => {
            const ok = ctx.downloadXlsxTemplate(
                'ARGE-Liste-Vorlage.xlsx',
                [
                    ['Kürzel', 'Name', 'Fächer'],
                    ['SPRACHEN', 'Sprachen', 'D,E,FS2'],
                    ['NAWI', 'Naturwissenschaften', 'BIO,CH,PH']
                ],
                'ARGEs'
            );
            if (!ok) ctx.setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
        });
    }
    if (ctx.btnTeachersTpl) {
        ctx.btnTeachersTpl.addEventListener('click', () => {
            const ok =
                window.ms365TeacherListImport && typeof window.ms365TeacherListImport.downloadTemplate === 'function'
                    ? window.ms365TeacherListImport.downloadTemplate()
                    : false;
            if (!ok) ctx.setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
        });
    }
    if (ctx.btnStudentsTpl) {
        ctx.btnStudentsTpl.addEventListener('click', () => {
            const ok =
                window.ms365StudentListImport && typeof window.ms365StudentListImport.downloadTemplate === 'function'
                    ? window.ms365StudentListImport.downloadTemplate()
                    : false;
            if (!ok) ctx.setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
        });
    }
    if (ctx.btnStudentsFillDemoParents) {
        ctx.btnStudentsFillDemoParents.addEventListener('click', (ev) => {
            if (!ctx.taStudents) return;
            const before = String(ctx.taStudents.value || '');
            if (!before.trim()) {
                ctx.setSummary('Schülerliste ist leer – zuerst Schüler:innen eintragen oder aus M365 einlesen.', 'warn');
                return;
            }
            const force = !!(ev && (ev.shiftKey || ev.altKey));
            const result = ctx.enrichStudentLinesWithDemoParents(before, { force: force });
            ctx.taStudents.value = result.text;
            if (typeof ctx.renderStudentsTableFromTextarea === 'function') {
                ctx.renderStudentsTableFromTextarea();
            }
            const changed = (result.enriched || 0) + (result.repaired || 0);
            const msg = force
                ? result.enriched +
                  ' Zeile(n) mit Demo-Eltern neu gesetzt (Shift/Alt = überschreiben). Bitte Speichern.'
                : changed
                  ? [
                        result.enriched
                            ? result.enriched + ' Zeile(n) um Demo-Eltern ergänzt'
                            : '',
                        result.repaired
                            ? result.repaired + ' fehlerhafte Namen repariert'
                            : '',
                        result.skipped
                            ? result.skipped + ' bereits mit Eltern übersprungen'
                            : ''
                    ]
                        .filter(Boolean)
                        .join(', ') + '. Bitte Speichern.'
                  : result.skipped
                    ? 'Alle ' +
                      result.skipped +
                      ' Zeilen haben bereits Eltern. Mit Shift+Klick überschreiben.'
                    : 'Keine Schülerzeilen gefunden.';
            ctx.setSummary(msg, changed ? 'ok' : 'warn');
        });
    }
    if (ctx.btnClassesTpl) {
        ctx.btnClassesTpl.addEventListener('click', () => {
            const ok =
                window.ms365ClassListImport && typeof window.ms365ClassListImport.downloadTemplate === 'function'
                    ? window.ms365ClassListImport.downloadTemplate()
                    : false;
            if (!ok) ctx.setSummary('Vorlage: Excel-Bibliothek nicht geladen – Seite neu laden.', 'warn');
        });
    }
    if (ctx.btnClassesExportCsv) {
        ctx.btnClassesExportCsv.addEventListener('click', () => {
            const api = window.ms365ClassListImport;
            const rows = ctx.getClassesFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Klassenliste ist leer – nichts zu exportieren.', 'warn');
                return;
            }
            if (!api || typeof api.exportCsv !== 'function' || !api.exportCsv(rows)) {
                ctx.setSummary('CSV-Export fehlgeschlagen.', 'warn');
                return;
            }
            ctx.setSummary('Klassenliste als CSV exportiert (' + rows.length + ').', 'ok');
        });
    }
    if (ctx.btnClassesExportXlsx) {
        ctx.btnClassesExportXlsx.addEventListener('click', () => {
            const api = window.ms365ClassListImport;
            const rows = ctx.getClassesFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Klassenliste ist leer – nichts zu exportieren.', 'warn');
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
            ctx.setSummary('Klassenliste als XLSX exportiert (' + rows.length + ').', 'ok');
        });
    }

    ctx.btnReloads.forEach((btn) => btn.addEventListener('click', () => ctx.renderFromStorage()));

    ctx.onSchoolYearChangedFromMenu = function onSchoolYearChangedFromMenu(ev) {
        const d = ev && ev.detail;
        if (!d || !d.year) return;
        ctx.patchMatchedGroupId('studentCouncil', null);
        ctx.renderFromStorage();
    }
    if (!window.__ms365TenantSchoolYearListener) {
        window.__ms365TenantSchoolYearListener = true;
        window.addEventListener('ms365-school-year-changed', ctx.onSchoolYearChangedFromMenu);
    }
    ctx.bindSchoolYearControls();
    window.addEventListener('ms365-menu-header-ready', () => ctx.bindSchoolYearControls());
    window.addEventListener('ms365-auth-state-changed', () => ctx.bindSchoolYearControls());

    if (ctx.btnExport) {
        ctx.btnExport.addEventListener('click', () => {
            try {
                if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.downloadBackup === 'function') {
                    window.ms365BrowserBackup.downloadBackup();
                    return;
                }
                if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.exportJson === 'function') {
                    ctx.downloadJson('ms365-schooltool-data-v2.json', window.ms365AppDataV2.exportJson());
                    return;
                }
            } catch {
                // ignore
            }
            ctx.downloadJson('schule-einstellungen.json', ctx.load());
        });
    }
    if (ctx.btnExportHeader) {
        ctx.btnExportHeader.addEventListener('click', () => {
            try {
                if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.downloadBackup === 'function') {
                    window.ms365BrowserBackup.downloadBackup();
                    return;
                }
                if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.exportJson === 'function') {
                    ctx.downloadJson('ms365-schooltool-data-v2.json', window.ms365AppDataV2.exportJson());
                    return;
                }
            } catch {
                // ignore
            }
            ctx.downloadJson('schule-einstellungen.json', ctx.load());
        });
    }

    if (ctx.btnClear) {
        ctx.btnClear.addEventListener('click', () => {
            try {
                localStorage.removeItem('ms365-tenant-settings-v1');
            } catch {
                // ignore
            }
            // UI/Domain wieder auf Standard zurücksetzen
            try {
                if (ctx.schoolNameInput) ctx.schoolNameInput.value = '';
                ctx.domainInput = document.getElementById('schoolEmailDomain');
                if (ctx.domainInput) ctx.domainInput.value = 'ms365.schule';
                if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                    window.ms365SetSchoolDomainNoAt('ms365.schule');
                }
            } catch {
                // ignore
            }
            if (ctx.taSubjects) ctx.taSubjects.value = '';
            if (ctx.taArges) ctx.taArges.value = '';
            if (ctx.taTeachers) ctx.taTeachers.value = '';
            if (ctx.taAdminBundle) ctx.taAdminBundle.value = '';
            ctx.administrationGroupsSnapshot = [];
            if (ctx.taAdmin) ctx.taAdmin.value = '';
            if (ctx.taAdminRoles) ctx.taAdminRoles.value = '';
            if (ctx.selSgaMode) ctx.selSgaMode.value = 'group';
            if (ctx.taSga) ctx.taSga.value = '';
            if (ctx.taStudents) ctx.taStudents.value = '';
            if (ctx.taStudentCouncil) ctx.taStudentCouncil.value = '';
            if (ctx.taClasses) ctx.taClasses.value = '';
            ctx.renderSubjectsTableFromTextarea();
            ctx.renderArgesTableFromTextarea();
            ctx.renderTeachersTableFromTextarea();
            ctx.renderAdminRolesTableFromTextarea();
            ctx.renderAdminTableFromTextarea();
            ctx.renderSgaTableFromTextarea();
            ctx.renderStudentsTableFromTextarea();
            ctx.renderStudentCouncilTableFromTextarea();
            ctx.renderClassesTableFromTextarea();
            ctx.setSummary('Stammdaten gelöscht (nur lokaler Browser-Speicher).', 'warn');
        });
    }

    if (ctx.fileImport) {
        ctx.fileImport.addEventListener('change', async (e) => {
            const f = e.target.files && e.target.files[0];
            if (!f) return;
            try {
                const text = await f.text();
                const obj = ctx.safeJsonParse(text);
                if (!obj) {
                    ctx.setSummary('Import fehlgeschlagen: keine gültige JSON-Datei.', 'warn');
                    return;
                }
                if (window.ms365BrowserBackup && typeof window.ms365BrowserBackup.isBackupPayload === 'function' && window.ms365BrowserBackup.isBackupPayload(obj)) {
                    let backupMsg =
                        'Dieses Browser-Backup ersetzt alle lokalen Schuldaten und wiederherstellbaren Werkzeug-Zwischenstände in diesem Browser. Microsoft-Anmeldung und PIN-Freischaltungen bleiben unberührt. Fortfahren?';
                    const backupSubjects = ctx.subjectsFromImportPayload(obj);
                    if (backupSubjects.length) {
                        const shrinkMsg = ctx.formatSubjectCatalogShrinkConfirmDe(
                            ctx.subjectCatalogShrinkReport(
                                typeof ctx.load === 'function' ? ctx.load().subjects || [] : [],
                                backupSubjects
                            ),
                            ''
                        );
                        if (shrinkMsg) backupMsg += '\n\n' + shrinkMsg;
                    }
                    const ok = await ctx.dlgConfirm(backupMsg, {
                        title: 'Backup importieren',
                        okText: 'Importieren',
                        cancelText: 'Abbrechen',
                        danger: true
                    });
                    if (!ok) return;
                    if (typeof window.ms365BrowserBackup.applyBackupAsync === 'function') {
                        await window.ms365BrowserBackup.applyBackupAsync(obj);
                    } else {
                        window.ms365BrowserBackup.applyBackup(obj);
                    }
                    location.reload();
                    return;
                }
                const isV2 =
                    obj.version >= 2 && obj.core && obj.structure && obj.match;
                const partialWarn =
                    'Diese Datei ist kein vollständiges Browser-Backup.\n\n' +
                    'Übernommen werden nur Stammdaten' +
                    (isV2 ? ' (Schuldaten v2)' : '') +
                    '.\n' +
                    'Nicht übernommen: E-Mail-Muster, Schul-Apps-Einstellungen, Dashboard, Power Automate und Werkzeugstände.\n\n' +
                    'Für die Demo auf einem anderen PC bitte „Browser-Backup“ aus der Datensicherung verwenden.\n\n' +
                    'Trotzdem nur Stammdaten importieren?';
                const okPartial = await ctx.dlgConfirm(partialWarn, {
                    title: isV2 ? 'Nur Stammdaten (v2)' : 'Nur Stammdaten',
                    okText: 'Nur Stammdaten',
                    cancelText: 'Abbrechen',
                    danger: true
                });
                if (!okPartial) {
                    ctx.setSummary('Import abgebrochen – bitte vollständiges Browser-Backup verwenden.', 'warn');
                    return;
                }
                const importSubjects = ctx.subjectsFromImportPayload(obj);
                if (importSubjects.length) {
                    const okSub = await ctx.confirmSubjectCatalogIfShrinks(
                        importSubjects,
                        isV2 ? 'JSON-Import (Schuldaten v2)' : 'JSON-Import (Stammdaten)'
                    );
                    if (!okSub) {
                        ctx.setSummary('Import abgebrochen – Fächerliste würde verkleinert worden sein.', 'warn');
                        return;
                    }
                }
                try {
                    if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.importJson === 'function') {
                        window.ms365AppDataV2.importJson(obj);
                    }
                } catch (errImport) {
                    ctx.setSummary('Import fehlgeschlagen: ' + (errImport?.message || String(errImport)), 'warn');
                    return;
                }
                const saved = isV2 ? ctx.load() : ctx.save(obj);
                if (ctx.schoolNameInput) ctx.schoolNameInput.value = ctx.normStr(saved.schoolName || '');
                if (ctx.taSubjects) ctx.taSubjects.value = (saved.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
                if (ctx.taArges) ctx.taArges.value = (saved.arges || []).map((x) => `${x.code};${x.name || ''};${(x.subjects || []).join(',')}`.trim()).join('\n');
                if (ctx.taTeachers) ctx.taTeachers.value = ctx.teachersToLines(saved.teachers || []);
                ctx.setAdministrationGroups(ctx.adminGroupsFromSettings(saved));
                if (ctx.taAdmin) ctx.taAdmin.value = (saved.admin || []).map((x) => `${x.role || ''};${x.name || ''};${x.email || ''}`.trim()).join('\n');
                if (ctx.taAdminRoles) ctx.taAdminRoles.value = (saved.adminRoles || []).map((x) => `${x.code || ''};${x.name || ''}`.trim()).join('\n');
                if (ctx.selSgaMode) ctx.selSgaMode.value = ctx.normStr(saved.sgaMode || 'group').toLowerCase() === 'distribution' ? 'distribution' : 'group';
                if (ctx.taSga) ctx.taSga.value = ctx.sgaToLines(saved.sga || []);
                if (ctx.taStudents) ctx.taStudents.value = ctx.studentsToLines(saved.students || []);
                if (ctx.taStudentCouncil) ctx.taStudentCouncil.value = ctx.studentCouncilToLines(saved.studentCouncil || []);
                if (ctx.taClasses) ctx.taClasses.value = (saved.classes || []).map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim()).join('\n');
                ctx.renderSubjectsTableFromTextarea();
                ctx.renderArgesTableFromTextarea();
                ctx.renderTeachersTableFromTextarea();
                ctx.renderAdminRolesTableFromTextarea();
                ctx.renderAdminTableFromTextarea();
                ctx.renderSgaTableFromTextarea();
                ctx.renderStudentsTableFromTextarea();
                ctx.renderStudentCouncilTableFromTextarea();
                ctx.renderClassesTableFromTextarea();
                const yImp = ctx.getDisplayedSchoolYearLabel() || ctx.currentSchoolYearLabel();
                ctx.setSummary(
                    `Import OK: schulweit ${(saved.subjects || []).length} Fächer, ${(saved.arges || []).length} ARGEs, ${(saved.admin || []).length} Verwaltung, ${(saved.sga || []).length} SGA-Einträge, ${(saved.teachers || []).length} Lehrkräfte — für Schuljahr ${yImp}: ${(saved.students || []).length} Schüler, ${(saved.studentCouncil || []).length} Schülervertretung, ${(saved.classes || []).length} Klassen.`,
                    'ok'
                );
            } catch (err) {
                const written = err && err.details && err.details.written;
                if (Array.isArray(written) && written.indexOf('ms365-schooltool-data-v2') >= 0) {
                    location.reload();
                    return;
                }
                ctx.setSummary('Import fehlgeschlagen: ' + (err?.message || String(err)), 'warn');
            } finally {
                ctx.fileImport.value = '';
            }
        });
    }

    if (ctx.domainInput) {
        ctx.domainInput.addEventListener('input', () => ctx.scheduleAutoSave());
        ctx.domainInput.addEventListener('change', () => ctx.scheduleAutoSave());
    }
    if (ctx.verifiedDomainsInput) {
        ctx.verifiedDomainsInput.addEventListener('input', () => ctx.scheduleAutoSave());
        ctx.verifiedDomainsInput.addEventListener('change', () => ctx.scheduleAutoSave());
    }
    if (ctx.schoolNameInput) {
        ctx.schoolNameInput.addEventListener('input', () => {
            ctx.clearStoredSchoolWideGroupMatches();
            ctx.refreshSchoolWideGroupNamingPlaceholders();
            ctx.scheduleAutoSave();
        });
        ctx.schoolNameInput.addEventListener('change', () => {
            ctx.clearStoredSchoolWideGroupMatches();
            ctx.refreshSchoolWideGroupNamingPlaceholders();
            ctx.scheduleAutoSave();
        });
    }
    // Kein Standard-Abschlussjahr mehr in den Stammdaten
}
