/**
 * tenant-settings-ui-bind-events-people.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindEventsPeople(ctx) {
    if (ctx.taAdminBundle) {
        ctx.taAdminBundle.addEventListener('input', () => {
            if (typeof window.ms365TenantSettingsParseAdminGroupsLines === 'function') {
                const parsed = window.ms365TenantSettingsParseAdminGroupsLines(ctx.taAdminBundle.value);
                ctx.administrationGroupsSnapshot = ctx.mergeAdminRolePolicyFromSnapshot(
                    parsed,
                    ctx.administrationGroupsSnapshot
                );
            }
            ctx.renderAdminRolesTableFromTextarea();
            ctx.renderAdminTableFromTextarea();
            ctx.refreshAdminAudienceBoardUi();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.taAdmin) {
        ctx.taAdmin.addEventListener('input', () => ctx.renderAdminTableFromTextarea());
        ctx.taAdmin.addEventListener('input', () => ctx.scheduleAutoSave());
    }
    if (ctx.taAdminRoles) {
        ctx.taAdminRoles.addEventListener('input', () => ctx.renderAdminRolesTableFromTextarea());
        ctx.taAdminRoles.addEventListener('input', () => ctx.renderAdminTableFromTextarea());
        ctx.taAdminRoles.addEventListener('input', () => ctx.scheduleAutoSave());
    }
    if (ctx.btnAdminRolesDefaults) {
        ctx.btnAdminRolesDefaults.addEventListener('click', () => {
            const defaults =
                typeof window.ms365TenantSettingsDefaultAdminRoles === 'function'
                    ? window.ms365TenantSettingsDefaultAdminRoles()
                    : [];
            const rows = ctx.groupsToDisplayRows(ctx.getAdministrationGroups());
            const seen = new Set(
                rows.map(function (row) {
                    return ctx.normStr(row.name).toLowerCase();
                })
            );
            defaults.forEach(function (d) {
                const name = ctx.normStr(d && d.name);
                const key = name.toLowerCase();
                if (!name || seen.has(key)) return;
                seen.add(key);
                rows.push({
                    code: ctx.normCode(d && d.code),
                    name: name,
                    personName: '',
                    email: '',
                    tier: ctx.inferAdminTierForRole(d)
                });
            });
            ctx.setAdministrationGroups(ctx.displayRowsToGroups(rows));
            ctx.renderAdminTableFromTextarea();
            ctx.refreshAdminAudienceBoardUi();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.btnVerifyVerwaltungGraph && !ctx.btnVerifyVerwaltungGraph.dataset.tenantVerifyVerwaltungBound) {
        ctx.btnVerifyVerwaltungGraph.dataset.tenantVerifyVerwaltungBound = '1';
        ctx.btnVerifyVerwaltungGraph.addEventListener('click', async () => {
            await ctx.runVerifyVerwaltungGraphBulk();
        });
    }
    if (ctx.selSgaMode) {
        ctx.selSgaMode.addEventListener('change', () => {
            ctx.clearStoredSchoolWideGroupMatches();
            ctx.scheduleAutoSave();
        });
    }
    ctx.bindSchoolWideGroupNamingField(ctx.inpSgaNewDisplayName, 'sga');
    ctx.bindSchoolWideGroupNamingField(ctx.inpSgaNewMailNick, 'sga');
    ctx.bindSchoolWideGroupNamingField(ctx.inpSvNewDisplayName, 'studentCouncil');
    ctx.bindSchoolWideGroupNamingField(ctx.inpSvNewMailNick, 'studentCouncil');
    if (ctx.taSga) {
        ctx.taSga.addEventListener('input', () => {
            ctx.sgaSortState.key = null;
            ctx.sgaSortState.dir = 1;
            ctx.renderSgaTableFromTextarea();
            ctx.renderSgaBoard();
            ctx.scheduleAutoSave();
        });
        ctx.taSga.addEventListener('change', () => {
            const part = ctx.partitionSgaForBoard();
            ctx.pushStudentCouncilFromSgaBoard(part.board);
        });
    }
    if (ctx.tabMainSga && !ctx.tabMainSga.dataset.tenantSgaBoardRefreshBound) {
        ctx.tabMainSga.dataset.tenantSgaBoardRefreshBound = '1';
        ctx.tabMainSga.addEventListener('click', () => {
            ctx.renderSgaBoard();
        });
    }
    if (ctx.btnAddSgaRow) {
        ctx.btnAddSgaRow.addEventListener('click', () => {
            const all = ctx.getSgaFromTextarea();
            all.push({ scope: 'teacher', name: '', email: '' });
            ctx.setSgaTextareaFromRows(all);
            ctx.renderSgaTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }

    if (ctx.taStudents) {
        let studentsTableRenderTimer = null;
        ctx.taStudents.addEventListener('input', () => {
            ctx.taStudents.dataset.ms365UserEdited = '1';
            ctx.studentsSortState.key = null;
            ctx.studentsSortState.dir = 1;
            if (studentsTableRenderTimer) clearTimeout(studentsTableRenderTimer);
            studentsTableRenderTimer = setTimeout(() => {
                studentsTableRenderTimer = null;
                ctx.renderStudentsTableFromTextarea();
            }, 250);
        });
        ctx.taStudents.addEventListener('input', () => ctx.scheduleAutoSave());
    }
    if (ctx.studentsClassFilter && !ctx.studentsClassFilter.dataset.studentsFilterBound) {
        ctx.studentsClassFilter.dataset.studentsFilterBound = '1';
        ctx.studentsClassFilter.addEventListener('change', () => {
            ctx.renderStudentsTableFromTextarea();
        });
    }
    if (ctx.studentsTextFilter && !ctx.studentsTextFilter.dataset.studentsFilterBound) {
        ctx.studentsTextFilter.dataset.studentsFilterBound = '1';
        let studentsTextFilterTimer = null;
        ctx.studentsTextFilter.addEventListener('input', () => {
            if (studentsTextFilterTimer) clearTimeout(studentsTextFilterTimer);
            studentsTextFilterTimer = setTimeout(() => {
                studentsTextFilterTimer = null;
                ctx.renderStudentsTableFromTextarea();
            }, 150);
        });
    }
    if (ctx.studentsTable && !ctx.studentsTable.dataset.studentsSortBound) {
        ctx.studentsTable.dataset.studentsSortBound = '1';
        ctx.studentsTable.querySelectorAll('th.is-sortable').forEach((th) => {
            th.addEventListener('click', () => ctx.applyStudentsSort(th.dataset.sortKey));
            th.addEventListener('keydown', (e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                    e.preventDefault();
                    ctx.applyStudentsSort(th.dataset.sortKey);
                }
            });
        });
    }
    if (ctx.btnAddStudentRow) {
        ctx.btnAddStudentRow.addEventListener('click', () => {
            const prevRows = ctx.getStudentsFromTextarea();
            const all = prevRows.slice();
            all.push({ klasse: '', name: '', email: '' });
            ctx.setStudentsTextareaFromRows(all);
            ctx.renderStudentsTableFromTextarea();
            ctx.renderStudentSammelLifecyclePanel(prevRows, all);
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.btnStudentsExportCsv) {
        ctx.btnStudentsExportCsv.addEventListener('click', () => {
            const api = window.ms365StudentListImport;
            const rows = ctx.getStudentsFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Schülerliste ist leer – nichts zu exportieren.', 'warn');
                return;
            }
            if (!api || typeof api.exportCsv !== 'function' || !api.exportCsv(rows)) {
                ctx.setSummary('CSV-Export fehlgeschlagen.', 'warn');
                return;
            }
            ctx.setSummary('Schülerliste als CSV exportiert (' + rows.length + ').', 'ok');
        });
    }
    if (ctx.btnStudentsExportXlsx) {
        ctx.btnStudentsExportXlsx.addEventListener('click', () => {
            const api = window.ms365StudentListImport;
            const rows = ctx.getStudentsFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Schülerliste ist leer – nichts zu exportieren.', 'warn');
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
            ctx.setSummary('Schülerliste als XLSX exportiert (' + rows.length + ').', 'ok');
        });
    }
    if (ctx.btnVerifyStudentsGraph && !ctx.btnVerifyStudentsGraph.dataset.tenantVerifyStudentsBound) {
        ctx.btnVerifyStudentsGraph.dataset.tenantVerifyStudentsBound = '1';
        ctx.btnVerifyStudentsGraph.addEventListener('click', async () => {
            await ctx.runVerifyGraphDirectoryRows(
                ctx.getStudentsFromTextarea(),
                function (r) {
                    return r.email;
                },
                'Schüler:innen',
                ctx.btnVerifyStudentsGraph,
                function () {
                    ctx.renderStudentsTableFromTextarea();
                }
            );
        });
    }
    if (ctx.taStudentCouncil) {
        ctx.taStudentCouncil.addEventListener('input', () => {
            ctx.studentCouncilSortState.key = null;
            ctx.studentCouncilSortState.dir = 1;
            ctx.renderStudentCouncilTableFromTextarea();
            ctx.renderStudentCouncilPickBoard();
            ctx.scheduleAutoSave();
        });
        ctx.taStudentCouncil.addEventListener('change', () => {
            ctx.pushSgaStudentsFromCouncilRows(ctx.getStudentCouncilFromTextarea());
        });
    }
    if (ctx.btnStudentCouncilAddPickSlot && !ctx.btnStudentCouncilAddPickSlot.dataset.tenantSvPickBound) {
        ctx.btnStudentCouncilAddPickSlot.dataset.tenantSvPickBound = '1';
        ctx.btnStudentCouncilAddPickSlot.addEventListener('click', () => {
            const all = ctx.getStudentCouncilFromTextarea().slice();
            all.push({ klasse: '', name: '', email: '' });
            ctx.setStudentCouncilTextareaFromRows(all);
            ctx.renderStudentCouncilTableFromTextarea();
            ctx.renderStudentCouncilPickBoard();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.tabMainSchuelervertretung && !ctx.tabMainSchuelervertretung.dataset.tenantSvPickRefreshBound) {
        ctx.tabMainSchuelervertretung.dataset.tenantSvPickRefreshBound = '1';
        ctx.tabMainSchuelervertretung.addEventListener('click', () => {
            ctx.renderStudentCouncilPickBoard();
        });
    }
    if (ctx.btnAddStudentCouncilRow) {
        ctx.btnAddStudentCouncilRow.addEventListener('click', () => {
            const all = ctx.getStudentCouncilFromTextarea();
            all.push({ klasse: '', name: '', email: '' });
            ctx.setStudentCouncilTextareaFromRows(all);
            ctx.renderStudentCouncilTableFromTextarea();
            ctx.renderStudentCouncilPickBoard();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.btnVerifySgaGraph && !ctx.btnVerifySgaGraph.dataset.tenantVerifySgaBound) {
        ctx.btnVerifySgaGraph.dataset.tenantVerifySgaBound = '1';
        ctx.btnVerifySgaGraph.addEventListener('click', async () => {
            await ctx.runVerifyGraphDirectoryRows(
                ctx.getSgaFromTextarea(),
                function (r) { return r.email; },
                'SGA-Mitglieder',
                ctx.btnVerifySgaGraph,
                function () { ctx.renderSgaTableFromTextarea(); }
            );
        });
    }
    if (ctx.btnVerifyStudentCouncilGraph && !ctx.btnVerifyStudentCouncilGraph.dataset.tenantVerifyStudentCouncilBound) {
        ctx.btnVerifyStudentCouncilGraph.dataset.tenantVerifyStudentCouncilBound = '1';
        ctx.btnVerifyStudentCouncilGraph.addEventListener('click', async () => {
            await ctx.runVerifyGraphDirectoryRows(
                ctx.getStudentCouncilFromTextarea(),
                function (r) { return r.email; },
                'Schülervertretung',
                ctx.btnVerifyStudentCouncilGraph,
                function () { ctx.renderStudentCouncilTableFromTextarea(); }
            );
        });
    }

    if (ctx.btnVerifySgaGroup && !ctx.btnVerifySgaGroup.dataset.tenantVerifySgaGroupBound) {
        ctx.btnVerifySgaGroup.dataset.tenantVerifySgaGroupBound = '1';
        ctx.btnVerifySgaGroup.addEventListener('click', async () => {
            if (ctx.btnVerifySgaGroup.disabled) return;
            try {
                ctx.btnVerifySgaGroup.disabled = true;
                ctx.btnVerifySgaGroup.setAttribute('aria-busy', 'true');
                ctx.setSummary('SGA-Gruppe: Prüfe in Microsoft 365 …', 'warn');
                await ctx.verifySgaGroupExistence();
            } catch (e) {
                ctx.setSummary('SGA-Gruppe: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnVerifySgaGroup.disabled = false;
                ctx.btnVerifySgaGroup.removeAttribute('aria-busy');
            }
        });
    }

    if (ctx.btnCreateSgaGroup && !ctx.btnCreateSgaGroup.dataset.tenantCreateSgaGroupBound) {
        ctx.btnCreateSgaGroup.dataset.tenantCreateSgaGroupBound = '1';
        ctx.btnCreateSgaGroup.addEventListener('click', async () => {
            if (ctx.btnCreateSgaGroup.disabled) return;
            const fold = document.getElementById('tenantSgaGroupFold');
            if (fold && !fold.open) {
                fold.open = true;
                document.getElementById('tenantSgaNewDisplayName')?.focus();
                return;
            }
            try {
                ctx.btnCreateSgaGroup.disabled = true;
                ctx.btnCreateSgaGroup.setAttribute('aria-busy', 'true');
                ctx.setSummary('SGA-Gruppe: Anlege in Microsoft 365 …', 'warn');
                await ctx.createSgaGroupExistence();
            } catch (e) {
                ctx.setSummary('SGA-Gruppe anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnCreateSgaGroup.disabled = false;
                ctx.btnCreateSgaGroup.removeAttribute('aria-busy');
            }
        });
    }

    if (ctx.btnVerifyStudentCouncilGroup && !ctx.btnVerifyStudentCouncilGroup.dataset.tenantVerifyStudentCouncilGroupBound) {
        ctx.btnVerifyStudentCouncilGroup.dataset.tenantVerifyStudentCouncilGroupBound = '1';
        ctx.btnVerifyStudentCouncilGroup.addEventListener('click', async () => {
            if (ctx.btnVerifyStudentCouncilGroup.disabled) return;
            try {
                ctx.btnVerifyStudentCouncilGroup.disabled = true;
                ctx.btnVerifyStudentCouncilGroup.setAttribute('aria-busy', 'true');
                ctx.setSummary('Schülervertretung: Prüfe in Microsoft 365 …', 'warn');
                await ctx.verifyStudentCouncilGroupExistence();
            } catch (e) {
                ctx.setSummary(
                    'Schülervertretung: ' + (e && e.message ? e.message : String(e)),
                    'warn'
                );
            } finally {
                ctx.btnVerifyStudentCouncilGroup.disabled = false;
                ctx.btnVerifyStudentCouncilGroup.removeAttribute('aria-busy');
            }
        });
    }

    if (ctx.btnCreateStudentCouncilGroup && !ctx.btnCreateStudentCouncilGroup.dataset.tenantCreateStudentCouncilGroupBound) {
        ctx.btnCreateStudentCouncilGroup.dataset.tenantCreateStudentCouncilGroupBound = '1';
        ctx.btnCreateStudentCouncilGroup.addEventListener('click', async () => {
            if (ctx.btnCreateStudentCouncilGroup.disabled) return;
            const fold = document.getElementById('tenantSvGroupFold');
            if (fold && !fold.open) {
                fold.open = true;
                document.getElementById('tenantSvNewDisplayName')?.focus();
                return;
            }
            try {
                ctx.btnCreateStudentCouncilGroup.disabled = true;
                ctx.btnCreateStudentCouncilGroup.setAttribute('aria-busy', 'true');
                ctx.setSummary('Schülervertretung: Anlege in Microsoft 365 …', 'warn');
                await ctx.createStudentCouncilGroupExistence();
            } catch (e) {
                ctx.setSummary(
                    'Schülervertretung anlegen: ' + (e && e.message ? e.message : String(e)),
                    'warn'
                );
            } finally {
                ctx.btnCreateStudentCouncilGroup.disabled = false;
                ctx.btnCreateStudentCouncilGroup.removeAttribute('aria-busy');
            }
        });
    }

    if (ctx.btnSearchStudentCouncilGroup && !ctx.btnSearchStudentCouncilGroup.dataset.tenantSearchSvGroupBound) {
        ctx.btnSearchStudentCouncilGroup.dataset.tenantSearchSvGroupBound = '1';
        ctx.btnSearchStudentCouncilGroup.addEventListener('click', async () => {
            if (ctx.btnSearchStudentCouncilGroup.disabled) return;
            try {
                ctx.btnSearchStudentCouncilGroup.disabled = true;
                await ctx.searchSchoolWideGroups('studentCouncil');
            } catch (e) {
                ctx.setSummary(
                    'Schülervertretung Suche: ' + (e && e.message ? e.message : String(e)),
                    'warn'
                );
            } finally {
                ctx.btnSearchStudentCouncilGroup.disabled = false;
            }
        });
    }
    if (ctx.inpStudentCouncilGroupSearch && !ctx.inpStudentCouncilGroupSearch.dataset.tenantSearchSvEnterBound) {
        ctx.inpStudentCouncilGroupSearch.dataset.tenantSearchSvEnterBound = '1';
        ctx.inpStudentCouncilGroupSearch.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                if (ctx.btnSearchStudentCouncilGroup) ctx.btnSearchStudentCouncilGroup.click();
            }
        });
    }
    if (ctx.btnUnmatchStudentCouncilGroup && !ctx.btnUnmatchStudentCouncilGroup.dataset.tenantUnmatchSvBound) {
        ctx.btnUnmatchStudentCouncilGroup.dataset.tenantUnmatchSvBound = '1';
        ctx.btnUnmatchStudentCouncilGroup.addEventListener('click', () => {
            ctx.unmatchSchoolWideGroup('studentCouncil');
        });
    }
    if (ctx.btnOpenStudentCouncilGroup && !ctx.btnOpenStudentCouncilGroup.dataset.tenantOpenSvBound) {
        ctx.btnOpenStudentCouncilGroup.dataset.tenantOpenSvBound = '1';
        ctx.btnOpenStudentCouncilGroup.addEventListener('click', () => {
            ctx.openSchoolWideGroupInEntra('studentCouncil');
        });
    }

    if (ctx.btnSearchSgaGroup && !ctx.btnSearchSgaGroup.dataset.tenantSearchSgaGroupBound) {
        ctx.btnSearchSgaGroup.dataset.tenantSearchSgaGroupBound = '1';
        ctx.btnSearchSgaGroup.addEventListener('click', async () => {
            if (ctx.btnSearchSgaGroup.disabled) return;
            try {
                ctx.btnSearchSgaGroup.disabled = true;
                await ctx.searchSchoolWideGroups('sga');
            } catch (e) {
                ctx.setSummary('SGA Suche: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnSearchSgaGroup.disabled = false;
            }
        });
    }
    if (ctx.inpSgaGroupSearch && !ctx.inpSgaGroupSearch.dataset.tenantSearchSgaEnterBound) {
        ctx.inpSgaGroupSearch.dataset.tenantSearchSgaEnterBound = '1';
        ctx.inpSgaGroupSearch.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                if (ctx.btnSearchSgaGroup) ctx.btnSearchSgaGroup.click();
            }
        });
    }
    if (ctx.btnUnmatchSgaGroup && !ctx.btnUnmatchSgaGroup.dataset.tenantUnmatchSgaBound) {
        ctx.btnUnmatchSgaGroup.dataset.tenantUnmatchSgaBound = '1';
        ctx.btnUnmatchSgaGroup.addEventListener('click', () => {
            ctx.unmatchSchoolWideGroup('sga');
        });
    }
    if (ctx.btnOpenSgaGroup && !ctx.btnOpenSgaGroup.dataset.tenantOpenSgaBound) {
        ctx.btnOpenSgaGroup.dataset.tenantOpenSgaBound = '1';
        ctx.btnOpenSgaGroup.addEventListener('click', () => {
            ctx.openSchoolWideGroupInEntra('sga');
        });
    }

}
