/**
 * tenant-settings-ui-bind-events-groups.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindEventsGroups(ctx) {
    if (ctx.btnVerifyLehrerGroup && !ctx.btnVerifyLehrerGroup.dataset.tenantVerifyLehrerGroupBound) {
        ctx.btnVerifyLehrerGroup.dataset.tenantVerifyLehrerGroupBound = '1';
        ctx.btnVerifyLehrerGroup.addEventListener('click', async () => {
            if (ctx.btnVerifyLehrerGroup.disabled) return;
            try {
                ctx.btnVerifyLehrerGroup.disabled = true;
                ctx.btnVerifyLehrerGroup.setAttribute('aria-busy', 'true');
                await ctx.verifyLehrerSammelGroupExistence();
            } catch (e) {
                ctx.setSummary('Lehrer:innen-Gruppe: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnVerifyLehrerGroup.disabled = false;
                ctx.btnVerifyLehrerGroup.removeAttribute('aria-busy');
            }
        });
    }
    if (ctx.btnCreateLehrerGroup && !ctx.btnCreateLehrerGroup.dataset.tenantCreateLehrerGroupBound) {
        ctx.btnCreateLehrerGroup.dataset.tenantCreateLehrerGroupBound = '1';
        ctx.btnCreateLehrerGroup.addEventListener('click', async () => {
            if (ctx.btnCreateLehrerGroup.disabled) return;
            try {
                ctx.btnCreateLehrerGroup.disabled = true;
                ctx.btnCreateLehrerGroup.setAttribute('aria-busy', 'true');
                ctx.setSummary('Lehrer:innen-Sammelgruppe: Anlege in Microsoft 365 …', 'warn');
                await ctx.createLehrerSammelGroupExistence();
            } catch (e) {
                ctx.setSummary('Lehrer:innen-Gruppe anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnCreateLehrerGroup.disabled = false;
                ctx.btnCreateLehrerGroup.removeAttribute('aria-busy');
            }
        });
    }
    if (ctx.btnSyncLehrerGroupMembers && !ctx.btnSyncLehrerGroupMembers.dataset.tenantSyncLehrerMembersBound) {
        ctx.btnSyncLehrerGroupMembers.dataset.tenantSyncLehrerMembersBound = '1';
        ctx.btnSyncLehrerGroupMembers.addEventListener('click', async () => {
            if (ctx.btnSyncLehrerGroupMembers.disabled) return;
            try {
                ctx.btnSyncLehrerGroupMembers.disabled = true;
                ctx.btnSyncLehrerGroupMembers.setAttribute('aria-busy', 'true');
                await ctx.syncLehrerSammelGroupMembersAdditive();
            } catch (e) {
                ctx.setSummary('Lehrer:innen-Gruppe: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnSyncLehrerGroupMembers.disabled = false;
                ctx.btnSyncLehrerGroupMembers.removeAttribute('aria-busy');
            }
        });
    }
    if (ctx.btnSearchLehrerGroup && !ctx.btnSearchLehrerGroup.dataset.tenantSearchLehrerGroupBound) {
        ctx.btnSearchLehrerGroup.dataset.tenantSearchLehrerGroupBound = '1';
        ctx.btnSearchLehrerGroup.addEventListener('click', async () => {
            if (ctx.btnSearchLehrerGroup.disabled) return;
            try {
                ctx.btnSearchLehrerGroup.disabled = true;
                await ctx.searchLehrerSammelGroups();
            } catch (e) {
                ctx.setSummary('Lehrer:innen Suche: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnSearchLehrerGroup.disabled = false;
            }
        });
    }
    if (ctx.inpLehrerGroupSearch && !ctx.inpLehrerGroupSearch.dataset.tenantSearchLehrerEnterBound) {
        ctx.inpLehrerGroupSearch.dataset.tenantSearchLehrerEnterBound = '1';
        ctx.inpLehrerGroupSearch.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                if (ctx.btnSearchLehrerGroup) ctx.btnSearchLehrerGroup.click();
            }
        });
    }
    if (ctx.btnUnmatchLehrerGroup && !ctx.btnUnmatchLehrerGroup.dataset.tenantUnmatchLehrerBound) {
        ctx.btnUnmatchLehrerGroup.dataset.tenantUnmatchLehrerBound = '1';
        ctx.btnUnmatchLehrerGroup.addEventListener('click', () => {
            ctx.unmatchLehrerSammelGroup();
        });
    }
    if (ctx.btnOpenLehrerGroup && !ctx.btnOpenLehrerGroup.dataset.tenantOpenLehrerBound) {
        ctx.btnOpenLehrerGroup.dataset.tenantOpenLehrerBound = '1';
        ctx.btnOpenLehrerGroup.addEventListener('click', () => {
            ctx.openLehrerSammelGroupInEntra();
        });
    }

    if (ctx.btnVerifySchuelerGroup && !ctx.btnVerifySchuelerGroup.dataset.tenantVerifySchuelerGroupBound) {
        ctx.btnVerifySchuelerGroup.dataset.tenantVerifySchuelerGroupBound = '1';
        ctx.btnVerifySchuelerGroup.addEventListener('click', async () => {
            if (ctx.btnVerifySchuelerGroup.disabled) return;
            try {
                ctx.btnVerifySchuelerGroup.disabled = true;
                ctx.btnVerifySchuelerGroup.setAttribute('aria-busy', 'true');
                await ctx.verifySchuelerSammelGroupExistence();
            } catch (e) {
                ctx.setSummary('Schüler:innen-Gruppe: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnVerifySchuelerGroup.disabled = false;
                ctx.btnVerifySchuelerGroup.removeAttribute('aria-busy');
            }
        });
    }
    if (ctx.btnCreateSchuelerGroup && !ctx.btnCreateSchuelerGroup.dataset.tenantCreateSchuelerGroupBound) {
        ctx.btnCreateSchuelerGroup.dataset.tenantCreateSchuelerGroupBound = '1';
        ctx.btnCreateSchuelerGroup.addEventListener('click', async () => {
            if (ctx.btnCreateSchuelerGroup.disabled) return;
            try {
                ctx.btnCreateSchuelerGroup.disabled = true;
                ctx.btnCreateSchuelerGroup.setAttribute('aria-busy', 'true');
                ctx.setSummary('Schüler:innen-Sammelgruppe: Anlege in Microsoft 365 …', 'warn');
                await ctx.createSchuelerSammelGroupExistence();
            } catch (e) {
                ctx.setSummary('Schüler:innen-Gruppe anlegen: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnCreateSchuelerGroup.disabled = false;
                ctx.btnCreateSchuelerGroup.removeAttribute('aria-busy');
            }
        });
    }
    if (ctx.btnSyncSchuelerGroupMembers && !ctx.btnSyncSchuelerGroupMembers.dataset.tenantSyncSchuelerMembersBound) {
        ctx.btnSyncSchuelerGroupMembers.dataset.tenantSyncSchuelerMembersBound = '1';
        ctx.btnSyncSchuelerGroupMembers.addEventListener('click', async () => {
            if (ctx.btnSyncSchuelerGroupMembers.disabled) return;
            try {
                ctx.btnSyncSchuelerGroupMembers.disabled = true;
                ctx.btnSyncSchuelerGroupMembers.setAttribute('aria-busy', 'true');
                await ctx.syncSchuelerSammelGroupMembersAdditive();
            } catch (e) {
                ctx.setSummary('Schüler:innen-Gruppe: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnSyncSchuelerGroupMembers.disabled = false;
                ctx.btnSyncSchuelerGroupMembers.removeAttribute('aria-busy');
            }
        });
    }
    if (ctx.btnSearchSchuelerGroup && !ctx.btnSearchSchuelerGroup.dataset.tenantSearchSchuelerGroupBound) {
        ctx.btnSearchSchuelerGroup.dataset.tenantSearchSchuelerGroupBound = '1';
        ctx.btnSearchSchuelerGroup.addEventListener('click', async () => {
            if (ctx.btnSearchSchuelerGroup.disabled) return;
            try {
                ctx.btnSearchSchuelerGroup.disabled = true;
                await ctx.searchSchuelerSammelGroups();
            } catch (e) {
                ctx.setSummary('Schüler:innen Suche: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                ctx.btnSearchSchuelerGroup.disabled = false;
            }
        });
    }
    if (ctx.inpSchuelerGroupSearch && !ctx.inpSchuelerGroupSearch.dataset.tenantSearchSchuelerEnterBound) {
        ctx.inpSchuelerGroupSearch.dataset.tenantSearchSchuelerEnterBound = '1';
        ctx.inpSchuelerGroupSearch.addEventListener('keydown', (ev) => {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                if (ctx.btnSearchSchuelerGroup) ctx.btnSearchSchuelerGroup.click();
            }
        });
    }
    if (ctx.btnUnmatchSchuelerGroup && !ctx.btnUnmatchSchuelerGroup.dataset.tenantUnmatchSchuelerBound) {
        ctx.btnUnmatchSchuelerGroup.dataset.tenantUnmatchSchuelerBound = '1';
        ctx.btnUnmatchSchuelerGroup.addEventListener('click', () => {
            ctx.unmatchSchuelerSammelGroup();
        });
    }
    if (ctx.btnOpenSchuelerGroup && !ctx.btnOpenSchuelerGroup.dataset.tenantOpenSchuelerBound) {
        ctx.btnOpenSchuelerGroup.dataset.tenantOpenSchuelerBound = '1';
        ctx.btnOpenSchuelerGroup.addEventListener('click', () => {
            ctx.openSchuelerSammelGroupInEntra();
        });
    }

    ctx.btnTeacherLifecycleApply = document.getElementById('tenantTeacherLifecycleApply');
    if (ctx.btnTeacherLifecycleApply) {
        ctx.btnTeacherLifecycleApply.addEventListener('click', () => ctx.applyTeacherLifecyclePreview());
    }
    ctx.btnTeacherLifecycleDismiss = document.getElementById('tenantTeacherLifecycleDismiss');
    if (ctx.btnTeacherLifecycleDismiss) {
        ctx.btnTeacherLifecycleDismiss.addEventListener('click', () => {
            const host = document.getElementById('tenantTeacherLifecycle');
            if (host) host.hidden = true;
            ctx.lastTeacherLifecycleJoin = [];
        });
    }

    ctx.btnStudentSammelLifecycleApply = document.getElementById('tenantStudentSammelLifecycleApply');
    if (ctx.btnStudentSammelLifecycleApply) {
        ctx.btnStudentSammelLifecycleApply.addEventListener('click', () => ctx.applyStudentSammelLifecyclePreview());
    }
    ctx.btnStudentSammelLifecycleDismiss = document.getElementById('tenantStudentSammelLifecycleDismiss');
    if (ctx.btnStudentSammelLifecycleDismiss) {
        ctx.btnStudentSammelLifecycleDismiss.addEventListener('click', () => {
            const host = document.getElementById('tenantStudentSammelLifecycle');
            if (host) host.hidden = true;
            ctx.lastStudentSammelLifecycleJoin = [];
        });
    }

    ctx.btnLifecycleApply = document.getElementById('tenantStudentLifecycleApply');
    if (ctx.btnLifecycleApply) ctx.btnLifecycleApply.addEventListener('click', () => ctx.applyStudentLifecyclePreview());
    ctx.btnLifecycleDismiss = document.getElementById('tenantStudentLifecycleDismiss');
    if (ctx.btnLifecycleDismiss) {
        ctx.btnLifecycleDismiss.addEventListener('click', () => {
            const host = document.getElementById('tenantStudentLifecycle');
            if (host) host.hidden = true;
        });
    }

    if (ctx.taClasses) {
        ctx.taClasses.addEventListener('input', () => {
            ctx.taClasses.dataset.ms365UserEdited = '1';
            ctx.classesSortState.key = null;
            ctx.classesSortState.dir = 1;
            ctx.renderClassesTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }
    ctx.klassenlehrerPanelApi = ctx.mountKlassenlehrerPanel({
        getClasses: ctx.getClassesFromTextarea,
        getTeachers: ctx.getTeachersFromTextarea,
        getBelegung: function () {
            if (
                window.ms365AppDataV2 &&
                typeof window.ms365AppDataV2.getUnterrichtsbelegung === 'function'
            ) {
                return window.ms365AppDataV2.getUnterrichtsbelegung();
            }
            return null;
        },
        saveBelegung: function (snapshot) {
            if (
                window.ms365AppDataV2 &&
                typeof window.ms365AppDataV2.setUnterrichtsbelegung === 'function'
            ) {
                window.ms365AppDataV2.setUnterrichtsbelegung(snapshot);
            }
        },
        toast: function (msg, kind) {
            ctx.setSummary(msg, kind === 'ok' ? 'ok' : kind === 'warn' ? 'warn' : 'ok');
        },
        onChanged: function () {
            ctx.renderStatusOverview();
        }
    });
    if (ctx.btnAddClassRow) {
        ctx.btnAddClassRow.addEventListener('click', () => {
            const all = ctx.getClassesFromTextarea();
            all.push({ code: '', year: '', name: '', headName: '', headEmail: '' });
            ctx.setClassesTextareaFromRows(all);
            ctx.classesTableEditFocus = { rowIdx: all.length - 1, field: 'code' };
            ctx.renderClassesTableFromTextarea();
            ctx.scheduleAutoSave();
        });
    }
    if (ctx.btnVerifyClassesGraph && !ctx.btnVerifyClassesGraph.dataset.tenantVerifyClassesBound) {
        ctx.btnVerifyClassesGraph.dataset.tenantVerifyClassesBound = '1';
        ctx.btnVerifyClassesGraph.addEventListener('click', async () => {
            const rows = ctx.getClassesFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Keine Klassen zum Prüfen vorhanden.', 'warn');
                return;
            }
            const btn = ctx.btnVerifyClassesGraph;
            let found = 0;
            let missed = 0;
            let skipped = 0;
            const seen = new Set();
            try {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
                ctx.setSummary('Klassengruppen-Abgleich läuft …', 'warn');
                let token = null;
                for (let i = 0; i < rows.length; i++) {
                    const key = ctx.classMatchKey(rows[i]);
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
                    const res = await ctx.verifyClassGroupForRow(rows[i], token ? { token: token } : {});
                    if (res && res.skipped) skipped++;
                    else if (res && res.found) found++;
                    else missed++;
                    if (typeof ctx.graphApi().sleep === 'function' && i < rows.length - 1) {
                        await ctx.graphApi().sleep(300);
                    }
                }
                ctx.renderClassesTableFromTextarea();
                ctx.renderStatusOverview();
                ctx.setSummary(
                    'Klassen: ' + found + ' gefunden, ' + missed + ' nicht gefunden, ' + skipped + ' ohne Abgleichsschlüssel',
                    missed ? 'warn' : 'ok'
                );
            } catch (e) {
                ctx.setSummary('Klassengruppen-Abgleich: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        });
    }
    if (ctx.btnEnrichClassesFromGroups && !ctx.btnEnrichClassesFromGroups.dataset.tenantEnrichClassesBound) {
        ctx.btnEnrichClassesFromGroups.dataset.tenantEnrichClassesBound = '1';
        ctx.btnEnrichClassesFromGroups.addEventListener('click', async () => {
            const rows = ctx.getClassesFromTextarea();
            if (!rows.length) {
                ctx.setSummary('Keine Klassen zum Ergänzen vorhanden.', 'warn');
                return;
            }
            const btn = ctx.btnEnrichClassesFromGroups;
            try {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
                let working = rows;
                const local = ctx.applyClassEnrichFromLocalLinks(working, { persist: false });
                working = local.classes || working;
                const ctx = ctx.buildClassEnrichContext();
                const graphResult = await ctx.enrichClassesKvFromGraphOwners(working, {
                    classTeams: ctx.classTeams,
                    classGroupMatchByKey: ctx.classGroupMatchByKey,
                    teachers: ctx.getTeachersFromTextarea(),
                    fetchGroupOwners: ctx.graphApi().fetchGroupOwners,
                    getGraphToken: function () {
                        return ctx.graphApi().getGraphToken();
                    }
                });
                working = graphResult.classes || working;
                const cur = ctx.load();
                const saved = ctx.save(Object.assign({}, cur, { classes: working }));
                if (ctx.taClasses) ctx.taClasses.value = ctx.classesToLines(saved.classes || []);
                ctx.renderClassesTableFromTextarea();
                ctx.renderStatusOverview();
                const loc = local.stats || {};
                const gr = graphResult.stats || {};
                ctx.setSummary(
                    'Ergänzt: ' +
                        (loc.yearFilled || 0) +
                        ' Abschlussjahr(e), ' +
                        ((loc.kvFilled || 0) + (gr.kvFilled || 0)) +
                        ' Klassenvorstand-Daten (lokal + ggf. Graph-Besitzer). Bitte prüfen und speichern ist bereits übernommen.',
                    'ok'
                );
                ctx.dispatchTenantSettingsChanged(saved, 'class-enrich-groups');
            } catch (e) {
                ctx.setSummary('Ergänzen fehlgeschlagen: ' + (e && e.message ? e.message : String(e)), 'warn');
            } finally {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        });
    }

    ctx.btnStatusRefresh = document.getElementById('tenantStatusRefresh');
    if (ctx.btnStatusRefresh && ctx.btnStatusRefresh.dataset.entraRefreshBound !== '1') {
        ctx.btnStatusRefresh.dataset.entraRefreshBound = '1';
        ctx.btnStatusRefresh.addEventListener('click', () => ctx.renderStatusOverview());
    }

    window.addEventListener('ms365-browser-backup-applying', function () {
        ctx.listsHydrated = false;
        if (ctx.autoSaveTimer) {
            clearTimeout(ctx.autoSaveTimer);
            ctx.autoSaveTimer = null;
        }
    });

    window.addEventListener('ms365-tenant-settings-changed', function (ev) {
        const d = ev && ev.detail;
        if (!d) return;
        if (d.reason === 'render' || d.reason === 'autosave') return;
        if (d.source === 'browser-backup-import' || d.source === 'spo-auto-pull') {
            ctx.renderFromStorage();
        }
    });

    ctx.renderFromStorage();
    ctx.refillListsFromStorageIfBlank();
    setTimeout(ctx.refillListsFromStorageIfBlank, 0);
    setTimeout(ctx.refillListsFromStorageIfBlank, 600);
    ctx.tenantMainTabs = document.getElementById('tenantMainTabs');
    if (ctx.tenantMainTabs && ctx.tenantMainTabs.dataset.ms365ListRefillBound !== '1') {
        ctx.tenantMainTabs.dataset.ms365ListRefillBound = '1';
        ctx.tenantMainTabs.addEventListener('click', function () {
            ctx.refillListsFromStorageIfBlank();
        });
    }
    window.addEventListener('hashchange', function () {
        ctx.refillListsFromStorageIfBlank();
    });

    // Schritt 4 / 5: Tenant-IST, Match (Dropdown+Speichern), Differenz + Graph-Anlage
}
