/**
 * tenant-settings-ui-render-storage.js
 * @param {Record<string, any>} ctx
 */
export function installTenantRenderStorageCore(ctx) {
    ctx.paintCoreLists = function paintCoreLists(s) {
        const src = s && typeof s === 'object' ? s : {};
        if (ctx.schoolNameInput) ctx.schoolNameInput.value = ctx.normStr(src.schoolName || '');
        if (ctx.taSubjects) {
            delete ctx.taSubjects.dataset.ms365UserEdited;
            ctx.taSubjects.value = (src.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
        }
        if (ctx.taArges) {
            delete ctx.taArges.dataset.ms365UserEdited;
            ctx.taArges.value = (src.arges || [])
                .map((x) => `${x.code};${x.name || ''};${(x.subjects || []).join(',')}`.trim())
                .join('\n');
        }
        if (ctx.taTeachers) {
            delete ctx.taTeachers.dataset.ms365UserEdited;
            ctx.taTeachers.value = ctx.teachersToLines(src.teachers || []);
        }
        if (ctx.taStudents) {
            delete ctx.taStudents.dataset.ms365UserEdited;
            ctx.taStudents.value = ctx.studentsToLines(src.students || []);
        }
        ctx.listEditSincePaint = false;
        if (typeof ctx.markRegisterClean === 'function') ctx.markRegisterClean();
        if (ctx.taClasses) {
            delete ctx.taClasses.dataset.ms365UserEdited;
            ctx.taClasses.value = (src.classes || [])
                .map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim())
                .join('\n');
        }
    }

    ctx.renderFromStorage = function renderFromStorage() {
        ctx.listsHydrated = false;
        let s = ctx.load();
        ctx.paintCoreLists(s);
        ctx.listsHydrated = true;
        try {
            const enriched = ctx.applyClassEnrichFromLocalLinks(s.classes || [], {
                persist: true,
                summary: false,
                reason: 'class-enrich-on-load'
            });
            if (enriched && enriched.classes) s = ctx.load();
            ctx.applySubjectEnrichFromSources(s.subjects || [], {
                persist: true,
                summary: false,
                reason: 'subject-enrich-on-load'
            });
        } catch {
            /* Listen bleiben aus dem ersten Paint stehen */
        }
        s = ctx.load();
        // Domain in UI-Feld zurückschreiben (wird auch von school-domain.js genutzt)
        try {
            if (ctx.schoolNameInput) ctx.schoolNameInput.value = ctx.normStr(s.schoolName || '');
            if (ctx.domainInput) ctx.domainInput.value = ctx.normStr(s.domain || '');
            if (ctx.verifiedDomainsInput) {
                const vd = Array.isArray(s.verifiedEmailDomains)
                    ? s.verifiedEmailDomains.join(', ')
                    : ctx.normStr(s.verifiedEmailDomains || '');
                ctx.verifiedDomainsInput.value = vd;
            }
            if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
                const d = ctx.normStr(s.domain || '').replace(/^@+/, '');
                if (d) window.ms365SetSchoolDomainNoAt(d);
            }
        } catch {
            // ignore
        }
        if (ctx.taSubjects) {
            ctx.taSubjects.value = (s.subjects || []).map((x) => `${x.code};${x.name || ''}`.trim()).join('\n');
        }
        if (ctx.taArges) {
            ctx.taArges.value = (s.arges || [])
                .map((x) => `${x.code};${x.name || ''};${(x.subjects || []).join(',')}`.trim())
                .join('\n');
        }
        if (ctx.taTeachers) {
            ctx.taTeachers.value = ctx.teachersToLines(s.teachers || []);
        }
        ctx.syncAdministrationSnapshotFromSettings(s);
        ctx.setAdministrationGroups(ctx.administrationGroupsSnapshot || []);
        ctx.syncAudienceSnapshotsFromSettings(s);
        if (ctx.audienceBoardRef) ctx.audienceBoardRef.refresh();
        if (ctx.taAdmin) {
            ctx.taAdmin.value = (s.admin || [])
                .map((x) => `${x.role || ''};${x.name || ''};${x.email || ''}`.trim())
                .join('\n');
        }
        if (ctx.taAdminRoles) {
            ctx.taAdminRoles.value = (s.adminRoles || [])
                .map((x) => `${x.code || ''};${x.name || ''}`.trim())
                .join('\n');
        }
        if (ctx.selSgaMode) {
            ctx.selSgaMode.value = ctx.normStr(s.sgaMode || 'group').toLowerCase() === 'distribution' ? 'distribution' : 'group';
        }
        if (ctx.taSga) ctx.taSga.value = ctx.sgaToLines(s.sga || []);
        if (ctx.taStudents) {
            ctx.taStudents.value = ctx.studentsToLines(s.students || []);
        }
        if (ctx.taStudentCouncil) {
            ctx.sgaStudentCouncilSyncLock = true;
            try {
                ctx.taStudentCouncil.value = ctx.studentCouncilToLines(s.studentCouncil || []);
            } finally {
                ctx.sgaStudentCouncilSyncLock = false;
            }
        }
        if (ctx.taClasses) {
            ctx.taClasses.value = (s.classes || [])
                .map((x) => `${x.code || ''};${x.year || ''};${x.name || ''};${x.headName || ''};${x.headEmail || ''}`.trim())
                .join('\n');
        }
        ctx.renderSubjectsTableFromTextarea();
        ctx.renderArgesTableFromTextarea();
        ctx.renderTeachersTableFromTextarea();
        ctx.renderAdminRolesTableFromTextarea();
        ctx.renderAdminTableFromTextarea();
        ctx.renderSgaTableFromTextarea();
        ctx.renderStudentsTableFromTextarea();
        ctx.reconcileSgaAndStudentCouncil();
        ctx.renderSgaBoard();
        ctx.renderStudentCouncilTableFromTextarea();
        ctx.renderStudentCouncilPickBoard();
        ctx.renderClassesTableFromTextarea();
        ctx.restoreSchoolWideGroupNamingFields();
        ctx.restoreStoredSchoolWideGroupMatchStatus();
        ctx.restoreLehrerSammelGroupMatchStatus();
        ctx.restoreSchuelerSammelGroupMatchStatus();
        ctx.renderSchoolYearSelectFromV2();
        ctx.renderStatusOverview();
        ctx.reloadTenantSchoolProfileForm();
        const orgNodes = (s.schulOrganigramm && s.schulOrganigramm.nodes) || [];
        const orgClassCount = orgNodes.filter((n) => {
            const kind = ctx.normStr(n && n.refKind).toLowerCase();
            const type = ctx.normStr(n && n.type).toLowerCase();
            return kind === 'class' || type === 'class';
        }).length;
        if (!(s.teachers || []).length && !(s.classes || []).length && orgClassCount) {
            ctx.setSummary(
                'In der Landkarte liegen noch ' +
                    orgClassCount +
                    ' Klassen. Die Stammdaten-Listen sind leer. Bitte die Datei ms365-browser-backup-….json erneut importieren.',
                'warn'
            );
        }
        ctx.dispatchTenantSettingsChanged(s, 'render');
    }

    ctx.fieldBlankAndUntouched = function fieldBlankAndUntouched(el) {
        if (!el) return false;
        if (el.dataset && el.dataset.ms365UserEdited === '1') return false;
        return !String(el.value || '').trim();
    }

    /**
     * Landkarte liest die Listen aus dem Speicher, die Reiter zeigen das Formular.
     * Wenn das Formular leer ist und der Speicher Zeilen hat, Felder und Tabellen nachziehen.
     */
    ctx.refillListsFromStorageIfBlank = function refillListsFromStorageIfBlank() {
        let s = null;
        try {
            s = typeof ctx.load === 'function' ? ctx.load() : null;
        } catch {
            return;
        }
        if (!s) return;
        const need =
            (ctx.fieldBlankAndUntouched(ctx.taTeachers) && (s.teachers || []).length) ||
            (ctx.fieldBlankAndUntouched(ctx.taClasses) && (s.classes || []).length) ||
            (ctx.fieldBlankAndUntouched(ctx.taSubjects) && (s.subjects || []).length) ||
            (ctx.fieldBlankAndUntouched(ctx.taArges) && (s.arges || []).length) ||
            (ctx.fieldBlankAndUntouched(ctx.taStudents) && (s.students || []).length);
        if (!need) return;
        ctx.listsHydrated = false;
        ctx.paintCoreLists(s);
        ctx.listsHydrated = true;
        try {
            ctx.renderSubjectsTableFromTextarea();
            ctx.renderArgesTableFromTextarea();
            ctx.renderTeachersTableFromTextarea();
            ctx.renderStudentsTableFromTextarea();
            ctx.renderClassesTableFromTextarea();
        } catch {
            /* Tabellen optional */
        }
    }

    ctx.renderSchoolYearSelectFromV2 = function renderSchoolYearSelectFromV2() {
        ctx.renderSchoolYearSelect(ctx.schoolYearSelect || document.getElementById('schoolYearSelect'));
    }

    ctx.downloadJson = function downloadJson(filename, obj) {
        const blob = new Blob([JSON.stringify(obj, null, 2)], { type: 'application/json;charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(() => URL.revokeObjectURL(url), 250);
    }

    ctx.importFileToRows = function importFileToRows(file, onRows) {
        ctx.importSpreadsheetFileToJsonRows(file, onRows, (msg) => ctx.setSummary(msg, 'warn'));
    }

    ctx.closeListImportDrop = function closeListImportDrop(fileInputEl) {
        try {
            const drop = fileInputEl && fileInputEl.closest('details.ts-list-import-drop');
            if (drop) drop.open = false;
        } catch {
            /* ignore */
        }
    }

}
