/**
 * tenant-settings-ui-bind-setup.js
 * @param {Record<string, any>} ctx
 */
export function installTenantBindSetup(ctx) {
        ctx.parseLinesToSubjects = window.ms365TenantSettingsParseSubjectsLines;
        ctx.parseLinesToArges = window.ms365TenantSettingsParseArgesLines;
        ctx.parseLinesToTeachers = window.ms365TenantSettingsParseTeachersLines;
        ctx.parseLinesToAdmin = window.ms365TenantSettingsParseAdminLines;
        ctx.parseLinesToAdminRoles = window.ms365TenantSettingsParseAdminRolesLines;
        ctx.parseLinesToStudents = window.ms365TenantSettingsParseStudentsLines;
        ctx.parseLinesToStudentCouncil = window.ms365TenantSettingsParseStudentCouncilLines;
        ctx.parseLinesToClasses = window.ms365TenantSettingsParseClassesLines;
        ctx.parseLinesToSga = window.ms365TenantSettingsParseSgaLines;
        ctx.load = window.ms365TenantSettingsLoad;
        ctx.save = function (settings) {
            return ctx.writeSchoolData(settings, { reason: 'tenant-register' }).settings;
        };

        ctx.taSubjects = document.getElementById('tenantSubjectsLines');
        ctx.subjectsTbody = document.getElementById('tenantSubjectsTableBody');
        ctx.subjectsTable = ctx.subjectsTbody ? ctx.subjectsTbody.closest('table') : null;
        ctx.subjectsSortState = { key: null, dir: 1 };
        ctx.btnAddSubjectRow = document.getElementById('tenantSubjectsAddRow');
        ctx.btnSubjectsImportM365 = document.getElementById('tenantSubjectsImportM365');
        ctx.btnEnrichSubjects = document.getElementById('tenantBtnEnrichSubjects');
        ctx.taArges = document.getElementById('tenantArgesLines');
        ctx.argesTbody = document.getElementById('tenantArgesTableBody');
        ctx.argesTable = ctx.argesTbody ? ctx.argesTbody.closest('table') : null;
        ctx.argesSortState = { key: null, dir: 1 };
        ctx.argesTableEditFocus = null;
        ctx.ARGE_TABLE_EDIT_FIELDS = ['code', 'name', 'subjects', 'headName', 'headEmail'];
        ctx.btnVerifyArgesGraph = document.getElementById('tenantBtnVerifyArgesGraph');
        ctx.btnAddArgeRow = document.getElementById('tenantArgesAddRow');
        ctx.taTeachers = document.getElementById('tenantTeachersLines');
        ctx.teachersTbody = document.getElementById('tenantTeachersTableBody');
        ctx.teachersTable = ctx.teachersTbody ? ctx.teachersTbody.closest('table') : null;
        ctx.btnAddTeacherRow = document.getElementById('tenantTeachersAddRow');
        ctx.btnVerifyTeachersGraph = document.getElementById('tenantBtnVerifyTeachersGraph');
        ctx.lehrerGroupMatchCell = document.getElementById('tenantLehrerGroupMatchCell');
        ctx.btnVerifyLehrerGroup = document.getElementById('tenantBtnVerifyLehrerGroup');
        ctx.btnCreateLehrerGroup = document.getElementById('tenantBtnCreateLehrerGroup');
        ctx.btnSyncLehrerGroupMembers = document.getElementById('tenantBtnSyncLehrerGroupMembers');
        ctx.btnUnmatchLehrerGroup = document.getElementById('tenantBtnUnmatchLehrerGroup');
        ctx.btnOpenLehrerGroup = document.getElementById('tenantBtnOpenLehrerGroup');
        ctx.inpLehrerGroupSearch = document.getElementById('tenantLehrerGroupSearch');
        ctx.btnSearchLehrerGroup = document.getElementById('tenantBtnSearchLehrerGroup');
        ctx.lehrerGroupSearchResults = document.getElementById('tenantLehrerGroupSearchResults');
        ctx.schuelerGroupMatchCell = document.getElementById('tenantSchuelerGroupMatchCell');
        ctx.btnVerifySchuelerGroup = document.getElementById('tenantBtnVerifySchuelerGroup');
        ctx.btnCreateSchuelerGroup = document.getElementById('tenantBtnCreateSchuelerGroup');
        ctx.btnSyncSchuelerGroupMembers = document.getElementById('tenantBtnSyncSchuelerGroupMembers');
        ctx.btnUnmatchSchuelerGroup = document.getElementById('tenantBtnUnmatchSchuelerGroup');
        ctx.btnOpenSchuelerGroup = document.getElementById('tenantBtnOpenSchuelerGroup');
        ctx.inpSchuelerGroupSearch = document.getElementById('tenantSchuelerGroupSearch');
        ctx.btnSearchSchuelerGroup = document.getElementById('tenantBtnSearchSchuelerGroup');
        ctx.schuelerGroupSearchResults = document.getElementById('tenantSchuelerGroupSearchResults');
        ctx.btnTeachersExportCsv = document.getElementById('tenantTeachersExportCsv');
        ctx.btnTeachersExportXlsx = document.getElementById('tenantTeachersExportXlsx');
        /** @type {{ key: 'code'|'name'|'email'|null, dir: 1|-1 }} */
        ctx.teachersSortState = { key: null, dir: 1 };
        ctx.taAdminBundle = document.getElementById('tenantAdminBundleLines');
        ctx.taAdmin = document.getElementById('tenantAdminLines');
        ctx.adminTbody = document.getElementById('tenantAdminTableBody');
        ctx.taAdminRoles = document.getElementById('tenantAdminRoleLines');
        ctx.adminRolesTbody = document.getElementById('tenantAdminRolesTableBody');
        ctx.btnAdminRolesDefaults = document.getElementById('tenantAdminRolesDefaults');
        ctx.btnVerifyVerwaltungGraph = document.getElementById('tenantBtnVerifyVerwaltungGraph');
        ctx.tenantAdminAudienceBoardHost = document.getElementById('tenantAdminAudienceBoard');
        ctx.tenantVerwaltungStructureTreeHost = document.getElementById(
            'tenantVerwaltungStructureTree'
        );
        ctx.audienceGroupsSnapshot = [];
        /** @type {object[]|null} Verwaltung ohne Textfeld-Bundle */
        ctx.administrationGroupsSnapshot = null;
        ctx.audienceMembershipsSnapshot = [];
        /** @type {{ refresh: () => void } | null} */
        ctx.audienceBoardRef = null;
        /** @type {{ refresh: () => void }|null} */
        ctx.verwaltungTreeRef = null;
        ctx.selSgaMode = document.getElementById('tenantSgaMode');
        ctx.taSga = document.getElementById('tenantSgaLines');
        ctx.sgaTbody = document.getElementById('tenantSgaTableBody');
        ctx.sgaTable = ctx.sgaTbody ? ctx.sgaTbody.closest('table') : null;
        ctx.sgaSortState = { key: null, dir: 1 };
        ctx.btnAddSgaRow = document.getElementById('tenantSgaAddRow');
        ctx.btnVerifySgaGraph = document.getElementById('tenantBtnVerifySgaGraph');
        ctx.sgaGroupMatchCell = document.getElementById('tenantSgaGroupMatchCell');
        ctx.btnVerifySgaGroup = document.getElementById('tenantBtnVerifySgaGroup');
        ctx.btnCreateSgaGroup = document.getElementById('tenantBtnCreateSgaGroup');
        ctx.btnUnmatchSgaGroup = document.getElementById('tenantBtnUnmatchSgaGroup');
        ctx.btnOpenSgaGroup = document.getElementById('tenantBtnOpenSgaGroup');
        ctx.inpSgaGroupSearch = document.getElementById('tenantSgaGroupSearch');
        ctx.btnSearchSgaGroup = document.getElementById('tenantBtnSearchSgaGroup');
        ctx.sgaGroupSearchResults = document.getElementById('tenantSgaGroupSearchResults');
        ctx.inpSgaNewDisplayName = document.getElementById('tenantSgaNewDisplayName');
        ctx.inpSgaNewMailNick = document.getElementById('tenantSgaNewMailNick');
        ctx.inpSvNewDisplayName = document.getElementById('tenantSvNewDisplayName');
        ctx.inpSvNewMailNick = document.getElementById('tenantSvNewMailNick');
        ctx.sgaBoardTeachersHost = document.getElementById('tenantSgaBoardTeachers');
        ctx.sgaBoardStudentsHost = document.getElementById('tenantSgaBoardStudents');
        ctx.sgaBoardExternalsHost = document.getElementById('tenantSgaBoardExternals');
        ctx.tabMainSga = document.getElementById('tabMainSga');
        ctx.ENTRA_GROUP_DETAILS =
            'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/';
        ctx.taStudents = document.getElementById('tenantStudentsLines');
        ctx.studentsTbody = document.getElementById('tenantStudentsTableBody');
        ctx.studentsTable = ctx.studentsTbody ? ctx.studentsTbody.closest('table') : null;
        ctx.studentsClassFilter = document.getElementById('tenantStudentsClassFilter');
        ctx.studentsTextFilter = document.getElementById('tenantStudentsTextFilter');
        ctx.studentsFilterMeta = document.getElementById('tenantStudentsFilterMeta');
        ctx.btnAddStudentRow = document.getElementById('tenantStudentsAddRow');
        ctx.btnVerifyStudentsGraph = document.getElementById('tenantBtnVerifyStudentsGraph');
        ctx.btnStudentsExportCsv = document.getElementById('tenantStudentsExportCsv');
        ctx.btnStudentsExportXlsx = document.getElementById('tenantStudentsExportXlsx');
        /** @type {{ key: 'klasse'|'name'|'email'|null, dir: 1|-1 }} */
        ctx.studentsSortState = { key: null, dir: 1 };
        ctx.taStudentCouncil = document.getElementById('tenantStudentCouncilLines');
        ctx.studentCouncilTbody = document.getElementById('tenantStudentCouncilTableBody');
        ctx.studentCouncilTable = ctx.studentCouncilTbody ? ctx.studentCouncilTbody.closest('table') : null;
        ctx.studentCouncilSortState = { key: null, dir: 1 };
        ctx.btnAddStudentCouncilRow = document.getElementById('tenantStudentCouncilAddRow');
        ctx.btnVerifyStudentCouncilGraph = document.getElementById('tenantBtnVerifyStudentCouncilGraph');
        ctx.studentCouncilGroupMatchCell = document.getElementById('tenantStudentCouncilGroupMatchCell');
        ctx.btnVerifyStudentCouncilGroup = document.getElementById('tenantBtnVerifyStudentCouncilGroup');
        ctx.btnCreateStudentCouncilGroup = document.getElementById('tenantBtnCreateStudentCouncilGroup');
        ctx.btnUnmatchStudentCouncilGroup = document.getElementById('tenantBtnUnmatchStudentCouncilGroup');
        ctx.btnOpenStudentCouncilGroup = document.getElementById('tenantBtnOpenStudentCouncilGroup');
        ctx.inpStudentCouncilGroupSearch = document.getElementById('tenantStudentCouncilGroupSearch');
        ctx.btnSearchStudentCouncilGroup = document.getElementById('tenantBtnSearchStudentCouncilGroup');
        ctx.studentCouncilGroupSearchResults = document.getElementById('tenantStudentCouncilGroupSearchResults');
        ctx.studentCouncilPickSlotsHost = document.getElementById('tenantStudentCouncilPickSlots');
        ctx.btnStudentCouncilAddPickSlot = document.getElementById('tenantStudentCouncilAddPickSlot');
        ctx.tabMainSchuelervertretung = document.getElementById('tabMainSchuelervertretung');
        ctx.STUDENT_COUNCIL_PICK_MIN_SLOTS = 3;
        ctx.taClasses = document.getElementById('tenantClassesLines');
        ctx.classesTbody = document.getElementById('tenantClassesTableBody');
        ctx.classesTable = ctx.classesTbody ? ctx.classesTbody.closest('table') : null;
        ctx.classesSortState = { key: null, dir: 1 };
        /** Nach Render: Zelle für Inline-Edit fokussieren (z. B. neue Zeile, Tabulator). */
        ctx.classesTableEditFocus = null;
        ctx.CLASS_TABLE_EDIT_FIELDS = ['code', 'year', 'name', 'headName', 'headEmail'];
        ctx.btnAddClassRow = document.getElementById('tenantClassesAddRow');
        ctx.btnVerifyClassesGraph = document.getElementById('tenantBtnVerifyClassesGraph');
        ctx.btnClassesImportM365 = document.getElementById('tenantClassesImportM365');
        ctx.btnEnrichClassesFromGroups = document.getElementById('tenantBtnEnrichClassesFromGroups');
        ctx.klassenlehrerPanelApi = null;
        ctx.fileSubjects = document.getElementById('tenantSubjectsImportFile');
        ctx.fileArges = document.getElementById('tenantArgesImportFile');
        ctx.btnArgesImportM365 = document.getElementById('tenantArgesImportM365');
        ctx.fileTeachers = document.getElementById('tenantTeachersImportFile');
        ctx.fileStudents = document.getElementById('tenantStudentsImportFile');
        ctx.fileClasses = document.getElementById('tenantClassesImportFile');
        ctx.btnSubjectsTpl = document.getElementById('tenantSubjectsTemplateXlsx');
        ctx.btnArgesTpl = document.getElementById('tenantArgesTemplateXlsx');
        ctx.btnTeachersTpl = document.getElementById('tenantTeachersTemplateXlsx');
        ctx.btnStudentsTpl = document.getElementById('tenantStudentsTemplateXlsx');
        ctx.btnStudentsFillDemoParents = document.getElementById('tenantStudentsFillDemoParents');
        ctx.btnClassesTpl = document.getElementById('tenantClassesTemplateXlsx');
        ctx.btnClassesExportCsv = document.getElementById('tenantClassesExportCsv');
        ctx.btnClassesExportXlsx = document.getElementById('tenantClassesExportXlsx');
        ctx.btnSaves = document.querySelectorAll('[data-tenant-action="save"]');
        ctx.btnReloads = document.querySelectorAll('[data-tenant-action="reload"]');
        ctx.btnExport = document.getElementById('tenantSettingsExport');
        ctx.btnExportHeader = document.getElementById('tenantSettingsExportHeader');
        ctx.fileImport = document.getElementById('tenantSettingsImportFile');
        ctx.btnClear = document.getElementById('tenantSettingsClear');
        ctx.summary = document.getElementById('tenantSettingsSummary');
        ctx._inpDefaultGradYear = null;
        ctx.schoolNameInput = document.getElementById('schoolName');
        ctx.domainInput = document.getElementById('schoolEmailDomain');
        ctx.verifiedDomainsInput = document.getElementById('schoolVerifiedEmailDomains');
        ctx.schoolYearSelect = document.getElementById('schoolYearSelect');

        /** Schuljahr Sep–Aug (kanonisch, siehe shared/utils/school-year.js). */
        ctx.currentSchoolYearLabel = function currentSchoolYearLabel(now) {
            try {
                if (typeof window !== 'undefined' && window.ms365SchoolYear && typeof window.ms365SchoolYear.currentSchoolYearLabel === 'function') {
                    return window.ms365SchoolYear.currentSchoolYearLabel(now);
                }
            } catch {
                /* fallback */
            }
            const d = now instanceof Date && !isNaN(now.getTime()) ? now : new Date();
            const y = d.getMonth() < 8 ? d.getFullYear() - 1 : d.getFullYear();
            return String(y) + '/' + String(y + 1).slice(2);
        }

        ctx.getDisplayedSchoolYearLabel = function getDisplayedSchoolYearLabel() {
            try {
                if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                    const c = window.ms365AppDataV2.getContainer();
                    const cur = c && c.years ? String(c.years.current || '').trim() : '';
                    if (cur) return cur;
                }
            } catch {
                // ignore
            }
            if (ctx.schoolYearSelect) return String(ctx.schoolYearSelect.value || '').trim();
            return '';
        }

        ctx.autoSaveTimer = null;
        ctx.listsHydrated = false;
        /** True nur nach einer echten Eingabe. Seitenwechsel darf leere Felder nicht über volle Listen speichern. */
        ctx.listEditSincePaint = false;
        ctx.registerSaveDirty = false;
        if (typeof ctx.syncRegisterSaveBar === 'function') ctx.syncRegisterSaveBar();
        ctx.__syncGuard = 0;

        ctx.dispatchTenantSettingsChanged = function dispatchTenantSettingsChanged(saved, reason) {
            try {
                if (ctx.__syncGuard) return;
                const r = String(reason || '');
                if (r === 'render' || r === 'manual-save' || r === 'writeback') {
                    ctx.setStandSummary(r);
                } else if (r === 'autosave') {
                    ctx.updateRegisterLastSavedMeta(r);
                    ctx.refreshStandSummaryIfVisible(r);
                }
                window.dispatchEvent(
                    new CustomEvent('ms365-tenant-settings-changed', {
                        detail: { settings: saved, reason: r }
                    })
                );
            } catch {
                // ignore
            }
        }

        ctx.autoSaveNow = function autoSaveNow() {
            // Solange die Textfelder nicht aus dem Speicher gefüllt sind, würde ein
            // Autosave die gerade importierten Listen mit dem leeren Formular überschreiben
            // und nur das Organigramm stehen lassen.
            if (!ctx.listsHydrated) return;
            const subjects = typeof ctx.parseLinesToSubjects === 'function' ? ctx.parseLinesToSubjects(ctx.taSubjects ? ctx.taSubjects.value : '') : [];
            const arges = typeof ctx.parseLinesToArges === 'function' ? ctx.parseLinesToArges(ctx.taArges ? ctx.taArges.value : '') : [];
            const teachers = typeof ctx.parseLinesToTeachers === 'function' ? ctx.parseLinesToTeachers(ctx.taTeachers ? ctx.taTeachers.value : '') : [];
            const admin = ctx.getAdminFromTextarea();
            const adminRoles = ctx.getAdminRolesFromTextarea();
            const sga = typeof ctx.parseLinesToSga === 'function' ? ctx.parseLinesToSga(ctx.taSga ? ctx.taSga.value : '') : [];
            const sgaMode = ctx.selSgaMode ? ctx.normStr(ctx.selSgaMode.value || 'group').toLowerCase() : 'group';
            const students = typeof ctx.parseLinesToStudents === 'function' ? ctx.parseLinesToStudents(ctx.taStudents ? ctx.taStudents.value : '') : [];
            const studentCouncil =
                typeof ctx.parseLinesToStudentCouncil === 'function'
                    ? ctx.parseLinesToStudentCouncil(ctx.taStudentCouncil ? ctx.taStudentCouncil.value : '')
                    : [];
            const classes = typeof ctx.parseLinesToClasses === 'function' ? ctx.parseLinesToClasses(ctx.taClasses ? ctx.taClasses.value : '') : [];
            const schoolName = ctx.schoolNameInput ? ctx.normStr(ctx.schoolNameInput.value || '') : '';
            const domain =
                typeof window.ms365GetSchoolDomainNoAt === 'function' ? window.ms365GetSchoolDomainNoAt() : '';
            const verifiedEmailDomains = ctx.verifiedDomainsInput ? ctx.normStr(ctx.verifiedDomainsInput.value || '') : '';
            const administration = ctx.getAdministrationEntries();
            const prevSettings = typeof ctx.load === 'function' ? ctx.load() : {};
            const keepRows = function (parsed, prev) {
                const next = Array.isArray(parsed) ? parsed : [];
                const old = Array.isArray(prev) ? prev : [];
                if (ctx.listEditSincePaint) return next;
                if (next.length || !old.length) return next;
                return old;
            };
            const schoolNameOut =
                !ctx.listEditSincePaint && !schoolName && prevSettings && prevSettings.schoolName
                    ? ctx.normStr(prevSettings.schoolName)
                    : schoolName;
            const saved = ctx.save({
                schoolName: schoolNameOut,
                domain,
                verifiedEmailDomains,
                subjects: keepRows(subjects, prevSettings && prevSettings.subjects),
                arges: keepRows(arges, prevSettings && prevSettings.arges),
                teachers: keepRows(teachers, prevSettings && prevSettings.teachers),
                administration,
                admin,
                adminRoles,
                verwaltungAudienceGroups: ctx.audienceGroupsSnapshot,
                adminAudienceMemberships: ctx.audienceMembershipsSnapshot,
                verwaltungOrganigramm:
                    prevSettings && prevSettings.verwaltungOrganigramm
                        ? prevSettings.verwaltungOrganigramm
                        : { version: 1, nodes: [] },
                schulOrganigramm:
                    prevSettings && prevSettings.schulOrganigramm
                        ? prevSettings.schulOrganigramm
                        : { version: 1, nodes: [] },
                sgaMode,
                sga: keepRows(sga, prevSettings && prevSettings.sga),
                students: keepRows(students, prevSettings && prevSettings.students),
                studentCouncil: keepRows(studentCouncil, prevSettings && prevSettings.studentCouncil),
                classes: keepRows(classes, prevSettings && prevSettings.classes)
            });
            ctx.dispatchTenantSettingsChanged(saved, 'autosave');
            if (typeof ctx.markRegisterClean === 'function') ctx.markRegisterClean();
        }

}
