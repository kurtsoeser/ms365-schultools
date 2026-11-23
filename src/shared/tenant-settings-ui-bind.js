/**
 * tenant-settings-ui bindUi orchestration (Analyse 02 Phase C / Tabs-Split).
 * Assembles shared ctx and installs tab/panel modules.
 */
import {
    CLASS_LIST_SHEET_NAME,
    CLASS_LIST_XLSX_FILENAME,
    classListTemplateAoa,
    classesRowsToAoa,
    classesToLines as classesToLinesPure,
    downloadClassesCsv
} from './tenant-settings-ui-classes.js';
import { studentsToLines as studentsToLinesPure } from './tenant-settings-ui-students.js';
import { enrichStudentLinesWithDemoParents } from './demo-eltern-enrich.js';
import {
    resolveStudentRepSyncSource,
    sgaStudentsFromStudentCouncil,
    studentCouncilFromSgaStudents
} from './sga-student-council-sync.js';
import {
    SUBJECT_LIST_SHEET_NAME,
    SUBJECT_LIST_XLSX_FILENAME,
    subjectListTemplateAoa,
    subjectsFromSpreadsheetJsonRows,
    subjectsToLines as subjectsToLinesPure
} from './tenant-settings-ui-subjects.js';
import {
    pdfImportTargetFromFilename,
    resolveRegisterImportTarget,
    spreadsheetImportTargetFromAoa,
    REGISTER_IMPORT_TAB_BTN,
    REGISTER_IMPORT_TARGET_LABEL
} from './register-webuntis-file-import.js';
import { openSubjectPdfImportReview } from './subject-pdf-import-review-ui.js';
import {
    ADMIN_TIER_SCHULLEITUNG,
    ADMIN_TIER_VERWALTUNG,
    inferAdminTierForRole,
    normalizeAdminTier
} from './administration-audience-logic.js';
import {
    BUILTIN_AUDIENCE_SCHULLEITUNG,
    ensureAdminAudienceOnSettings,
    normalizeAdminAudienceMemberships,
    normalizeVerwaltungAudienceGroups
} from './administration-audience-groups.js';
import { mountTenantAdminAudienceBoard } from './tenant-admin-audience-board-ui.js';
import { mountVerwaltungStammdatenTree } from './verwaltung-stammdaten-tree-ui.js';
import { mountTenantSchoolEmailPatternUi } from './tenant-teacher-email-pattern-ui.js';
import { mountTenantIntranetQuickSync } from './tenant-intranet-quick-sync.js';
import { mountTenantSchoolProfileUi, reloadTenantSchoolProfileForm } from './tenant-school-profile-ui.js';
import { mountTenantStammdatenSetupNav } from './tenant-stammdaten-setup-nav.js';
import { mountTenantPlanerSitesUi } from './tenant-planer-sites-ui.js';
import { writeSchoolData } from './stammdaten-write.js';
import { isStammdatenLocalFirst, localFirstStatusMessage } from './stammdaten-local-first.js';
import {
    enrichClassesFromLinkedGroups,
    enrichClassesKvFromGraphOwners,
    collectPriorClassFieldsByCode,
    deriveClassEntryFromGroup,
    fillRegisterRowsHeadFromGroupOwners
} from './class-list-enrich.js';
import {
    enrichSubjectsFromCatalogSources,
    subjectCatalogShrinkReport,
    formatSubjectCatalogShrinkConfirmDe,
    subjectsFromImportPayload,
    deriveSubjectEntryFromGroup
} from './subject-list-enrich.js';
import { pickEntraGroupsMulti } from './entra-group-picker.js';
import { pickEntraUser } from './entra-user-picker.js';
import { studentBelongsToClassRow } from './class-student-match.js';
import {
    summarizeWebuntisMergeResult,
    diffTeachersImport,
    summarizeTeachersDiff,
    mergeTeachersImportLists,
    replaceTeachersImportList
} from './webuntis-stammdaten-wizard-logic.js';
import { mountTenantImportPanel, mergeSisImportViaAdapter } from './tenant-panel-import.js';
import { notifyTenantStudentsChanged } from './tenant-panel-schueler.js';
import {
    applyTabRegisterStatus,
    personMatchTabStatus,
    classMatchTabStatus,
    catalogLinkTabStatus,
    updateRegisterLastSavedMeta,
    formatRegisterLastSavedSuffix
} from './tenant-register-tab-status.js';
import { renderSchoolYearSelect, bindSchoolYearControls } from './school-year-ui.js';
import { summarizeClassTeaching } from './klassenlehrer-stammdaten-logic.js';
import { mountKlassenlehrerPanel } from './tenant-panel-klassenlehrer-ui.js';
import './schul-organigramm-stammdaten-sync.js';

import * as tenantSettingsUiUtils from './tenant-settings-ui-utils.js';
import { installTenantSettingsImportApis } from './tenant-settings-ui-import.js';
import { installTenantBindSetup } from './tenant-settings-ui-bind-setup.js';
import { installTenantArgesPanel } from './tenant-settings-ui-arges.js';
import { installTenantStatusPanel } from './tenant-settings-ui-status.js';
import { installTenantDirectoryHelpers } from './tenant-settings-ui-directory.js';
import { installTenantSubjectsPanel } from './tenant-settings-ui-subjects-panel.js';
import { installTenantTeachersPanel } from './tenant-settings-ui-teachers.js';
import { installTenantAdminPanel } from './tenant-settings-ui-admin.js';
import { installTenantSgaPanel } from './tenant-settings-ui-sga.js';
import { installTenantStudentCouncilPanel } from './tenant-settings-ui-student-council.js';
import { installTenantSchoolGroupsPanel } from './tenant-settings-ui-school-groups.js';
import { installTenantTeachersLifecycle } from './tenant-settings-ui-teachers-lifecycle.js';
import { installTenantImportFlows } from './tenant-settings-ui-import-flows.js';
import { installTenantStudentsPanel } from './tenant-settings-ui-students-panel.js';
import { installTenantClassesPanel } from './tenant-settings-ui-classes-panel.js';
import { installTenantRenderStorage } from './tenant-settings-ui-render.js';
import { installTenantBindEvents } from './tenant-settings-ui-bind-events.js';

installTenantSettingsImportApis();

function createCtx() {
    return Object.assign({}, tenantSettingsUiUtils, {
        CLASS_LIST_SHEET_NAME,
        CLASS_LIST_XLSX_FILENAME,
        classListTemplateAoa,
        classesRowsToAoa,
        classesToLinesPure,
        downloadClassesCsv,
        studentsToLinesPure,
        enrichStudentLinesWithDemoParents,
        resolveStudentRepSyncSource,
        sgaStudentsFromStudentCouncil,
        studentCouncilFromSgaStudents,
        SUBJECT_LIST_SHEET_NAME,
        SUBJECT_LIST_XLSX_FILENAME,
        subjectListTemplateAoa,
        subjectsFromSpreadsheetJsonRows,
        subjectsToLinesPure,
        pdfImportTargetFromFilename,
        resolveRegisterImportTarget,
        spreadsheetImportTargetFromAoa,
        REGISTER_IMPORT_TAB_BTN,
        REGISTER_IMPORT_TARGET_LABEL,
        openSubjectPdfImportReview,
        ADMIN_TIER_SCHULLEITUNG,
        ADMIN_TIER_VERWALTUNG,
        inferAdminTierForRole,
        normalizeAdminTier,
        BUILTIN_AUDIENCE_SCHULLEITUNG,
        ensureAdminAudienceOnSettings,
        normalizeAdminAudienceMemberships,
        normalizeVerwaltungAudienceGroups,
        mountTenantAdminAudienceBoard,
        mountVerwaltungStammdatenTree,
        mountTenantSchoolEmailPatternUi,
        mountTenantIntranetQuickSync,
        mountTenantSchoolProfileUi,
        reloadTenantSchoolProfileForm,
        mountTenantStammdatenSetupNav,
        mountTenantPlanerSitesUi,
        writeSchoolData,
        isStammdatenLocalFirst,
        localFirstStatusMessage,
        enrichClassesFromLinkedGroups,
        enrichClassesKvFromGraphOwners,
        collectPriorClassFieldsByCode,
        deriveClassEntryFromGroup,
        fillRegisterRowsHeadFromGroupOwners,
        enrichSubjectsFromCatalogSources,
        subjectCatalogShrinkReport,
        formatSubjectCatalogShrinkConfirmDe,
        subjectsFromImportPayload,
        deriveSubjectEntryFromGroup,
        pickEntraGroupsMulti,
        pickEntraUser,
        studentBelongsToClassRow,
        summarizeWebuntisMergeResult,
        diffTeachersImport,
        summarizeTeachersDiff,
        mergeTeachersImportLists,
        replaceTeachersImportList,
        mountTenantImportPanel,
        mergeSisImportViaAdapter,
        notifyTenantStudentsChanged,
        applyTabRegisterStatus,
        personMatchTabStatus,
        classMatchTabStatus,
        catalogLinkTabStatus,
        updateRegisterLastSavedMeta,
        formatRegisterLastSavedSuffix,
        renderSchoolYearSelect,
        bindSchoolYearControls,
        summarizeClassTeaching,
        mountKlassenlehrerPanel
    });
}

function bindUi() {
    const ctx = createCtx();
    ctx.exposeSubjectCatalogGuardOnWindow();
    const form = document.getElementById('tenantSettingsForm');
    if (!form) return;
    ctx.form = form;

    // Pane-Switcher sofort – unabhängig von Settings-Load/Save (sonst wirken Klicks tot).
    try {
        ctx.mountTenantStammdatenSetupNav();
    } catch {
        /* ignore */
    }

    if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
        return;
    }

    ctx.mountTenantSchoolEmailPatternUi();
    ctx.mountTenantSchoolProfileUi();
    // IT-Nav-Sichtbarkeit zuerst, dann Pane-Switcher (merkt sich z. B. Sicherungsbibliothek).
    ctx.mountTenantIntranetQuickSync();
    ctx.mountTenantStammdatenSetupNav();
    ctx.mountTenantPlanerSitesUi();
    if (ctx.isStammdatenLocalFirst()) {
        const banner = document.getElementById('tenantLocalFirstBanner');
        if (banner) {
            banner.hidden = true;
            banner.textContent = '';
        }
    }

    installTenantStatusPanel(ctx);
    installTenantBindSetup(ctx);
    installTenantArgesPanel(ctx);
    installTenantDirectoryHelpers(ctx);
    installTenantSubjectsPanel(ctx);
    installTenantTeachersPanel(ctx);
    installTenantAdminPanel(ctx);
    installTenantSgaPanel(ctx);
    installTenantStudentCouncilPanel(ctx);
    installTenantSchoolGroupsPanel(ctx);
    installTenantTeachersLifecycle(ctx);
    installTenantImportFlows(ctx);
    installTenantStudentsPanel(ctx);
    installTenantClassesPanel(ctx);
    installTenantRenderStorage(ctx);
    installTenantBindEvents(ctx);
}

// Im gebauten Bundle steht dieses Modul im <head> und läuft vor den defer-Skripten
// (tenant-settings-core.js). Ohne Warten bricht bindUi() dort still ab und die Reiter bleiben leer.
let bindUiDone = false;
function tenantCoreReady() {
    return (
        typeof window.ms365TenantSettingsLoad === 'function' &&
        typeof window.ms365TenantSettingsSave === 'function'
    );
}
function bindUiWhenReady() {
    if (bindUiDone) return true;
    if (document.getElementById('tenantSettingsForm') && !tenantCoreReady()) return false;
    bindUiDone = true;
    bindUi();
    return true;
}
if (!bindUiWhenReady()) {
    document.addEventListener('DOMContentLoaded', bindUiWhenReady, { once: true });
    window.addEventListener('load', bindUiWhenReady, { once: true });
    let bindUiTries = 0;
    const bindUiPoll = setInterval(function () {
        bindUiTries += 1;
        if (bindUiWhenReady() || bindUiTries > 100) clearInterval(bindUiPoll);
    }, 100);
}

export { bindUi, bindUiWhenReady };
