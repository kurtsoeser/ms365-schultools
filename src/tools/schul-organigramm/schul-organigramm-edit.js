/**
 * Schul-Organigramm bearbeiten (schreibt Stammdaten).
 * Phase 2: Rollen · Phase 3: ARGEs · Phase 4: Klassen/KV, Vertretung/SGA.
 * Custom-Verwaltungsgruppen → verwaltungAudienceGroups (+ Organigramm-Section).
 *
 * Split: reine Move-/Extract-Module (siehe ARCHITECTURE.md). Dieses File ist der
 * Entry/Re-Export + createSchulOrganigrammEditor-Wiring.
 */
import { pickEntraUser } from '../../shared/entra-user-picker.js';
import {
    moveSchulOrganigrammArea,
    applySchulOrganigrammToStammdaten,
    isAreaNode,
    normalizeSchulOrganigramm
} from './schul-organigramm-logic.js';
import {
    normStr,
    normEmail,
    normCode,
    escapeHtml,
    suggestTeacherCode,
    ensureTeacherInList,
    stripEphemeralSchoolSettings,
    plainJsonClone,
    settingsForPersist,
    suggestLocalPersonEmail,
    isBlankPerson,
    personMatchesRole,
    loadSettings as loadSettingsDefault,
    saveSettings as saveSettingsDefault,
    toast,
    confirmDlg
} from './schul-organigramm-edit-utils.js';
import {
    editDialogHtml,
    areaM365ObjectsHtml,
    klassenchatsToolHref,
    editFootHtml
} from './schul-organigramm-edit-html.js';
import {
    findAdminRole,
    peopleForAdminRole,
    studentsForClass,
    guardiansForClass,
    applyArgeHeadFromGraphOwners,
    applyClassHeadFromGraphOwners,
    teachersForArge,
    teachersForClassChat,
    listPeopleForSchulOrgNode,
    listM365PeopleForArea,
    listMailboxDelegatePeopleForArea,
    previewPeopleForSchulOrgNode,
    SCHUL_ORG_PEOPLE_PREVIEW_LIMIT,
    roleCodeFromName,
    syncAdministrationFromAdmin
} from './schul-organigramm-edit-people.js';
import {
    moveSchulOrgRole,
    moveSchulOrgPerson,
    normalizeSchulOrgAreaPersonRole,
    ensureClassHeadsNamedFromTeachers,
    healClassHeadsFromAreaPeople,
    addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea
} from './schul-organigramm-edit-ops.js';
import {
    findArge,
    findTeacherByEmail,
    findClass,
    findStudentCouncil,
    findSga,
    sgaScopeLabel,
    fillTeacherSelect,
    wireTeacherHeadFields
} from './schul-organigramm-edit-find.js';
import {
    closeOverlay,
    onEsc,
    setActiveOverlay,
    toggleInlineEntraUserPicker,
    toggleInlineEntraGroupPicker,
    rememberSchulOrgRoleParent,
    getLastSchulOrgRoleParent
} from './schul-organigramm-edit-overlay.js';
import { createRoleUi } from './schul-organigramm-edit-role-ui.js';
import { createArgeUi } from './schul-organigramm-edit-arge-ui.js';
import { createClassUi } from './schul-organigramm-edit-class-ui.js';
import { createPersonListUi } from './schul-organigramm-edit-person-list-ui.js';
import { createAreaUi } from './schul-organigramm-edit-area-ui.js';
import { createCanvasUi } from './schul-organigramm-edit-canvas-ui.js';

export {
    findAdminRole,
    peopleForAdminRole,
    studentsForClass,
    guardiansForClass,
    applyArgeHeadFromGraphOwners,
    applyClassHeadFromGraphOwners,
    teachersForArge,
    teachersForClassChat,
    listPeopleForSchulOrgNode,
    listM365PeopleForArea,
    listMailboxDelegatePeopleForArea,
    previewPeopleForSchulOrgNode,
    SCHUL_ORG_PEOPLE_PREVIEW_LIMIT,
    moveSchulOrgRole,
    moveSchulOrgPerson,
    normalizeSchulOrgAreaPersonRole,
    ensureClassHeadsNamedFromTeachers,
    healClassHeadsFromAreaPeople,
    addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea,
    findArge,
    findTeacherByEmail,
    findClass,
    findStudentCouncil,
    findSga,
    sgaScopeLabel,
    rememberSchulOrgRoleParent,
    getLastSchulOrgRoleParent
};

/**
 * @param {{ onChanged?: function(): void, loadSettings?: function(): object, saveSettings?: function(object): object }} [deps]
 */
export function createSchulOrganigrammEditor(deps) {
    const onChangedCb =
        deps && typeof deps.onChanged === 'function' ? deps.onChanged : null;
    const loadSettingsFn =
        deps && typeof deps.loadSettings === 'function' ? deps.loadSettings : loadSettingsDefault;
    const saveSettingsFn =
        deps && typeof deps.saveSettings === 'function' ? deps.saveSettings : saveSettingsDefault;

    function notify() {
        if (onChangedCb) onChangedCb();
    }

    function loadSettings() {
        return loadSettingsFn();
    }

    function saveSettings(settings) {
        const clean = settingsForPersist(settings);
        return saveSettingsFn(clean);
    }

    /** @type {Record<string, any>} */
    const ctx = {
        loadSettings: loadSettings,
        saveSettings: saveSettings,
        toast: toast,
        confirmDlg: confirmDlg,
        notify: notify,
        closeOverlay: closeOverlay,
        onEsc: onEsc,
        setActiveOverlay: setActiveOverlay,
        escapeHtml: escapeHtml,
        normStr: normStr,
        normEmail: normEmail,
        normCode: normCode,
        editDialogHtml: editDialogHtml,
        editFootHtml: editFootHtml,
        areaM365ObjectsHtml: areaM365ObjectsHtml,
        klassenchatsToolHref: klassenchatsToolHref,
        toggleInlineEntraUserPicker: toggleInlineEntraUserPicker,
        toggleInlineEntraGroupPicker: toggleInlineEntraGroupPicker,
        rememberSchulOrgRoleParent: rememberSchulOrgRoleParent,
        getLastSchulOrgRoleParent: getLastSchulOrgRoleParent,
        pickEntraUser: pickEntraUser,
        suggestTeacherCode: suggestTeacherCode,
        ensureTeacherInList: ensureTeacherInList,
        stripEphemeralSchoolSettings: stripEphemeralSchoolSettings,
        plainJsonClone: plainJsonClone,
        settingsForPersist: settingsForPersist,
        suggestLocalPersonEmail: suggestLocalPersonEmail,
        isBlankPerson: isBlankPerson,
        personMatchesRole: personMatchesRole,
        softToast: null,
        promptFreeAreaName: null,
        persistOrganigrammLayout: null,
        defaultCanvasSlotForNewArea: null
    };

    const areaUi = createAreaUi(ctx);
    ctx.softToast = areaUi.softToast;
    ctx.promptFreeAreaName = areaUi.promptFreeAreaName;
    ctx.persistOrganigrammLayout = areaUi.persistOrganigrammLayout;
    ctx.defaultCanvasSlotForNewArea = areaUi.defaultCanvasSlotForNewArea;

    const roleUi = createRoleUi(ctx);
    const argeUi = createArgeUi(ctx);
    const classUi = createClassUi(ctx);
    const personListUi = createPersonListUi(ctx);
    ctx.openRoleEditor = roleUi.openRoleEditor;
    ctx.openArgeEditor = argeUi.openArgeEditor;
    ctx.openClassEditor = classUi.openClassEditor;
    ctx.openCouncilEditor = personListUi.openCouncilEditor;
    ctx.openSgaEditor = personListUi.openSgaEditor;
    const canvasUi = createCanvasUi(ctx);

    return {
        openRoleEditor: roleUi.openRoleEditor,
        openAreaEditor: areaUi.openAreaEditor,
        createRoleUnderSection: roleUi.createRoleUnderSection,
        rememberRoleParent: rememberSchulOrgRoleParent,
        getLastRoleParent: getLastSchulOrgRoleParent,
        createAudienceSection: areaUi.createAudienceSection,
        createAreaAt: areaUi.createAreaAt,
        openAddCanvasPerson: canvasUi.openAddCanvasPerson,
        openEditCanvasPerson: canvasUi.openEditCanvasPerson,
        pickAndAddPersonToArea: canvasUi.pickAndAddPersonToArea,
        removeAudienceSection: areaUi.removeAudienceSection,
        removeArea: areaUi.removeArea,
        openArgeEditor: argeUi.openArgeEditor,
        createArgeUnderSection: argeUi.createArgeUnderSection,
        openClassEditor: classUi.openClassEditor,
        createClassUnderSection: classUi.createClassUnderSection,
        openCouncilEditor: personListUi.openCouncilEditor,
        createCouncilUnderSection: personListUi.createCouncilUnderSection,
        openSgaEditor: personListUi.openSgaEditor,
        createSgaUnderSection: personListUi.createSgaUnderSection,
        close: closeOverlay,
        movePersonToNode: personListUi.movePersonToNode,
        moveRoleToArea: function (roleNodeId, newParentId, liveLayoutOrBefore, beforeNodeId) {
            const cur = loadSettings();
            let beforeId = '';
            if (liveLayoutOrBefore && Array.isArray(liveLayoutOrBefore.nodes)) {
                cur.schulOrganigramm = normalizeSchulOrganigramm(liveLayoutOrBefore);
                beforeId = typeof beforeNodeId === 'string' ? beforeNodeId : '';
            } else if (typeof liveLayoutOrBefore === 'string') {
                beforeId = liveLayoutOrBefore;
            }
            const result = moveSchulOrgRole(cur, roleNodeId, newParentId, beforeId);
            if (!result.ok || !result.settings) return result;
            saveSettings(result.settings);
            notify();
            return result;
        },
        moveAreaToParent: function (areaId, newParentId, beforeId) {
            const cur = loadSettings();
            const layout = normalizeSchulOrganigramm(cur.schulOrganigramm);
            const node = layout.nodes.find(function (n) {
                return n.id === areaId && isAreaNode(n);
            });
            if (!node) {
                return { ok: false, message: 'Bereich nicht gefunden.' };
            }
            const parentId =
                newParentId == null || newParentId === '' ? null : String(newParentId);
            if (parentId === String(areaId)) {
                return { ok: false, message: 'Bereich kann nicht in sich selbst abgelegt werden.' };
            }
            const moved = moveSchulOrganigrammArea(layout, areaId, parentId, beforeId || '');
            const after = moved.nodes.find(function (n) {
                return n.id === areaId;
            });
            const afterParent =
                after && after.parentId != null && after.parentId !== ''
                    ? String(after.parentId)
                    : '';
            const wantParent = parentId == null ? '' : String(parentId);
            if (!after || afterParent !== wantParent) {
                return { ok: false, message: 'Bereich kann hier nicht abgelegt werden.' };
            }
            cur.schulOrganigramm = moved;
            const applied = applySchulOrganigrammToStammdaten(cur, moved);
            saveSettings(applied.settings);
            notify();
            return {
                ok: true,
                settings: applied.settings,
                layout: applied.layout,
                message: 'Bereich verschoben.'
            };
        },
        addPersonToArea: function (areaNode, person) {
            const cur = loadSettings();
            const res = addSchulOrgPersonToArea(cur, areaNode, person);
            if (res.ok && res.settings) {
                saveSettings(res.settings);
                notify();
            }
            return res;
        },
        removePersonFromArea: function (areaNode, email) {
            const cur = loadSettings();
            const res = removeSchulOrgPersonFromArea(cur, areaNode, email);
            if (res.ok && res.settings) {
                saveSettings(res.settings);
                notify();
            }
            return res;
        },
        listPeopleForNode: function (node) {
            return listPeopleForSchulOrgNode(loadSettings(), node);
        },
        findAdminRole: findAdminRole,
        peopleForAdminRole: peopleForAdminRole,
        findArge: findArge,
        findTeacherByEmail: findTeacherByEmail,
        findClass: findClass,
        findStudentCouncil: findStudentCouncil,
        findSga: findSga
    };
}

if (typeof window !== 'undefined') {
    window.ms365SchulOrganigrammEdit = {
        createSchulOrganigrammEditor: createSchulOrganigrammEditor
    };
}

export default {
    createSchulOrganigrammEditor: createSchulOrganigrammEditor,
    moveSchulOrgRole: moveSchulOrgRole,
    findAdminRole: findAdminRole,
    peopleForAdminRole: peopleForAdminRole,
    studentsForClass: studentsForClass,
    guardiansForClass: guardiansForClass,
    teachersForClassChat: teachersForClassChat,
    teachersForArge: teachersForArge,
    applyArgeHeadFromGraphOwners: applyArgeHeadFromGraphOwners,
    applyClassHeadFromGraphOwners: applyClassHeadFromGraphOwners,
    normalizeSchulOrgAreaPersonRole: normalizeSchulOrgAreaPersonRole,
    ensureClassHeadsNamedFromTeachers: ensureClassHeadsNamedFromTeachers,
    healClassHeadsFromAreaPeople: healClassHeadsFromAreaPeople,
    listPeopleForSchulOrgNode: listPeopleForSchulOrgNode,
    listM365PeopleForArea: listM365PeopleForArea,
    listMailboxDelegatePeopleForArea: listMailboxDelegatePeopleForArea,
    previewPeopleForSchulOrgNode: previewPeopleForSchulOrgNode,
    moveSchulOrgPerson: moveSchulOrgPerson,
    addSchulOrgPersonToArea: addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea: removeSchulOrgPersonFromArea,
    SCHUL_ORG_PEOPLE_PREVIEW_LIMIT: SCHUL_ORG_PEOPLE_PREVIEW_LIMIT,
    findArge: findArge,
    findTeacherByEmail: findTeacherByEmail,
    findClass: findClass,
    findStudentCouncil: findStudentCouncil,
    findSga: findSga,
    sgaScopeLabel: sgaScopeLabel
};
