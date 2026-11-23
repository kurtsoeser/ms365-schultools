/**
 * Area / audience remove helpers for Schul-Organigramm.
 */
import {
    removeCustomVerwaltungAudienceGroup,
    normalizeVerwaltungAudienceGroups
} from '../../shared/administration-audience-groups.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import { patchAudienceGroupGraphId } from '../../shared/tenant-admin-audience-m365.js';
import {
    audienceIdForArea,
    canDeleteSchulOrganigrammArea,
    isOptionalDeletableBuiltinArea,
    withRemovedSchulOrganigrammArea,
    isAreaNode,
    normalizeSchulOrganigramm,
    applySchulOrganigrammToStammdaten,
    childrenOf,
    NODE,
    audienceSectionId,
    schulOrgEntityKindForNode,
    removeClassEntityFromSettings,
    removeArgeEntityFromSettings
} from './schul-organigramm-logic.js';

/**
 * @param {Record<string, any>} ctx
 * @param {{ persistOrganigrammLayout?: function }} [extra]
 */
export function createAreaRemoveUi(ctx, extra) {
    const loadSettings = ctx.loadSettings;
    const saveSettings = ctx.saveSettings;
    const toast = ctx.toast;
    const confirmDlg = ctx.confirmDlg;
    const notify = ctx.notify;
    const closeOverlay = ctx.closeOverlay;
    const onEsc = ctx.onEsc;
    const setActiveOverlay = ctx.setActiveOverlay;
    const escapeHtml = ctx.escapeHtml;
    const normStr = ctx.normStr;
    const normEmail = ctx.normEmail;
    const normCode = ctx.normCode;
    const editDialogHtml = ctx.editDialogHtml;
    const editFootHtml = ctx.editFootHtml;
    const areaM365ObjectsHtml = ctx.areaM365ObjectsHtml;
    const klassenchatsToolHref = ctx.klassenchatsToolHref;
    const toggleInlineEntraUserPicker = ctx.toggleInlineEntraUserPicker;
    const toggleInlineEntraGroupPicker = ctx.toggleInlineEntraGroupPicker;
    const rememberSchulOrgRoleParent = ctx.rememberSchulOrgRoleParent;
    const getLastSchulOrgRoleParent = ctx.getLastSchulOrgRoleParent;
    const pickEntraUser = ctx.pickEntraUser;
    const suggestTeacherCode = ctx.suggestTeacherCode;
    const ensureTeacherInList = ctx.ensureTeacherInList;
    const stripEphemeralSchoolSettings = ctx.stripEphemeralSchoolSettings;
    const plainJsonClone = ctx.plainJsonClone;
    const settingsForPersist = ctx.settingsForPersist;
    const suggestLocalPersonEmail = ctx.suggestLocalPersonEmail;
    const isBlankPerson = ctx.isBlankPerson;
    const personMatchesRole = ctx.personMatchesRole;
    const persistOrganigrammLayout =
        (extra && extra.persistOrganigrammLayout) || ctx.persistOrganigrammLayout;

    async function removeAudienceSection(audienceGroupId, opts) {
        const id = normStr(audienceGroupId);
        if (!id || id.indexOf('va-') !== 0) {
            toast('Nur selbst angelegte Bereiche können entfernt werden.', 'warning');
            return false;
        }
        if (!(opts && opts.skipConfirm)) {
            let ok = true;
            if (typeof window.ms365AppDialogConfirm === 'function') {
                ok = await window.ms365AppDialogConfirm(
                    'Bereich und Verwaltungsgruppe wirklich entfernen? Rollen bleiben erhalten und wandern nach „Verwaltung (Personal)“.',
                    { title: 'Bereich entfernen', okText: 'Entfernen', cancelText: 'Abbrechen' }
                );
            } else {
                ok = window.confirm('Bereich entfernen?');
            }
            if (!ok) return false;
        }

        const cur = loadSettings();
        const removed = removeCustomVerwaltungAudienceGroup(
            id,
            cur.verwaltungAudienceGroups,
            cur.adminAudienceMemberships
        );
        cur.verwaltungAudienceGroups = removed.groups;
        cur.adminAudienceMemberships = removed.memberships;
        const sid = audienceSectionId(id);
        if (cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)) {
            cur.schulOrganigramm = normalizeSchulOrganigramm({
                version: cur.schulOrganigramm.version,
                view: cur.schulOrganigramm.view,
                nodes: cur.schulOrganigramm.nodes
                    .filter(function (n) {
                        return n.id !== sid;
                    })
                    .map(function (n) {
                        if (n.parentId === sid) {
                            const hasVerwaltung = (cur.schulOrganigramm.nodes || []).some(
                                function (x) {
                                    return x && x.id === 'sec_verwaltung';
                                }
                            );
                            return Object.assign({}, n, {
                                parentId: hasVerwaltung ? 'sec_verwaltung' : null
                            });
                        }
                        return n;
                    })
            });
        }
        saveSettings(cur);
        toast('Bereich entfernt.', 'success');
        notify();
        return true;
    }

    /**
     * Bereich aus dem Organigramm entfernen (Custom-va-* inkl. Stammdaten, sonst Layout-Knoten).
     * @param {string|object} areaIdOrNode
     */
    async function removeArea(areaIdOrNode) {
        const settings = loadSettings();
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const areaId =
            typeof areaIdOrNode === 'string'
                ? normStr(areaIdOrNode)
                : normStr(areaIdOrNode && areaIdOrNode.id);
        const area = layout.nodes.find(function (n) {
            return n.id === areaId && isAreaNode(n);
        });
        if (!area) {
            toast('Bereich nicht gefunden.', 'warning');
            return false;
        }
        const entityKind = schulOrgEntityKindForNode(area);
        // Spiegel-/abgeleitete Knoten: Sync legt sie neu an – kein reines Layout-Löschen
        if (entityKind === 'elternClass' || entityKind === 'klassenchatClass' || entityKind === 'kursteam') {
            toast(
                'Dieser Bereich wird aus den Stammdaten abgeleitet und erscheint beim nächsten Sync wieder. Bitte die Klasse in den Stammdaten bzw. im Klassen-Editor ändern.',
                'warning'
            );
            return false;
        }
        // Klasse/ARGE: Entity-Löschen (Listen + Landkarte), kein Layout-only
        if (entityKind === 'class' || entityKind === 'arge') {
            const code = normStr(area.refKey).toUpperCase() || normStr(area.label).toUpperCase();
            if (entityKind === 'class') {
                const preview = removeClassEntityFromSettings(settings, code);
                const studentHint =
                    preview.studentCount > 0
                        ? ' ' +
                          preview.studentCount +
                          ' Schüler:innen behalten die Kennung „' +
                          code +
                          '“.'
                        : '';
                const okClass = await confirmDlg(
                    'Klasse „' +
                        (area.label || code) +
                        '“ aus Stammdaten und Landkarte löschen?' +
                        studentHint,
                    { title: 'Klasse löschen', okText: 'Löschen', cancelText: 'Abbrechen' }
                );
                if (!okClass) return false;
                const cur = loadSettings();
                const removed = removeClassEntityFromSettings(cur, code);
                Object.keys(removed.settings).forEach(function (k) {
                    cur[k] = removed.settings[k];
                });
                saveSettings(cur);
                toast('Klasse gelöscht.', 'success');
                notify();
                return true;
            }
            const okArge = await confirmDlg(
                'ARGE „' + (area.label || code) + '“ aus Stammdaten und Landkarte löschen?',
                { title: 'ARGE löschen', okText: 'Löschen', cancelText: 'Abbrechen' }
            );
            if (!okArge) return false;
            const curArge = loadSettings();
            const nextArge = removeArgeEntityFromSettings(curArge, code);
            Object.keys(nextArge).forEach(function (k) {
                curArge[k] = nextArge[k];
            });
            saveSettings(curArge);
            toast('ARGE gelöscht.', 'success');
            notify();
            return true;
        }
        if (!canDeleteSchulOrganigrammArea(area)) {
            toast(
                'Dieser Standard-Bereich (z. B. Lehrer, Schüler, ARGEs) kann nicht gelöscht werden.',
                'warning'
            );
            return false;
        }
        const audId = audienceIdForArea(area);
        const isCustomAud = audId && String(audId).indexOf('va-') === 0;
        const isOptionalBuiltin = isOptionalDeletableBuiltinArea(area);
        let ok = true;
        if (typeof window.ms365AppDialogConfirm === 'function') {
            ok = await window.ms365AppDialogConfirm(
                isCustomAud
                    ? 'Bereich und Verwaltungsgruppe wirklich entfernen? Rollen bleiben erhalten und wandern nach „Verwaltung (Personal)“ (falls vorhanden) bzw. eine Ebene höher.'
                    : isOptionalBuiltin
                      ? '„' +
                        (area.label || 'Bereich') +
                        '“ von der Landkarte entfernen? Untergeordnete Einträge wandern eine Ebene höher. Der Bereich wird nicht automatisch wieder angelegt.'
                      : 'Bereich „' +
                        (area.label || 'Bereich') +
                        '“ wirklich entfernen? Untergeordnete Einträge wandern eine Ebene höher.',
                { title: 'Bereich entfernen', okText: 'Entfernen', cancelText: 'Abbrechen' }
            );
        } else {
            ok = window.confirm('Bereich entfernen?');
        }
        if (!ok) return false;

        if (isCustomAud) {
            return removeAudienceSection(audId, { skipConfirm: true });
        }

        const cur = loadSettings();
        const fallbackParent =
            area.parentId == null || area.parentId === '' ? null : String(area.parentId);
        const nextNodes = (cur.schulOrganigramm.nodes || [])
            .filter(function (n) {
                return n.id !== area.id;
            })
            .map(function (n) {
                if (n.parentId === area.id) {
                    return Object.assign({}, n, { parentId: fallbackParent });
                }
                return n;
            });
        let nextLayout = normalizeSchulOrganigramm({
            version: cur.schulOrganigramm.version,
            view: cur.schulOrganigramm.view,
            nodes: nextNodes,
            removedAreaIds: cur.schulOrganigramm.removedAreaIds
        });
        if (isOptionalBuiltin) {
            nextLayout = withRemovedSchulOrganigrammArea(nextLayout, area.id);
        }
        cur.schulOrganigramm = nextLayout;
        saveSettings(cur);
        toast('Bereich entfernt.', 'success');
        notify();
        return true;
    }

    return { removeAudienceSection, removeArea };
}
