/**
 * Area create/remove helpers for Schul-Organigramm.
 */
import {
    addCustomVerwaltungAudienceGroup,
    removeCustomVerwaltungAudienceGroup,
    normalizeVerwaltungAudienceGroups
} from '../../shared/administration-audience-groups.js';
import { writeSchoolData, writeSchulOrganigramm } from '../../shared/stammdaten-write.js';
import {
    getAudienceGroupGraphId,
    patchAudienceGroupGraphId
} from '../../shared/tenant-admin-audience-m365.js';
import {
    audienceSectionId,
    audienceIdForArea,
    sammelgruppeKeyForArea,
    ensureSchulOrganigrammSgaStructure,
    canDeleteSchulOrganigrammArea,
    isOptionalDeletableBuiltinArea,
    withRemovedSchulOrganigrammArea,
    areaAcceptsDirectPeople,
    isRoleCapableSection,
    canPlaceSchulOrgNodeInSection,
    applySchulOrganigrammToStammdaten,
    isAreaNode,
    normalizeSchulOrganigramm,
    setSchulOrganigrammAreaM365,
    setSchulOrganigrammNodePosition,
    childrenOf,
    NODE,
    normalizeSchulOrgGroupKind,
    suggestSchulOrgMailNickname,
    normalizeSchulOrgM365,
    emptySchulOrgM365Bindings,
    jahrgangAreaId,
    schulstufeFromClassCode,
    isKlassenchatClassNode,
    isArgeAreaNode,
    isClassAreaNode
} from './schul-organigramm-logic.js';
import { findClass } from './schul-organigramm-edit-find.js';
import { resolveClassGraphGroupId } from '../../shared/membership-hygiene.js';

/**
 * @param {Record<string, any>} ctx
 */
import { createAreaRemoveUi } from './schul-organigramm-edit-area-remove-ui.js';

export function createAreaCreateUi(ctx) {
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

    async function createAudienceSection() {
        let name = '';
        if (typeof window.ms365AppDialogPrompt === 'function') {
            name = await window.ms365AppDialogPrompt(
                'Name des neuen Bereichs (wird auch unter Stammdaten → Verwaltung gespeichert)',
                '',
                {
                    title: 'Verwaltungsgruppe anlegen',
                    inputLabel: 'z. B. Erweiterte Schulleitung, Schulwarte …'
                }
            );
        } else {
            name = window.prompt('Name der neuen Verwaltungsgruppe', '') || '';
        }
        const label = normStr(name);
        if (!label) return null;

        const cur = loadSettings();
        const groups = normalizeVerwaltungAudienceGroups(cur.verwaltungAudienceGroups);
        const dup = groups.some(function (g) {
            return normStr(g.label).toLowerCase() === label.toLowerCase();
        });
        if (dup) {
            toast('Diesen Bereich gibt es bereits.', 'warning');
            return null;
        }
        const added = addCustomVerwaltungAudienceGroup(label, groups);
        if (!added.id) {
            toast('Gruppe konnte nicht angelegt werden.', 'warning');
            return null;
        }
        cur.verwaltungAudienceGroups = added.groups;
        if (!Array.isArray(cur.adminAudienceMemberships)) cur.adminAudienceMemberships = [];

        const sid = audienceSectionId(added.id);
        const layout = normalizeSchulOrganigramm(
            cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)
                ? cur.schulOrganigramm
                : { version: 2, nodes: [] }
        );
        const maxOrder = layout.nodes.reduce(function (m, n) {
            return Math.max(m, typeof n.order === 'number' ? n.order : 0);
        }, 0);
        // Nach Verwaltung (Personal), vor ARGEs einfügen
        const staff = layout.nodes.find(function (n) {
            return n.id === 'sec_verwaltung';
        });
        const insertOrder =
            staff && typeof staff.order === 'number' ? staff.order + 0.5 : maxOrder + 1;
        layout.nodes.push({
            id: sid,
            type: 'area',
            label: label,
            parentId: null,
            order: insertOrder,
            tone: 'staff',
            refKind: 'audience',
            refKey: added.id
        });
        // Orders neu verdichten (nur Top-Level Areas)
        const roots = layout.nodes
            .filter(function (n) {
                return isAreaNode(n) && (n.parentId == null || n.parentId === '');
            })
            .sort(function (a, b) {
                return (a.order || 0) - (b.order || 0);
            });
        roots.forEach(function (n, i) {
            n.order = i;
        });
        cur.schulOrganigramm = normalizeSchulOrganigramm(layout);
        saveSettings(cur);
        toast('Bereich „' + label + '“ angelegt (auch in Stammdaten → Verwaltung).', 'success');
        notify();
        return { id: added.id, sectionId: sid, label: label };
    }

    /**
     * Landkarte direkt in v2 speichern (ohne vollen Stammdaten-Save – der verschluckt Fehler).
     * @param {object} layout
     * @param {string} reason
     */
    function persistOrganigrammLayout(layout, reason) {
        const clean = plainJsonClone(
            normalizeSchulOrganigramm(layout),
            normalizeSchulOrganigramm(layout)
        );
        if (typeof writeSchulOrganigramm === 'function') {
            writeSchulOrganigramm(clean, {
                reason: reason || 'schul-org-layout',
                materializeLists: false
            });
            return clean;
        }
        if (
            typeof window !== 'undefined' &&
            window.ms365AppDataV2 &&
            typeof window.ms365AppDataV2.getContainer === 'function' &&
            typeof window.ms365AppDataV2.setContainer === 'function'
        ) {
            const c = window.ms365AppDataV2.getContainer();
            if (!c || !c.core) throw new Error('App-Daten fehlen.');
            c.core.schulOrganigramm = clean;
            window.ms365AppDataV2.setContainer(c);
            return clean;
        }
        const cur = stripEphemeralSchoolSettings(loadSettings());
        cur.schulOrganigramm = clean;
        try {
            window.__ms365SchulOrgSkipSyncOnce = true;
        } catch {
            /* ignore */
        }
        saveSettings(cur);
        return clean;
    }

    /**
     * Namensdialog für neue Bereiche – eigener Overlay (z-index hoch),
     * unabhängig von App-Dialog hinter Landkarten-Diff.
     * @param {{ title: string, message: string, inputLabel: string }} opts
     * @returns {Promise<string|null>}
     */
    function promptFreeAreaName(opts) {
        const title = (opts && opts.title) || 'Bereich anlegen';
        const message = (opts && opts.message) || 'Name des neuen Bereichs';
        const inputLabel =
            (opts && opts.inputLabel) || 'Bezeichnung (z. B. IT-Team, Sekretariat, …)';
        return new Promise(function (resolve) {
            const existing = document.getElementById('schulOrgCreateAreaOverlay');
            if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

            const overlay = document.createElement('div');
            overlay.id = 'schulOrgCreateAreaOverlay';
            overlay.className = 'modal-overlay open schul-org-edit-overlay schul-org-create-area-overlay';
            overlay.style.zIndex = '14000';
            overlay.innerHTML =
                '<div class="modal-box schul-org-edit-box" role="dialog" aria-modal="true" aria-labelledby="schulOrgCreateAreaTitle" tabindex="-1">' +
                '<header class="schul-org-edit-head">' +
                '<div class="schul-org-edit-head-main">' +
                '<span class="schul-org-edit-icon" aria-hidden="true"><i class="bi bi-plus-square"></i></span>' +
                '<div class="schul-org-edit-titles">' +
                '<h3 id="schulOrgCreateAreaTitle" class="schul-org-edit-title">' +
                escapeHtml(title) +
                '</h3>' +
                '<p class="schul-org-edit-sub">' +
                escapeHtml(message) +
                '</p></div></div></header>' +
                '<div class="schul-org-edit-body">' +
                '<div class="field"><label for="schulOrgCreateAreaName">' +
                escapeHtml(inputLabel) +
                '</label>' +
                '<input type="text" id="schulOrgCreateAreaName" maxlength="80" autocomplete="off" spellcheck="false"></div>' +
                '</div>' +
                '<footer class="schul-org-edit-foot">' +
                '<button type="button" class="btn" data-cancel>Abbrechen</button>' +
                '<button type="button" class="btn btn-success" data-ok>Anlegen</button>' +
                '</footer></div>';

            function finish(value) {
                document.removeEventListener('keydown', onKey, true);
                if (overlay.parentNode) overlay.parentNode.removeChild(overlay);
                resolve(value);
            }
            function onKey(ev) {
                if (ev.key === 'Escape') {
                    ev.preventDefault();
                    finish(null);
                } else if (ev.key === 'Enter') {
                    const t = /** @type {HTMLElement} */ (ev.target);
                    if (t && t.id === 'schulOrgCreateAreaName') {
                        ev.preventDefault();
                        const inp = /** @type {HTMLInputElement|null} */ (
                            overlay.querySelector('#schulOrgCreateAreaName')
                        );
                        finish(inp ? inp.value : '');
                    }
                }
            }

            overlay.addEventListener('click', function (ev) {
                if (ev.target === overlay) finish(null);
            });
            const cancelBtn = overlay.querySelector('[data-cancel]');
            const okBtn = overlay.querySelector('[data-ok]');
            if (cancelBtn) {
                cancelBtn.addEventListener('click', function () {
                    finish(null);
                });
            }
            if (okBtn) {
                okBtn.addEventListener('click', function () {
                    const inp = /** @type {HTMLInputElement|null} */ (
                        overlay.querySelector('#schulOrgCreateAreaName')
                    );
                    finish(inp ? inp.value : '');
                });
            }
            document.body.appendChild(overlay);
            document.addEventListener('keydown', onKey, true);
            requestAnimationFrame(function () {
                const inp = /** @type {HTMLInputElement|null} */ (
                    overlay.querySelector('#schulOrgCreateAreaName')
                );
                if (inp) inp.focus();
            });
        });
    }

    /**
     * Kurzer Toast ohne blockierenden Modal-Dialog.
     * @param {string} msg
     * @param {string} [kind]
     */
    function softToast(msg, kind) {
        if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg, {
                kind: kind || 'info',
                title: 'Schul-Organigramm'
            });
            return;
        }
        if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg, { kind: kind || 'info', title: 'Schul-Organigramm' });
            return;
        }
        toast(msg, kind);
    }

    /**
     * Freien Bereich auf der Landkarte anlegen (Top-Level oder Unterbereich).
     * Keine Zwangs-Verwaltungsgruppe – nur Layout-Knoten (refKind none).
     * @param {{ parentId?: string|null, x?: number, y?: number, label?: string }} [opts]
     */
    async function createAreaAt(opts) {
        const o = opts && typeof opts === 'object' ? opts : {};
        const parentId =
            o.parentId == null || o.parentId === '' ? null : String(o.parentId);

        try {
            let name = normStr(o.label);
            if (!name) {
                const promptTitle = parentId ? 'Unterbereich anlegen' : 'Bereich anlegen';
                const promptMsg = parentId
                    ? 'Name des neuen Unterbereichs'
                    : 'Name des neuen Bereichs auf der Landkarte';
                name = await promptFreeAreaName({
                    title: promptTitle,
                    message: promptMsg,
                    inputLabel: 'Bezeichnung (z. B. IT-Team, Sekretariat, …)'
                });
                name = normStr(name);
            }
            if (!name) {
                softToast('Anlegen abgebrochen – kein Name eingegeben.', 'warning');
                return null;
            }

            const cur = loadSettings();
            const layout = normalizeSchulOrganigramm(cur.schulOrganigramm);
            let parent = null;
            if (parentId) {
                parent = layout.nodes.find(function (n) {
                    return n.id === parentId && isAreaNode(n);
                });
                if (!parent) {
                    softToast('Eltern-Bereich nicht gefunden.', 'warning');
                    return null;
                }
            }

            const id =
                'area_' +
                Math.random().toString(36).slice(2, 8) +
                Date.now().toString(36).slice(-4);

            if (parentId) {
                const siblings = childrenOf(layout, parentId).filter(isAreaNode);
                siblings.forEach(function (s) {
                    s.order = (typeof s.order === 'number' ? s.order : 0) + 1;
                });
            }

            const rootCount = layout.nodes.filter(function (n) {
                return isAreaNode(n) && (n.parentId == null || n.parentId === '');
            }).length;
            const slot = !parentId ? defaultCanvasSlotForNewArea(rootCount, o.x, o.y) : null;

            const node = {
                id: id,
                type: 'area',
                label: name,
                parentId: parentId,
                order: parentId
                    ? 0
                    : layout.nodes.reduce(function (m, n) {
                          return Math.max(m, typeof n.order === 'number' ? n.order : 0);
                      }, 0) + 1,
                tone: (parent && parent.tone) || 'custom',
                refKind: 'none',
                refKey: ''
            };
            if (slot) {
                node.x = slot.x;
                node.y = slot.y;
            }
            layout.nodes.push(node);

            const saved = persistOrganigrammLayout(layout, 'schul-org-create-area');

            // Sofort prüfen, ob der Knoten wirklich im Kanon liegt
            let verified = false;
            let verifiedLayout = saved;
            try {
                const api = window.ms365AppDataV2;
                const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
                if (c && c.core && c.core.schulOrganigramm) {
                    verifiedLayout = normalizeSchulOrganigramm(c.core.schulOrganigramm);
                    verified = verifiedLayout.nodes.some(function (n) {
                        return n && n.id === id;
                    });
                } else {
                    verified = !!(saved && Array.isArray(saved.nodes) && saved.nodes.some(function (n) {
                        return n && n.id === id;
                    }));
                }
            } catch {
                verified = !!(saved && Array.isArray(saved.nodes));
            }
            if (!verified) {
                throw new Error('Bereich wurde nicht gespeichert (Speicher prüfen).');
            }

            softToast(
                (parentId ? 'Unterbereich' : 'Bereich') + ' „' + name + '“ angelegt.',
                'success'
            );
            // Kein notify()/Voll-Remount hier – Aufrufer aktualisiert die Karte weich
            // (sonst kann flushLandkartePositions den neuen Knoten kurz „übersehen“).
            return {
                sectionId: id,
                label: name,
                parentId: parentId,
                layout: verifiedLayout,
                x: node.x,
                y: node.y
            };
        } catch (err) {
            console.error('[schul-org] createAreaAt', err);
            softToast(
                'Bereich anlegen fehlgeschlagen: ' +
                    (err && err.message ? err.message : String(err)),
                'error'
            );
            return null;
        }
    }

    /**
     * Canvas-Ort für neuen Top-Level-Bereich.
     * @param {number} rootIndex
     * @param {unknown} x
     * @param {unknown} y
     */
    function defaultCanvasSlotForNewArea(rootIndex, x, y) {
        const cx = Number(x);
        const cy = Number(y);
        if (Number.isFinite(cx) && Number.isFinite(cy)) {
            return { x: Math.round(cx), y: Math.round(cy) };
        }
        const col = rootIndex % 3;
        const row = Math.floor(rootIndex / 3);
        return { x: 48 + col * 920, y: 48 + row * 720 };
    }

    const removed = createAreaRemoveUi(ctx, {
        persistOrganigrammLayout: persistOrganigrammLayout
    });

    return {
        createAudienceSection,
        createAreaAt,
        softToast,
        promptFreeAreaName,
        persistOrganigrammLayout,
        defaultCanvasSlotForNewArea,
        removeAudienceSection: removed.removeAudienceSection,
        removeArea: removed.removeArea
    };
}
