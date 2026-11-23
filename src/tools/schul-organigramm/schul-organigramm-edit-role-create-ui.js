/**
 * Create-role dialog for Schul-Organigramm.
 */
import {
    listSchulOrgRoleParentAreas,
    normalizeSchulOrganigramm,
    NODE,
    isRoleCapableSection,
    canPlaceSchulOrgNodeInSection,
    childrenOf
} from './schul-organigramm-logic.js';
import { roleCodeFromName, syncAdministrationFromAdmin } from './schul-organigramm-edit-people.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import { readLastRoleParentId } from './schul-organigramm-edit-overlay.js';

/**
 * @param {Record<string, any>} ctx
 * @param {{ openRoleEditor: function(string): void }} api
 */
export function createRoleCreateUi(ctx, api) {
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
    const softToast = ctx.softToast;
    const promptFreeAreaName = ctx.promptFreeAreaName;
    const persistOrganigrammLayout = ctx.persistOrganigrammLayout;
    const defaultCanvasSlotForNewArea = ctx.defaultCanvasSlotForNewArea;
    const openRoleEditor = api.openRoleEditor;

    async function createRoleUnderSection(parentSectionId) {
        closeOverlay();
        const settings = loadSettings();
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const parentAreas = listSchulOrgRoleParentAreas(layout);
        if (!parentAreas.length) {
            toast(
                'Kein geeigneter Bereich für Funktionen – zuerst z. B. Verwaltung oder einen freien Bereich anlegen.',
                'warning'
            );
            return null;
        }

        const preferred = normStr(parentSectionId);
        const remembered = readLastRoleParentId();
        let preId = '';
        if (
            preferred &&
            parentAreas.some(function (a) {
                return a.id === preferred;
            })
        ) {
            preId = preferred;
        } else if (
            remembered &&
            parentAreas.some(function (a) {
                return a.id === remembered;
            })
        ) {
            preId = remembered;
        } else {
            preId = parentAreas[0].id;
        }

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-person-badge',
            title: 'Funktion anlegen',
            sub: 'Funktion auf der Landkarte platzieren – Bereich jederzeit änderbar.',
            body:
                '<div class="field"><label for="schulOrgNewRoleName">Bezeichnung</label>' +
                '<input type="text" id="schulOrgNewRoleName" maxlength="80" autocomplete="off" placeholder="z. B. Qualitätsmanagement"></div>' +
                '<div class="field"><label for="schulOrgNewRoleArea">Bereich auf der Landkarte</label>' +
                '<select id="schulOrgNewRoleArea">' +
                parentAreas
                    .map(function (a) {
                        return (
                            '<option value="' +
                            escapeHtml(a.id) +
                            '"' +
                            (a.id === preId ? ' selected' : '') +
                            '>' +
                            escapeHtml(a.label || a.id) +
                            '</option>'
                        );
                    })
                    .join('') +
                '</select>' +
                '<p class="muted schul-org-edit-hint">Die Funktion erscheint in diesem Bereich und lässt sich später per Ziehen oder im Editor einem anderen Bereich zuordnen.</p></div>' +
                '<div class="field"><label for="schulOrgNewRoleSeat">Besetzung</label>' +
                '<select id="schulOrgNewRoleSeat">' +
                '<option value="single">Einzelrolle – typisch eine Person</option>' +
                '<option value="multi">Team-Rolle – mehrere Personen</option>' +
                '</select></div>',
            foot:
                '<span class="schul-org-edit-foot-spacer"></span>' +
                '<button type="button" class="btn alt" data-schul-org-close>Abbrechen</button>' +
                '<button type="button" class="btn btn-success" data-schul-org-save-new-role>' +
                '<i class="bi bi-check-lg" aria-hidden="true"></i> Anlegen</button>'
        });
        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgNewRoleName'));
        const areaEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgNewRoleArea'));
        const seatEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgNewRoleSeat'));
        const saveBtn = /** @type {HTMLButtonElement|null} */ (
            overlay.querySelector('[data-schul-org-save-new-role]')
        );
        nameEl.focus();

        /**
         * Inline-Hinweis im Dialog (sichtbar auch wenn Toast hinter Overlay lag).
         * @param {string} msg
         * @param {'warning'|'error'|'success'} [kind]
         */
        function showFormFeedback(msg, kind) {
            let el = overlay.querySelector('[data-schul-org-new-role-feedback]');
            if (!el) {
                el = document.createElement('p');
                el.setAttribute('data-schul-org-new-role-feedback', '');
                el.setAttribute('role', 'status');
                el.style.margin = '0 0 10px';
                el.style.padding = '8px 10px';
                el.style.borderRadius = '8px';
                el.style.fontSize = '0.92rem';
                const body = overlay.querySelector('.schul-org-edit-body');
                if (body) body.insertBefore(el, body.firstChild);
                else overlay.appendChild(el);
            }
            const k = kind || 'warning';
            el.style.background =
                k === 'error'
                    ? 'color-mix(in srgb, #dc2626 14%, transparent)'
                    : k === 'success'
                      ? 'color-mix(in srgb, #16a34a 14%, transparent)'
                      : 'color-mix(in srgb, #f59e0b 16%, transparent)';
            el.textContent = msg;
            toast(msg, k);
        }

        /**
         * @returns {{ code: string, parentId: string }|null}
         */
        function commitNewRole() {
            try {
                const label = normStr(nameEl.value);
                if (!label) {
                    showFormFeedback('Bitte eine Bezeichnung eingeben.', 'warning');
                    nameEl.focus();
                    return null;
                }
                const cur = stripEphemeralSchoolSettings(
                    plainJsonClone(loadSettings(), loadSettings())
                );
                const exists = (cur.adminRoles || []).some(function (r) {
                    return normStr(r.name).toLowerCase() === label.toLowerCase();
                });
                if (exists) {
                    showFormFeedback('Diese Funktion gibt es bereits.', 'warning');
                    return null;
                }
                const used = new Set(
                    (cur.adminRoles || []).map(function (r) {
                        return normStr(r.code).toUpperCase();
                    })
                );
                let code = roleCodeFromName(label);
                let i = 2;
                while (used.has(code)) {
                    code = (roleCodeFromName(label) + String(i)).slice(0, 24);
                    i += 1;
                }
                let nextLayout = normalizeSchulOrganigramm(cur.schulOrganigramm);
                const resolvedParent = normStr(areaEl && areaEl.value);
                if (
                    !resolvedParent ||
                    !canPlaceSchulOrgNodeInSection(
                        { type: NODE.ROLE, refKind: 'adminRole' },
                        resolvedParent,
                        nextLayout
                    )
                ) {
                    showFormFeedback(
                        'Dieser Bereich nimmt keine Verwaltungsfunktionen auf. Bitte z. B. Verwaltung oder einen freien Bereich wählen.',
                        'warning'
                    );
                    return null;
                }

                const tier =
                    /direktion|direktor|schulleit/i.test(label) ||
                    resolvedParent === 'sec_schulleitung'
                        ? 'schulleitung'
                        : 'verwaltung';
                const role = {
                    code: code,
                    name: label,
                    tier: tier,
                    seatMode: seatEl && seatEl.value === 'multi' ? 'multi' : 'single'
                };
                cur.adminRoles = (cur.adminRoles || []).concat([role]);
                if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                    cur.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(
                        cur.adminRoles,
                        cur.admin || []
                    );
                } else {
                    // Fallback: leere Gruppe anhängen, damit normalizeSettings die Rolle nicht verwirft
                    const groups = Array.isArray(cur.administration) ? cur.administration.slice() : [];
                    groups.push({
                        code: role.code,
                        name: role.name,
                        people: [],
                        tier: role.tier,
                        seatMode: role.seatMode
                    });
                    cur.administration = groups;
                }
                const maxOrder = nextLayout.nodes.reduce(function (m, n) {
                    return Math.max(m, n.order || 0);
                }, 0);
                nextLayout.nodes.push({
                    id: 'role_' + code,
                    type: 'role',
                    label: label,
                    parentId: resolvedParent,
                    order: maxOrder + 1,
                    tone:
                        tier === 'schulleitung' || resolvedParent === 'sec_schulleitung'
                            ? 'lead'
                            : 'staff',
                    refKind: 'adminRole',
                    refKey: code
                });
                cur.schulOrganigramm = plainJsonClone(
                    normalizeSchulOrganigramm(nextLayout),
                    nextLayout
                );
                rememberSchulOrgRoleParent(resolvedParent);

                if (typeof writeSchoolData === 'function') {
                    writeSchoolData(cur, { reason: 'schul-org-create-role' });
                } else {
                    saveSettings(cur);
                }

                // Verifizieren: Rolle muss im Kanon liegen
                let verified = false;
                try {
                    const api = window.ms365AppDataV2;
                    const c =
                        api && typeof api.getContainer === 'function' ? api.getContainer() : null;
                    const roles =
                        c && c.core && Array.isArray(c.core.adminRoles) ? c.core.adminRoles : [];
                    verified = roles.some(function (r) {
                        return normStr(r && r.code).toUpperCase() === code.toUpperCase();
                    });
                } catch {
                    verified = false;
                }
                if (!verified) {
                    throw new Error('Funktion wurde nicht gespeichert (Speicher prüfen).');
                }

                const parentLabel =
                    (
                        nextLayout.nodes.find(function (n) {
                            return n.id === resolvedParent;
                        }) || {}
                    ).label || resolvedParent;
                toast('Funktion „' + label + '“ in „' + parentLabel + '“ angelegt.', 'success');
                closeOverlay();
                notify();
                try {
                    openRoleEditor(code);
                } catch (openErr) {
                    console.warn('[schul-org] openRoleEditor after create', openErr);
                }
                return { code: code, parentId: resolvedParent };
            } catch (err) {
                console.error('[schul-org] commitNewRole', err);
                showFormFeedback(
                    'Anlegen fehlgeschlagen: ' +
                        (err && err.message ? err.message : String(err)),
                    'error'
                );
                return null;
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        if (saveBtn) {
            saveBtn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                commitNewRole();
            });
        }
        nameEl.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                commitNewRole();
            }
        });
        return null;
    }

    return { createRoleUnderSection };
}
