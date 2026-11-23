/**
 * Role editor UI for Schul-Organigramm.
 */
import {
    listSchulOrgRoleParentAreas,
    normalizeSchulOrganigramm,
    NODE,
    applySchulOrganigrammToStammdaten,
    moveSchulOrganigrammNode,
    isRoleCapableSection,
    canPlaceSchulOrgNodeInSection,
    childrenOf
} from './schul-organigramm-logic.js';
import { findAdminRole, peopleForAdminRole, roleCodeFromName, syncAdministrationFromAdmin } from './schul-organigramm-edit-people.js';
import { syncSeatMode, patchSchulOrganigrammRefLabel, removeSchulOrganigrammByRef, moveSchulOrgRole } from './schul-organigramm-edit-ops.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import { normalizeAdminAudienceMemberships } from '../../shared/administration-audience-groups.js';
import { createRoleCreateUi } from './schul-organigramm-edit-role-create-ui.js';

/**
 * @param {Record<string, any>} ctx
 */
export function createRoleUi(ctx) {
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

    function openRoleEditor(refKey) {
        closeOverlay();
        const settings = loadSettings();
        let role = findAdminRole(settings, refKey);
        if (!role) {
            toast('Rolle nicht in den Stammdaten gefunden.', 'warning');
            return;
        }
        // Arbeitskopie
        role = Object.assign({}, role);
        let people = peopleForAdminRole(settings, role).map(function (p) {
            return Object.assign({}, p);
        });
        const layout0 = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const roleNode0 = layout0.nodes.find(function (n) {
            return (
                (n.type === NODE.ROLE || n.type === 'role') &&
                (normStr(n.refKey).toUpperCase() === normStr(role.code).toUpperCase() ||
                    n.id === 'role_' + normStr(role.code))
            );
        });
        const parentAreas = listSchulOrgRoleParentAreas(layout0);
        const currentParentId = roleNode0 ? normStr(roleNode0.parentId) : '';

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-person-vcard',
            title: 'Rolle bearbeiten',
            sub: 'Änderungen schreiben in die Stammdaten (schulweit).',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Stammdaten</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgRoleName">Bezeichnung</label>' +
                '<input type="text" id="schulOrgRoleName" maxlength="80" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgRoleCode">Kürzel</label>' +
                '<input type="text" id="schulOrgRoleCode" maxlength="24" autocomplete="off" spellcheck="false"></div>' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgRoleArea">Bereich auf der Landkarte</label>' +
                '<select id="schulOrgRoleArea">' +
                (parentAreas.length
                    ? parentAreas
                          .map(function (a) {
                              return (
                                  '<option value="' +
                                  escapeHtml(a.id) +
                                  '"' +
                                  (a.id === currentParentId ? ' selected' : '') +
                                  '>' +
                                  escapeHtml(a.label || a.id) +
                                  '</option>'
                              );
                          })
                          .join('')
                    : '<option value="">— kein Bereich —</option>') +
                '</select>' +
                '<p class="schul-org-edit-hint">Wohin die Funktion auf der Landkarte gehört (auch per Ziehen verschiebbar).</p></div>' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgRoleSeat">Besetzung (Hinweis)</label>' +
                '<select id="schulOrgRoleSeat">' +
                '<option value="single">Einzelrolle – typisch eine Person</option>' +
                '<option value="multi">Team-Rolle – mehrere Personen</option>' +
                '</select>' +
                '<p class="schul-org-edit-hint">Weitere Personen sind immer möglich; die Farbe richtet sich nach der Besetzung.</p></div>' +
                '</div></section>' +
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-panel-head">' +
                '<h4 class="schul-org-edit-panel-title">Personen</h4>' +
                '<p class="schul-org-edit-hint" id="schulOrgPeopleHint"></p></div>' +
                '<div class="schul-org-edit-table-wrap"><table class="schul-org-edit-table" aria-label="Personen der Rolle">' +
                '<thead><tr><th>Name</th><th>E-Mail</th><th></th></tr></thead>' +
                '<tbody id="schulOrgPeopleBody"></tbody></table></div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-success" data-schul-org-add-person><i class="bi bi-person-plus" aria-hidden="true"></i> Person</button>' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-add-m365><i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-role', 'Rolle löschen', 'data-schul-org-save-role')
        });

        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgRoleName'));
        const codeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgRoleCode'));
        const areaEl = /** @type {HTMLSelectElement|null} */ (overlay.querySelector('#schulOrgRoleArea'));
        const seatEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgRoleSeat'));
        const tbody = overlay.querySelector('#schulOrgPeopleBody');
        const hintEl = overlay.querySelector('#schulOrgPeopleHint');
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        nameEl.value = role.name || '';
        codeEl.value = role.code || '';
        seatEl.value = role.seatMode === 'multi' ? 'multi' : 'single';
        if (areaEl && currentParentId) areaEl.value = currentParentId;

        function updateHint() {
            const n = people.filter(function (p) {
                return !isBlankPerson(p);
            }).length;
            if (titleEl) titleEl.textContent = nameEl.value.trim() || role.name || 'Rolle';
            if (!hintEl) return;
            if (n === 0) {
                hintEl.textContent = 'Noch niemand – Person hinzufügen oder aus Microsoft 365 übernehmen.';
            } else if (n === 1) {
                hintEl.textContent =
                    'Aktuell eine Person (Einzelbesetzung). Weitere jederzeit möglich → Team-Rolle.';
            } else {
                hintEl.textContent = 'Team-Rolle mit ' + n + ' Personen.';
            }
            if (n >= 2) seatEl.value = 'multi';
        }

        function renderPeople() {
            if (!tbody) return;
            tbody.replaceChildren();
            if (!people.length) {
                const tr = document.createElement('tr');
                const td = document.createElement('td');
                td.colSpan = 3;
                td.className = 'muted';
                td.textContent = 'Keine Personen.';
                tr.appendChild(td);
                tbody.appendChild(tr);
                updateHint();
                return;
            }
            people.forEach(function (person, idx) {
                const tr = document.createElement('tr');
                const tdN = document.createElement('td');
                const inpN = document.createElement('input');
                inpN.type = 'text';
                inpN.value = person.name || '';
                inpN.setAttribute('aria-label', 'Name');
                inpN.addEventListener('input', function () {
                    people[idx].name = inpN.value;
                    updateHint();
                });
                tdN.appendChild(inpN);
                const tdE = document.createElement('td');
                const inpE = document.createElement('input');
                inpE.type = 'email';
                inpE.value = person.email || '';
                inpE.setAttribute('aria-label', 'E-Mail');
                inpE.addEventListener('input', function () {
                    people[idx].email = inpE.value;
                });
                tdE.appendChild(inpE);
                const tdA = document.createElement('td');
                tdA.className = 'schul-org-edit-action';
                const btnDel = document.createElement('button');
                btnDel.type = 'button';
                btnDel.className = 'btn btn-sm btn-danger';
                btnDel.title = 'Person entfernen';
                btnDel.innerHTML = '<i class="bi bi-x-lg" aria-hidden="true"></i>';
                btnDel.addEventListener('click', function () {
                    people.splice(idx, 1);
                    renderPeople();
                });
                tdA.appendChild(btnDel);
                tr.append(tdN, tdE, tdA);
                tbody.appendChild(tr);
            });
            updateHint();
        }

        function persistRole(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextCode = normStr(codeEl.value).toUpperCase() || roleCodeFromName(nextName);
            if (!nextName) {
                toast('Bitte eine Bezeichnung eingeben.', 'warning');
                return;
            }
            const cur = loadSettings();
            let roles = Array.isArray(cur.adminRoles) ? cur.adminRoles.slice() : [];
            let admin = Array.isArray(cur.admin) ? cur.admin.slice() : [];
            const oldName = role.name;
            const oldCode = role.code;

            // Rename people role labels if name changed
            if (nextName !== oldName && typeof window.ms365TenantSettingsRenameAdminRole === 'function') {
                const renamed = window.ms365TenantSettingsRenameAdminRole(roles, admin, oldName, nextName);
                roles = renamed.roles;
                admin = renamed.admin;
            }

            let found =
                roles.find(function (r) {
                    return (
                        normStr(r.name).toLowerCase() === nextName.toLowerCase() ||
                        normStr(r.code).toUpperCase() === normStr(oldCode).toUpperCase()
                    );
                }) || null;
            if (!found) {
                found = {
                    code: nextCode,
                    name: nextName,
                    seatMode: seatEl.value === 'multi' ? 'multi' : 'single',
                    tier: role.tier || 'verwaltung'
                };
                roles.push(found);
            } else {
                found.name = nextName;
                const clash = roles.some(function (r) {
                    return r !== found && normStr(r.code).toUpperCase() === nextCode;
                });
                if (!clash) found.code = nextCode;
                found.seatMode = seatEl.value === 'multi' ? 'multi' : 'single';
                if (role.m365ResourceType) found.m365ResourceType = role.m365ResourceType;
                if (role.m365ResourceId) found.m365ResourceId = role.m365ResourceId;
                if (role.m365ResourceEmail) found.m365ResourceEmail = role.m365ResourceEmail;
                if (role.m365ResourceLabel) found.m365ResourceLabel = role.m365ResourceLabel;
            }
            syncSeatMode(found, people);

            // Replace people for this role
            admin = admin.filter(function (p) {
                return !personMatchesRole(p, { name: oldName, code: oldCode }) && !personMatchesRole(p, found);
            });
            people.forEach(function (p) {
                admin.push({
                    role: found.name,
                    name: normStr(p.name),
                    email: normEmail(p.email),
                    defaultKey: found.name
                });
            });

            cur.adminRoles = roles;
            cur.admin = admin;
            if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                cur.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(roles, admin);
            }
            patchSchulOrganigrammRefLabel(cur, 'role', oldCode || refKey, found.name, found.code);
            const wantParent = areaEl ? normStr(areaEl.value) : '';
            if (wantParent && cur.schulOrganigramm) {
                const layout = normalizeSchulOrganigramm(cur.schulOrganigramm);
                const roleNode = layout.nodes.find(function (n) {
                    return (
                        (n.type === NODE.ROLE || n.type === 'role') &&
                        (normStr(n.refKey).toUpperCase() === normStr(found.code).toUpperCase() ||
                            n.id === 'role_' + normStr(found.code) ||
                            n.id === 'role_' + normStr(oldCode))
                    );
                });
                if (roleNode && normStr(roleNode.parentId) !== wantParent) {
                    const movedRole = moveSchulOrgRole(cur, roleNode.id, wantParent);
                    if (movedRole.ok && movedRole.settings) {
                        Object.assign(cur, movedRole.settings);
                        rememberSchulOrgRoleParent(wantParent);
                    } else if (movedRole.message) {
                        toast(movedRole.message, 'warning');
                    }
                }
            }
            saveSettings(cur);
            role = Object.assign({}, found);
            toast('Rolle gespeichert.', 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                people = peopleForAdminRole(loadSettings(), found).map(function (p) {
                    return Object.assign({}, p);
                });
                nameEl.value = found.name || '';
                codeEl.value = found.code || '';
                seatEl.value = found.seatMode === 'multi' ? 'multi' : 'single';
                renderPeople();
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-role]').addEventListener('click', function () {
            persistRole(true);
        });
        overlay.querySelector('[data-schul-org-add-person]').addEventListener('click', function () {
            people.push({ role: role.name, name: '', email: '', defaultKey: role.name });
            syncSeatMode(role, people);
            if (people.filter(function (p) {
                return !isBlankPerson(p);
            }).length +
                people.filter(isBlankPerson).length >=
                2) {
                seatEl.value = 'multi';
            }
            renderPeople();
            const last = tbody && tbody.querySelector('tr:last-child input');
            if (last) last.focus();
        });
        overlay.querySelector('[data-schul-org-add-m365]').addEventListener('click', function () {
            const existing = new Set(
                people
                    .map(function (p) {
                        return normEmail(p.email);
                    })
                    .filter(function (e) {
                        return e.indexOf('@') !== -1;
                    })
            );
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                multi: true,
                hint:
                    'Personen für „' +
                    (nameEl.value || role.name) +
                    '“ suchen, anhaken und übernehmen.',
                excludeEmails: existing,
                onApply: function (picked) {
                    if (!picked || !picked.length) return;
                    let added = 0;
                    picked.forEach(function (u) {
                        const mail = normEmail(u.mail || u.userPrincipalName || u.email);
                        const name = normStr(u.displayName || u.name);
                        if (!mail || mail.indexOf('@') === -1 || existing.has(mail)) return;
                        const draft = people.find(isBlankPerson);
                        if (draft) {
                            draft.name = name;
                            draft.email = mail;
                        } else {
                            people.push({
                                role: role.name,
                                name: name,
                                email: mail,
                                defaultKey: role.name
                            });
                        }
                        existing.add(mail);
                        added += 1;
                    });
                    if (!added) {
                        toast('Keine neue Person übernommen.', 'info');
                        return;
                    }
                    seatEl.value =
                        people.filter(function (p) {
                            return !isBlankPerson(p);
                        }).length >= 2
                            ? 'multi'
                            : seatEl.value;
                    renderPeople();
                    toast(added + ' Person(en) aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-role]').addEventListener('click', function () {
            void confirmDlg('Rolle „' + (role.name || role.code) + '“ wirklich löschen? Personen müssen vorher entfernt sein.', {
                title: 'Rolle löschen',
                okText: 'Löschen',
                cancelText: 'Abbrechen'
            }).then(function (ok) {
                if (!ok) return;
                const filled = people.filter(function (p) {
                    return !isBlankPerson(p);
                });
                if (filled.length) {
                    toast('Zuerst alle Personen entfernen.', 'warning');
                    return;
                }
                const cur = loadSettings();
                const code = normStr(role.code).toUpperCase();
                cur.adminRoles = (cur.adminRoles || []).filter(function (r) {
                    return normStr(r.code).toUpperCase() !== code;
                });
                cur.admin = (cur.admin || []).filter(function (p) {
                    return !personMatchesRole(p, role);
                });
                if (typeof window.ms365TenantSettingsAdminRolesAndAdminToGroups === 'function') {
                    cur.administration = window.ms365TenantSettingsAdminRolesAndAdminToGroups(
                        cur.adminRoles,
                        cur.admin
                    );
                }
                if (cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)) {
                    cur.schulOrganigramm = {
                        version: 1,
                        nodes: cur.schulOrganigramm.nodes.filter(function (n) {
                            return !(
                                n.type === 'role' &&
                                normStr(n.refKey).toUpperCase() === code
                            );
                        })
                    };
                }
                saveSettings(cur);
                toast('Rolle gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        renderPeople();
        nameEl.focus();
    }

    const createUi = createRoleCreateUi(ctx, { openRoleEditor: openRoleEditor });
    const createRoleUnderSection = createUi.createRoleUnderSection;

    return { openRoleEditor, createRoleUnderSection };
}
