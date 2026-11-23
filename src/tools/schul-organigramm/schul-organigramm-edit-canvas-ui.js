/**
 * Canvas person add/edit helpers for Schul-Organigramm.
 */
import {
    normalizeSchulOrganigramm,
    isAreaNode,
    areaAcceptsDirectPeople,
    isSchulOrgLocalPlaceholderEmail,
    isClassNode,
    isArgeNode,
    NODE,
    SCHUL_ORG_BOARD_REF,
    isSgaBoardArea,
    isSammelgruppeArea,
    isSgaScopeArea,
    sammelgruppeKeyForArea,
    upsertSchulOrganigrammAreaPerson,
    removeSchulOrganigrammAreaPerson
} from './schul-organigramm-logic.js';
import {
    addSchulOrgPersonToArea,
    removeSchulOrgPersonFromArea,
    moveSchulOrgPerson,
    findStudentIndex,
    findAdminIndex
} from './schul-organigramm-edit-ops.js';
import { syncAdministrationFromAdmin } from './schul-organigramm-edit-people.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';

/**
 * @param {Record<string, any>} ctx
 */
export function createCanvasUi(ctx) {
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
    const openClassEditor = ctx.openClassEditor;
    const openArgeEditor = ctx.openArgeEditor;
    const openRoleEditor = ctx.openRoleEditor;
    const openCouncilEditor = ctx.openCouncilEditor;
    const openSgaEditor = ctx.openSgaEditor;

    function openEditCanvasPerson(person, contextNode) {
        closeOverlay();
        const p = person && typeof person === 'object' ? person : {};
        const kind = normStr(p.kind).toLowerCase();
        const email0 = normEmail(p.email);
        const name0 = normStr(p.name);

        // Leitung/KV auch unbesetzt → Editor zum Zuordnen
        if (kind === 'head' && contextNode) {
            if (isClassNode(contextNode) && contextNode.refKey) {
                openClassEditor(contextNode.refKey);
                return;
            }
            if (isArgeNode(contextNode) && contextNode.refKey) {
                openArgeEditor(contextNode.refKey);
                return;
            }
        }

        if (!email0 && !name0) {
            toast('Person nicht gefunden.', 'warning');
            return;
        }
        if (kind === 'admin' && contextNode && contextNode.type === NODE.ROLE && contextNode.refKey) {
            openRoleEditor(contextNode.refKey);
            return;
        }
        if (kind === 'council') {
            openCouncilEditor(SCHUL_ORG_BOARD_REF);
            return;
        }
        if (kind === 'sga') {
            openSgaEditor(SCHUL_ORG_BOARD_REF);
            return;
        }

        const settings = loadSettings();
        let klasse0 = '';
        if (kind === 'student') {
            const hit = (Array.isArray(settings.students) ? settings.students : []).find(function (s) {
                return normEmail(s && s.email) === email0 || (!email0 && normStr(s && s.name) === name0);
            });
            if (hit) klasse0 = normStr(hit.klasse);
        }

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-person-gear',
            title: 'Person bearbeiten',
            sub:
                kind === 'student'
                    ? 'Schüler:in in den Stammdaten'
                    : kind === 'teacher'
                      ? 'Lehrkraft in den Stammdaten'
                      : 'Person in den Stammdaten',
            body:
                '<div class="schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgEditPersonName">Name</label>' +
                '<input type="text" id="schulOrgEditPersonName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgEditPersonEmail">E-Mail</label>' +
                '<input type="email" id="schulOrgEditPersonEmail" maxlength="160" autocomplete="email"></div>' +
                (kind === 'student'
                    ? '<div class="field schul-org-edit-field-full"><label for="schulOrgEditPersonKlasse">Klasse</label>' +
                      '<input type="text" id="schulOrgEditPersonKlasse" maxlength="24" autocomplete="off"></div>'
                    : '') +
                '</div>',
            foot:
                '<span class="schul-org-edit-foot-spacer"></span>' +
                '<button type="button" class="btn alt" data-schul-org-close>Abbrechen</button>' +
                '<button type="button" class="btn btn-success" data-schul-org-save>Speichern</button>'
        });
        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgEditPersonName'));
        const emailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgEditPersonEmail'));
        const klasseEl = /** @type {HTMLInputElement|null} */ (
            overlay.querySelector('#schulOrgEditPersonKlasse')
        );
        nameEl.value = name0;
        emailEl.value = email0;
        if (klasseEl) klasseEl.value = klasse0;
        nameEl.focus();

        overlay.querySelector('[data-schul-org-close]').addEventListener('click', closeOverlay);
        overlay.querySelector('[data-schul-org-save]').addEventListener('click', function () {
            const nextName = normStr(nameEl.value);
            const nextEmail = normEmail(emailEl.value);
            if (!nextName && !nextEmail) {
                toast('Name oder E-Mail angeben.', 'warning');
                return;
            }
            const cur = loadSettings();
            let ok = false;
            if (kind === 'student') {
                if (!Array.isArray(cur.students)) cur.students = [];
                const idx = findStudentIndex(cur.students, { email: email0, name: name0 }, klasse0);
                if (idx >= 0) {
                    cur.students[idx] = Object.assign({}, cur.students[idx], {
                        name: nextName || cur.students[idx].name,
                        email: nextEmail || cur.students[idx].email,
                        klasse: klasseEl ? normStr(klasseEl.value) || cur.students[idx].klasse : cur.students[idx].klasse
                    });
                    ok = true;
                }
            } else if (kind === 'teacher') {
                if (!Array.isArray(cur.teachers)) cur.teachers = [];
                const idx = cur.teachers.findIndex(function (t) {
                    return normEmail(t && t.email) === email0 || (!email0 && normStr(t && t.name) === name0);
                });
                if (idx >= 0) {
                    cur.teachers[idx] = Object.assign({}, cur.teachers[idx], {
                        name: nextName || cur.teachers[idx].name,
                        email: nextEmail || cur.teachers[idx].email
                    });
                    ok = true;
                }
            } else if (kind === 'areaPerson' || kind === 'person' || !kind) {
                const areaId =
                    (contextNode && (contextNode.id || contextNode.parentId)) || '';
                if (areaId) {
                    cur.schulOrganigramm = upsertSchulOrganigrammAreaPerson(
                        cur.schulOrganigramm,
                        areaId,
                        { email: nextEmail || email0, name: nextName || name0 }
                    );
                    if (email0 && nextEmail && email0 !== nextEmail) {
                        cur.schulOrganigramm = removeSchulOrganigrammAreaPerson(
                            cur.schulOrganigramm,
                            email0,
                            areaId
                        );
                    }
                    ok = true;
                }
            } else if (kind === 'admin') {
                if (!Array.isArray(cur.admin)) cur.admin = [];
                const idx = findAdminIndex(cur.admin, { email: email0, name: name0 });
                if (idx >= 0) {
                    cur.admin[idx] = Object.assign({}, cur.admin[idx], {
                        name: nextName || cur.admin[idx].name,
                        email: nextEmail || cur.admin[idx].email
                    });
                    syncAdministrationFromAdmin(cur);
                    ok = true;
                }
            }
            if (!ok) {
                toast('Person in den Stammdaten nicht gefunden.', 'warning');
                return;
            }
            saveSettings(cur);
            toast('Person gespeichert.', 'success');
            closeOverlay();
            notify();
        });
    }

    async function openAddCanvasPerson(preselectArea) {
        closeOverlay();
        const settings = loadSettings();
        const layout = normalizeSchulOrganigramm(settings.schulOrganigramm);
        const areas = layout.nodes
            .filter(function (n) {
                if (!isAreaNode(n)) return false;
                if (isSgaBoardArea(n)) return false;
                if (
                    isSammelgruppeArea(n) &&
                    !isSgaScopeArea(n) &&
                    sammelgruppeKeyForArea(n) !== 'lehrer'
                ) {
                    return false;
                }
                return true;
            })
            .sort(function (a, b) {
                return normStr(a.label).localeCompare(normStr(b.label), 'de');
            });
        if (!areas.length) {
            toast('Kein geeigneter Bereich vorhanden – zuerst einen Bereich anlegen.', 'warning');
            return false;
        }
        const preId = preselectArea && preselectArea.id ? String(preselectArea.id) : '';
        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-person-plus',
            title: 'Person anlegen',
            sub: 'Name eintragen und hinzufügen – Microsoft 365 ist optional.',
            body:
                '<div class="field"><label for="schulOrgAddPersonArea">Bereich</label>' +
                '<select id="schulOrgAddPersonArea">' +
                areas
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
                '</select></div>' +
                '<div class="schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgAddPersonName">Name</label>' +
                '<input type="text" id="schulOrgAddPersonName" maxlength="120" autocomplete="name" placeholder="Vor- und Nachname"></div>' +
                '<div class="field"><label for="schulOrgAddPersonEmail">E-Mail <span class="muted">(optional)</span></label>' +
                '<input type="email" id="schulOrgAddPersonEmail" maxlength="160" autocomplete="email" placeholder="name@schule.at"></div>' +
                '</div>' +
                '<p class="muted schul-org-edit-hint">Ohne E-Mail wird eine lokale Kennung erzeugt – später jederzeit änderbar bzw. aus Microsoft 365 zuordenbar.</p>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm alt" data-add-person-m365>' +
                '<i class="bi bi-microsoft" aria-hidden="true"></i> Optional: aus Microsoft 365</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>',
            foot:
                '<span class="schul-org-edit-foot-spacer"></span>' +
                '<button type="button" class="btn alt" data-schul-org-close>Abbrechen</button>' +
                '<button type="button" class="btn btn-success" data-schul-org-save>' +
                '<i class="bi bi-check-lg" aria-hidden="true"></i> Person hinzufügen</button>'
        });
        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const areaEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgAddPersonArea'));
        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAddPersonName'));
        const emailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgAddPersonEmail'));
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };
        nameEl.focus();

        function submitAddPerson() {
            const areaId = areaEl.value;
            const areaNode = layout.nodes.find(function (n) {
                return n.id === areaId;
            });
            if (!areaNode) {
                toast('Bereich nicht gefunden.', 'warning');
                return;
            }
            const name = normStr(nameEl.value);
            let email = normEmail(emailEl.value);
            if (!name && !email) {
                toast('Bitte mindestens einen Namen eingeben.', 'warning');
                nameEl.focus();
                return;
            }
            if (!email) {
                email = suggestLocalPersonEmail(name);
                emailEl.value = email;
            }
            const cur = loadSettings();
            const res = addSchulOrgPersonToArea(cur, areaNode, {
                name: name || email,
                email: email
            });
            if (!res.ok) {
                toast(res.message || 'Speichern fehlgeschlagen.', 'warning');
                return;
            }
            saveSettings(res.settings);
            toast('Person hinzugefügt.', 'success');
            closeOverlay();
            notify();
        }

        overlay.querySelector('[data-schul-org-close]').addEventListener('click', closeOverlay);
        overlay.querySelector('[data-add-person-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'Optional: Person in Microsoft 365 suchen – Name und E-Mail werden übernommen.',
                onPick: function (picked) {
                    nameEl.value = picked.displayName || nameEl.value;
                    emailEl.value = picked.mail || picked.userPrincipalName || emailEl.value;
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-save]').addEventListener('click', submitAddPerson);
        nameEl.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                emailEl.focus();
            }
        });
        emailEl.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                submitAddPerson();
            }
        });
        return true;
    }

    async function pickAndAddPersonToArea(areaNode) {
        if (!areaNode || !isAreaNode(areaNode)) {
            toast('Bitte zuerst einen Bereich wählen.', 'warning');
            return false;
        }
        if (isSgaBoardArea(areaNode)) {
            toast('Rechtsklick auf Schüler, Lehrer oder Eltern unter SGA.', 'info');
            return false;
        }
        if (
            isSammelgruppeArea(areaNode) &&
            !isSgaScopeArea(areaNode) &&
            sammelgruppeKeyForArea(areaNode) !== 'lehrer'
        ) {
            toast(
                'Schüler:innen bitte über die Klassen bzw. Schülerliste pflegen.',
                'warning'
            );
            return false;
        }
        try {
            const picked = await pickEntraUser({
                title: 'Person für „' + (areaNode.label || 'Bereich') + '“',
                hint: 'Wird dem Bereich zugeordnet.'
            });
            if (!picked) return false;
            const cur = loadSettings();
            const res = addSchulOrgPersonToArea(cur, areaNode, {
                email: picked.mail || picked.userPrincipalName || '',
                name: picked.displayName || ''
            });
            if (!res.ok) {
                toast(res.message || 'Hinzufügen fehlgeschlagen.', 'warning');
                return false;
            }
            saveSettings(res.settings);
            notify();
            toast('Person hinzugefügt.', 'success');
            return true;
        } catch (e) {
            toast('Entra-Auswahl fehlgeschlagen: ' + (e && e.message ? e.message : e), 'warning');
            return false;
        }
    }

    return { openEditCanvasPerson, openAddCanvasPerson, pickAndAddPersonToArea };
}
