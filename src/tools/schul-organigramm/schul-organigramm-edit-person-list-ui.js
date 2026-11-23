/**
 * Person list / council / SGA editors for Schul-Organigramm.
 */
import {
    normalizeSchulOrganigramm,
    NODE,
    SCHUL_ORG_BOARD_REF,
    SCHUL_ORG_SV_SEATS,
    ensureSchulOrganigrammSgaStructure,
    applySchulOrganigrammToStammdaten,
    childrenOf,
    sgaScopeForArea,
    svSeatForNode
} from './schul-organigramm-logic.js';
import { findStudentCouncil, findSga, sgaScopeLabel, normalizeSgaScope } from './schul-organigramm-edit-find.js';
import { moveSchulOrgPerson, patchSchulOrganigrammRefLabel, removeSchulOrganigrammByRef } from './schul-organigramm-edit-ops.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';
import { createBoardUi } from './schul-organigramm-edit-board-ui.js';

/**
 * @param {Record<string, any>} ctx
 */
export function createPersonListUi(ctx) {
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

    const boardUi = createBoardUi(ctx);
    const openBoardListEditor = boardUi.openBoardListEditor;

    function openPersonListEditor(kind, refKey) {
        closeOverlay();
        const isCouncil = kind === 'council';
        const settings = loadSettings();
        let row = isCouncil ? findStudentCouncil(settings, refKey) : findSga(settings, refKey);
        if (!row) {
            toast((isCouncil ? 'Schülervertretung' : 'SGA-Eintrag') + ' nicht gefunden.', 'warning');
            return;
        }
        row = Object.assign({}, row);
        const listKey = isCouncil ? 'studentCouncil' : 'sga';
        const nodeType = isCouncil ? 'council' : 'sga';
        const titleDefault = isCouncil ? 'Schülervertretung' : 'SGA';

        const classOptions = (Array.isArray(settings.classes) ? settings.classes : [])
            .map(function (c) {
                return normStr(c.code).toUpperCase() || normStr(c.name);
            })
            .filter(Boolean)
            .sort(function (a, b) {
                return a.localeCompare(b, 'de');
            });

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: isCouncil ? 'bi-people' : 'bi-building',
            title: titleDefault + ' bearbeiten',
            sub: isCouncil
                ? 'Klassenvertretung: Name, E-Mail und Klasse.'
                : 'Schulgemeinschaftsausschuss: Name, E-Mail und Bereich.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Person</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgPersName">Name</label>' +
                '<input type="text" id="schulOrgPersName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgPersEmail">E-Mail</label>' +
                '<input type="email" id="schulOrgPersEmail" maxlength="160" autocomplete="email"></div>' +
                (isCouncil
                    ? '<div class="field"><label for="schulOrgPersKlasse">Klasse</label>' +
                      '<select id="schulOrgPersKlasse"><option value="">– wählen / frei –</option></select></div>' +
                      '<div class="field"><label for="schulOrgPersKlasseFree">Klasse (frei)</label>' +
                      '<input type="text" id="schulOrgPersKlasseFree" maxlength="24" autocomplete="off" spellcheck="false"></div>'
                    : '<div class="field schul-org-edit-field-full"><label for="schulOrgPersScope">Bereich</label>' +
                      '<select id="schulOrgPersScope">' +
                      '<option value="teacher">Lehrer</option>' +
                      '<option value="student">Schüler</option>' +
                      '<option value="external">Extern</option>' +
                      '</select></div>') +
                '</div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-pers-m365><i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-pers', 'Löschen', 'data-schul-org-save-pers')
        });

        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgPersName'));
        const emailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgPersEmail'));
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };
        nameEl.value = row.name || '';
        emailEl.value = row.email || '';

        /** @type {HTMLSelectElement|null} */
        let klasseEl = null;
        /** @type {HTMLInputElement|null} */
        let klasseFreeEl = null;
        /** @type {HTMLSelectElement|null} */
        let scopeEl = null;

        if (isCouncil) {
            klasseEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgPersKlasse'));
            klasseFreeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgPersKlasseFree'));
            classOptions.forEach(function (c) {
                const opt = document.createElement('option');
                opt.value = c;
                opt.textContent = c;
                klasseEl.appendChild(opt);
            });
            const k = normStr(row.klasse).toUpperCase();
            if (k && classOptions.indexOf(k) !== -1) {
                klasseEl.value = k;
                klasseFreeEl.value = '';
            } else {
                klasseEl.value = '';
                klasseFreeEl.value = normStr(row.klasse);
            }
        } else {
            scopeEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgPersScope'));
            const sc = normalizeSgaScope(row.scope) || 'teacher';
            scopeEl.value = sc;
        }

        function updateTitle() {
            if (titleEl) titleEl.textContent = (nameEl.value.trim() || row.name || titleDefault) + ' bearbeiten';
        }
        nameEl.addEventListener('input', updateTitle);
        updateTitle();

        let oldRef = normEmail(row.email) || normStr(row.name) || refKey;

        function persistPerson(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextEmail = normEmail(emailEl.value);
            if (!nextName && !nextEmail) {
                toast('Bitte Name oder E-Mail eingeben.', 'warning');
                return;
            }
            if (nextEmail && nextEmail.indexOf('@') === -1) {
                toast('E-Mail ungültig.', 'warning');
                return;
            }
            const cur = loadSettings();
            const list = Array.isArray(cur[listKey]) ? cur[listKey].slice() : [];
            let found =
                (normEmail(row.email)
                    ? list.find(function (s) {
                          return normEmail(s.email) === normEmail(row.email);
                      })
                    : null) ||
                list.find(function (s) {
                    return (
                        normStr(s.name).toLowerCase() === normStr(row.name).toLowerCase() &&
                        (!isCouncil || normStr(s.klasse).toLowerCase() === normStr(row.klasse).toLowerCase()) &&
                        (isCouncil || normalizeSgaScope(s.scope) === normalizeSgaScope(row.scope))
                    );
                });
            if (nextEmail) {
                const clash = list.some(function (s) {
                    return s !== found && normEmail(s.email) === nextEmail;
                });
                if (clash) {
                    toast('Diese E-Mail ist bereits eingetragen.', 'warning');
                    return;
                }
            }
            if (!found) {
                found = isCouncil
                    ? { klasse: '', name: nextName, email: nextEmail }
                    : { scope: 'teacher', name: nextName, email: nextEmail };
                list.push(found);
            }
            found.name = nextName;
            found.email = nextEmail;
            if (isCouncil) {
                const fromSelect = klasseEl ? normStr(klasseEl.value).toUpperCase() : '';
                const fromFree = klasseFreeEl ? normStr(klasseFreeEl.value).toUpperCase() : '';
                found.klasse = fromSelect || fromFree;
            } else {
                found.scope = scopeEl ? normalizeSgaScope(scopeEl.value) || 'teacher' : 'teacher';
            }
            cur[listKey] = list;
            const nextRef = nextEmail || nextName;
            patchSchulOrganigrammRefLabel(cur, nodeType, oldRef, nextName || nextEmail, nextRef);
            saveSettings(cur);
            row = Object.assign({}, found);
            oldRef = nextRef;
            toast(titleDefault + ' gespeichert.', 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                nameEl.value = found.name || '';
                emailEl.value = found.email || '';
                updateTitle();
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-pers]').addEventListener('click', function () {
            persistPerson(true);
        });
        overlay.querySelector('[data-schul-org-pers-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: titleDefault + ' in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    const mail = normEmail(picked.mail || picked.userPrincipalName || picked.email);
                    const name = normStr(picked.displayName || picked.name);
                    if (mail && mail.indexOf('@') !== -1) emailEl.value = mail;
                    if (name) nameEl.value = name;
                    updateTitle();
                    toast('Aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-pers]').addEventListener('click', function () {
            void confirmDlg('„' + (row.name || row.email || titleDefault) + '“ wirklich löschen?', {
                title: titleDefault + ' löschen',
                okText: 'Löschen',
                cancelText: 'Abbrechen'
            }).then(function (ok) {
                if (!ok) return;
                const cur = loadSettings();
                const mail = normEmail(row.email);
                const nameL = normStr(row.name).toLowerCase();
                cur[listKey] = (cur[listKey] || []).filter(function (s) {
                    if (mail && normEmail(s.email) === mail) return false;
                    if (
                        !mail &&
                        nameL &&
                        normStr(s.name).toLowerCase() === nameL &&
                        (!isCouncil || normStr(s.klasse).toLowerCase() === normStr(row.klasse).toLowerCase()) &&
                        (isCouncil || normalizeSgaScope(s.scope) === normalizeSgaScope(row.scope))
                    ) {
                        return false;
                    }
                    return true;
                });
                removeSchulOrganigrammByRef(cur, nodeType, oldRef);
                saveSettings(cur);
                toast(titleDefault + ' gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        nameEl.focus();
    }

    /**
     * Sammel-Editor: alle Personen der Schülervertretung bzw. SGA (nach Scope).
     * @param {string} refKey
     */
    function openCouncilEditor(refKey) {
        if (!refKey || refKey === SCHUL_ORG_BOARD_REF || refKey === 'council_board') {
            openBoardListEditor('council');
            return;
        }
        const seat = svSeatForNode({ id: '', refKey: refKey, type: 'council' }) ||
            SCHUL_ORG_SV_SEATS.find(function (s) {
                return s.id === refKey || s.refKey === refKey;
            });
        if (seat) {
            // Einzelner Sitz → gesamtes SV-Board (drei Funktionen)
            openBoardListEditor('council');
            return;
        }
        openPersonListEditor('council', refKey);
    }

    function openSgaEditor(refKey) {
        if (!refKey || refKey === SCHUL_ORG_BOARD_REF || refKey === 'sga_board') {
            openBoardListEditor('sga');
            return;
        }
        openPersonListEditor('sga', refKey);
    }

    /**
     * @param {string} [parentSectionId]
     * @param {'council'|'sga'} kind
     */
    async function createPersonUnderSection(parentSectionId, kind) {
        openBoardListEditor(kind);
    }

    async function createCouncilUnderSection(parentSectionId) {
        return createPersonUnderSection(parentSectionId, 'council');
    }

    async function createSgaUnderSection(parentSectionId) {
        return createPersonUnderSection(parentSectionId, 'sga');
    }

    function movePersonToNode(person, targetNode) {
        const cur = loadSettings();
        const result = moveSchulOrgPerson(cur, person, targetNode);
        if (!result.ok || !result.settings) return result;
        saveSettings(result.settings);
        notify();
        return result;
    }

    return {
        openPersonListEditor,
        openBoardListEditor,
        openCouncilEditor,
        openSgaEditor,
        createPersonUnderSection,
        createCouncilUnderSection,
        createSgaUnderSection,
        movePersonToNode
    };
}
