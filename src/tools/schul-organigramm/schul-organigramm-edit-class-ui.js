/**
 * Class editor UI for Schul-Organigramm.
 */
import {
    normalizeSchulOrganigramm,
    NODE,
    isClassNode,
    canonicalSchulOrgClassCode,
    removeClassEntityFromSettings,
    renameClassEntityInSettings,
    applySchulOrganigrammToStammdaten,
    canPlaceSchulOrgNodeInSection,
    childrenOf,
    jahrgangAreaId,
    schulstufeFromClassCode
} from './schul-organigramm-logic.js';
import { findClass, findTeacherByEmail, fillTeacherSelect, wireTeacherHeadFields } from './schul-organigramm-edit-find.js';
import { roleCodeFromName } from './schul-organigramm-edit-people.js';
import { patchSchulOrganigrammRefLabel, removeSchulOrganigrammByRef, clearHeadOnRecord, setHeadOnRecord, ensureClassHeadsNamedFromTeachers } from './schul-organigramm-edit-ops.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';

/**
 * @param {Record<string, any>} ctx
 */
export function createClassUi(ctx) {
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

    function openClassEditor(refKey) {
        closeOverlay();
        const settings = loadSettings();
        let cls = findClass(settings, refKey);
        if (!cls) {
            toast('Klasse nicht in den Stammdaten gefunden.', 'warning');
            return;
        }
        cls = Object.assign({}, cls);

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-mortarboard',
            title: 'Klasse bearbeiten',
            sub: 'Klassenvorstand aus Lehrerliste oder Microsoft 365 zuordnen.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Klasse</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgClassName">Bezeichnung</label>' +
                '<input type="text" id="schulOrgClassName" maxlength="80" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgClassCode">Kürzel</label>' +
                '<input type="text" id="schulOrgClassCode" maxlength="24" autocomplete="off" spellcheck="false"></div>' +
                '<div class="field"><label for="schulOrgClassYear">Abschlussjahr</label>' +
                '<input type="text" id="schulOrgClassYear" maxlength="4" inputmode="numeric" pattern="[0-9]{4}" placeholder="z. B. 2030" autocomplete="off"></div>' +
                '</div></section>' +
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-panel-head">' +
                '<h4 class="schul-org-edit-panel-title">Klassenvorstand</h4>' +
                '<p class="schul-org-edit-hint" id="schulOrgClassHeadHint"></p></div>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgClassTeacher">Aus Lehrerliste</label>' +
                '<select id="schulOrgClassTeacher"><option value="">– manuell / M365 –</option></select></div>' +
                '<div class="field"><label for="schulOrgClassHeadName">Name KV</label>' +
                '<input type="text" id="schulOrgClassHeadName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgClassHeadEmail">E-Mail KV</label>' +
                '<input type="email" id="schulOrgClassHeadEmail" maxlength="160" autocomplete="email"></div>' +
                '</div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-class-m365><i class="bi bi-microsoft" aria-hidden="true"></i> KV aus Microsoft 365</button>' +
                '<button type="button" class="btn btn-sm schul-org-edit-btn-ghost" data-schul-org-class-clear-head><i class="bi bi-x-lg" aria-hidden="true"></i> KV leeren</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-class', 'Klasse löschen', 'data-schul-org-save-class')
        });

        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassName'));
        const codeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassCode'));
        const yearEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassYear'));
        const teacherEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgClassTeacher'));
        const headNameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassHeadName'));
        const headEmailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgClassHeadEmail'));
        const hintEl = overlay.querySelector('#schulOrgClassHeadHint');
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        nameEl.value = cls.name || '';
        codeEl.value = cls.code || '';
        yearEl.value = cls.year || '';
        headNameEl.value = cls.headName || '';
        headEmailEl.value = cls.headEmail || '';

        const teachers = Array.isArray(settings.teachers) ? settings.teachers.slice() : [];
        fillTeacherSelect(teacherEl, teachers);
        const headWire = wireTeacherHeadFields({
            teacherEl: teacherEl,
            headNameEl: headNameEl,
            headEmailEl: headEmailEl,
            hintEl: hintEl,
            teachers: teachers,
            vacantText: 'Kein Klassenvorstand – KV ist „offen“.'
        });

        nameEl.addEventListener('input', function () {
            if (titleEl) titleEl.textContent = nameEl.value.trim() || cls.name || 'Klasse';
        });
        if (titleEl) titleEl.textContent = nameEl.value.trim() || cls.name || 'Klasse';

        function persistClass(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextCode = normStr(codeEl.value).toUpperCase() || roleCodeFromName(nextName);
            const yearRaw = normStr(yearEl.value);
            const nextYear = /^\d{4}$/.test(yearRaw) ? yearRaw : '';
            if (!nextName && !nextCode) {
                toast('Bitte Bezeichnung oder Kürzel eingeben.', 'warning');
                return;
            }
            if (!nextCode) {
                toast('Bitte ein Kürzel eingeben.', 'warning');
                return;
            }
            if (yearRaw && !nextYear) {
                toast('Abschlussjahr muss vierstellig sein (z. B. 2030).', 'warning');
                return;
            }
            const cur = loadSettings();
            const classes = Array.isArray(cur.classes) ? cur.classes.slice() : [];
            const oldCode = normStr(cls.code).toUpperCase() || normStr(refKey).toUpperCase();
            let found = classes.find(function (c) {
                return normStr(c.code).toUpperCase() === oldCode;
            });
            const clash = classes.some(function (c) {
                return c !== found && normStr(c.code).toUpperCase() === nextCode;
            });
            if (clash) {
                toast('Klassenkürzel bereits vergeben.', 'warning');
                return;
            }
            if (!found) {
                found = {
                    code: oldCode || nextCode,
                    name: nextName || nextCode,
                    year: nextYear,
                    headName: '',
                    headEmail: '',
                    stableMailNickname: ''
                };
                classes.push(found);
                cur.classes = classes;
            }
            // Rename propagiert Schüler/Council/Chats/Belegung + Landkarte
            const renamed = renameClassEntityInSettings(cur, oldCode || nextCode, nextCode, {
                name: nextName || nextCode
            });
            Object.keys(renamed.settings).forEach(function (k) {
                cur[k] = renamed.settings[k];
            });
            const nextClasses = Array.isArray(cur.classes) ? cur.classes : [];
            found =
                nextClasses.find(function (c) {
                    return normStr(c.code).toUpperCase() === nextCode;
                }) || found;
            found.year = nextYear;
            found.headName = normStr(headNameEl.value);
            found.headEmail = normEmail(headEmailEl.value);
            cur.classes = nextClasses.map(function (c) {
                return normStr(c.code).toUpperCase() === nextCode ? found : c;
            });
            saveSettings(cur);
            cls = Object.assign({}, found);
            const extra =
                renamed.renamedStudents > 0
                    ? ' · ' + renamed.renamedStudents + ' Schüler:innen umgehängt'
                    : '';
            toast('Klasse gespeichert.' + extra, 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                nameEl.value = found.name || '';
                codeEl.value = found.code || '';
                yearEl.value = found.year || '';
                headNameEl.value = found.headName || '';
                headEmailEl.value = found.headEmail || '';
                headWire.syncTeacherSelect();
                headWire.updateHeadHint();
                if (titleEl) titleEl.textContent = found.name || found.code || 'Klasse';
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-class]').addEventListener('click', function () {
            persistClass(true);
        });
        overlay.querySelector('[data-schul-org-class-clear-head]').addEventListener('click', function () {
            headNameEl.value = '';
            headEmailEl.value = '';
            teacherEl.value = '';
            headWire.updateHeadHint();
        });
        overlay.querySelector('[data-schul-org-class-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'Klassenvorstand in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    const mail = normEmail(picked.mail || picked.userPrincipalName || picked.email);
                    const name = normStr(picked.displayName || picked.name);
                    if (!mail || mail.indexOf('@') === -1) {
                        toast('Keine gültige E-Mail in der Auswahl.', 'warning');
                        return;
                    }
                    headEmailEl.value = mail;
                    if (name) headNameEl.value = name;
                    headWire.syncTeacherSelect();
                    headWire.updateHeadHint();
                    toast('Klassenvorstand aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-class]').addEventListener('click', function () {
            const code = normStr(cls.code).toUpperCase();
            const preview = removeClassEntityFromSettings(loadSettings(), code);
            const studentHint =
                preview.studentCount > 0
                    ? ' ' +
                      preview.studentCount +
                      ' Schüler:innen behalten die Kennung „' +
                      code +
                      '“ (werden nicht gelöscht).'
                    : '';
            void confirmDlg(
                'Klasse „' +
                    (cls.name || cls.code) +
                    '“ aus Stammdaten und Landkarte löschen?' +
                    studentHint,
                {
                    title: 'Klasse löschen',
                    okText: 'Löschen',
                    cancelText: 'Abbrechen'
                }
            ).then(function (ok) {
                if (!ok) return;
                const cur = loadSettings();
                const removed = removeClassEntityFromSettings(cur, code);
                Object.keys(removed.settings).forEach(function (k) {
                    cur[k] = removed.settings[k];
                });
                saveSettings(cur);
                toast('Klasse gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        nameEl.focus();
    }

    /**
     * @param {string} [parentSectionId]
     */
    async function createClassUnderSection(parentSectionId) {
        let codeIn = '';
        if (typeof window.ms365AppDialogPrompt === 'function') {
            codeIn = await window.ms365AppDialogPrompt('Kürzel der neuen Klasse (z. B. 1A)', '', {
                title: 'Klasse anlegen',
                inputLabel: 'Kürzel'
            });
        } else {
            codeIn = window.prompt('Kürzel der neuen Klasse (z. B. 1A)', '') || '';
        }
        const code = normStr(codeIn).toUpperCase();
        if (!code) return;
        const cur = loadSettings();
        if ((cur.classes || []).some(function (c) {
            return normStr(c.code).toUpperCase() === code;
        })) {
            toast('Diese Klasse gibt es bereits.', 'warning');
            return;
        }
        const cls = {
            code: code,
            name: code,
            year: '',
            headName: '',
            headEmail: '',
            stableMailNickname: ''
        };
        cur.classes = (cur.classes || []).concat([cls]);
        const layout =
            cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)
                ? cur.schulOrganigramm
                : { version: 1, nodes: [] };
        const maxOrder = layout.nodes.reduce(function (m, n) {
            return Math.max(m, n.order || 0);
        }, 0);
        layout.nodes.push({
            id: 'class_' + code,
            type: 'class',
            label: code,
            parentId: parentSectionId || jahrgangAreaId(schulstufeFromClassCode(code)),
            order: maxOrder + 1,
            tone: 'class',
            refKind: 'class',
            refKey: code
        });
        cur.schulOrganigramm = layout;
        saveSettings(cur);
        toast('Klasse angelegt.', 'success');
        notify();
        openClassEditor(code);
    }

    return { openClassEditor, createClassUnderSection };
}
