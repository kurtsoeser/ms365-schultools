/**
 * ARGE editor UI for Schul-Organigramm.
 */
import {
    normalizeSchulOrganigramm,
    NODE,
    isArgeNode,
    listArgeCatalogGroupRefs,
    removeArgeEntityFromSettings,
    applySchulOrganigrammToStammdaten,
    canPlaceSchulOrgNodeInSection,
    childrenOf
} from './schul-organigramm-logic.js';
import { findArge, findTeacherByEmail, fillTeacherSelect, wireTeacherHeadFields } from './schul-organigramm-edit-find.js';
import { roleCodeFromName } from './schul-organigramm-edit-people.js';
import { patchSchulOrganigrammRefLabel, removeSchulOrganigrammByRef, clearHeadOnRecord, setHeadOnRecord } from './schul-organigramm-edit-ops.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';

/**
 * @param {Record<string, any>} ctx
 */
export function createArgeUi(ctx) {
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

    function openArgeEditor(refKey) {
        closeOverlay();
        const settings = loadSettings();
        let arge = findArge(settings, refKey);
        if (!arge) {
            toast('ARGE nicht in den Stammdaten gefunden.', 'warning');
            return;
        }
        arge = Object.assign({}, arge, {
            subjects: Array.isArray(arge.subjects) ? arge.subjects.slice() : []
        });

        const overlay = document.createElement('div');
        overlay.id = 'schulOrgEditOverlay';
        overlay.className = 'modal-overlay open schul-org-edit-overlay';
        overlay.innerHTML = editDialogHtml({
            icon: 'bi-journal-bookmark',
            title: 'ARGE bearbeiten',
            sub: 'Leitung aus Lehrerliste wählen oder aus Microsoft 365 übernehmen.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<h4 class="schul-org-edit-panel-title">Stammdaten</h4>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field"><label for="schulOrgArgeName">Bezeichnung</label>' +
                '<input type="text" id="schulOrgArgeName" maxlength="80" autocomplete="off"></div>' +
                '<div class="field"><label for="schulOrgArgeCode">Kürzel</label>' +
                '<input type="text" id="schulOrgArgeCode" maxlength="24" autocomplete="off" spellcheck="false"></div>' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgArgeSubjects">Fächer (Komma-getrennt)</label>' +
                '<input type="text" id="schulOrgArgeSubjects" maxlength="200" autocomplete="off" placeholder="z. B. MATHE, PHYSIK"></div>' +
                '</div></section>' +
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-panel-head">' +
                '<h4 class="schul-org-edit-panel-title">Leitung (Lehrer-Bezug)</h4>' +
                '<p class="schul-org-edit-hint" id="schulOrgArgeHeadHint"></p></div>' +
                '<div class="form-grid schul-org-edit-grid">' +
                '<div class="field schul-org-edit-field-full"><label for="schulOrgArgeTeacher">Aus Lehrerliste</label>' +
                '<select id="schulOrgArgeTeacher"><option value="">– manuell / M365 –</option></select></div>' +
                '<div class="field"><label for="schulOrgArgeHeadName">Name Leitung</label>' +
                '<input type="text" id="schulOrgArgeHeadName" maxlength="120" autocomplete="name"></div>' +
                '<div class="field"><label for="schulOrgArgeHeadEmail">E-Mail Leitung</label>' +
                '<input type="email" id="schulOrgArgeHeadEmail" maxlength="160" autocomplete="email"></div>' +
                '</div>' +
                '<div class="schul-org-edit-people-actions">' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-arge-m365><i class="bi bi-microsoft" aria-hidden="true"></i> Leitung aus Microsoft 365</button>' +
                '<button type="button" class="btn btn-sm schul-org-edit-btn-ghost" data-schul-org-arge-clear-head><i class="bi bi-x-lg" aria-hidden="true"></i> Leitung leeren</button>' +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '</section>',
            foot: editFootHtml('data-schul-org-del-arge', 'ARGE löschen', 'data-schul-org-save-arge')
        });

        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const nameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeName'));
        const codeEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeCode'));
        const subjEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeSubjects'));
        const teacherEl = /** @type {HTMLSelectElement} */ (overlay.querySelector('#schulOrgArgeTeacher'));
        const headNameEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeHeadName'));
        const headEmailEl = /** @type {HTMLInputElement} */ (overlay.querySelector('#schulOrgArgeHeadEmail'));
        const hintEl = overlay.querySelector('#schulOrgArgeHeadHint');
        const titleEl = overlay.querySelector('#schulOrgEditTitle');
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        nameEl.value = arge.name || '';
        codeEl.value = arge.code || '';
        subjEl.value = (arge.subjects || []).join(', ');
        headNameEl.value = arge.headName || '';
        headEmailEl.value = arge.headEmail || '';

        const teachers = Array.isArray(settings.teachers) ? settings.teachers.slice() : [];
        teachers
            .filter(function (t) {
                return normEmail(t && t.email).indexOf('@') !== -1;
            })
            .sort(function (a, b) {
                return normStr(a.name).localeCompare(normStr(b.name), 'de');
            })
            .forEach(function (t) {
                const opt = document.createElement('option');
                opt.value = normEmail(t.email);
                opt.textContent =
                    (normStr(t.name) || normEmail(t.email)) +
                    (t.code ? ' (' + normStr(t.code) + ')' : '') +
                    ' – ' +
                    normEmail(t.email);
                teacherEl.appendChild(opt);
            });

        function syncTeacherSelect() {
            const em = normEmail(headEmailEl.value);
            teacherEl.value = em && teachers.some(function (t) {
                return normEmail(t.email) === em;
            })
                ? em
                : '';
        }

        function updateHeadHint() {
            if (titleEl) titleEl.textContent = nameEl.value.trim() || arge.name || 'ARGE';
            if (!hintEl) return;
            const em = normEmail(headEmailEl.value);
            if (!em) {
                hintEl.textContent = 'Keine Leitung gesetzt – ARGE ist „offen“.';
                return;
            }
            const t = findTeacherByEmail({ teachers: teachers }, em);
            if (t) {
                hintEl.textContent =
                    'Verknüpft mit Lehrerliste: ' +
                    (t.name || t.code || em) +
                    (t.code ? ' · ' + t.code : '');
            } else {
                hintEl.textContent =
                    'E-Mail gesetzt, aber kein Treffer in der Lehrerliste (trotzdem gültig).';
            }
        }

        teacherEl.addEventListener('change', function () {
            const em = normEmail(teacherEl.value);
            if (!em) return;
            const t = findTeacherByEmail({ teachers: teachers }, em);
            headEmailEl.value = em;
            if (t && t.name) headNameEl.value = t.name;
            updateHeadHint();
        });
        headEmailEl.addEventListener('input', function () {
            syncTeacherSelect();
            updateHeadHint();
        });
        headNameEl.addEventListener('input', updateHeadHint);
        nameEl.addEventListener('input', updateHeadHint);

        syncTeacherSelect();
        updateHeadHint();

        function persistArge(closeAfter) {
            const nextName = normStr(nameEl.value);
            const nextCode = normStr(codeEl.value).toUpperCase() || roleCodeFromName(nextName);
            if (!nextName) {
                toast('Bitte eine Bezeichnung eingeben.', 'warning');
                return;
            }
            if (!nextCode) {
                toast('Bitte ein Kürzel eingeben.', 'warning');
                return;
            }
            const subjects = String(subjEl.value || '')
                .split(/[,;\s]+/)
                .map(function (x) {
                    return normStr(x).toUpperCase();
                })
                .filter(Boolean);
            const cur = loadSettings();
            const arges = Array.isArray(cur.arges) ? cur.arges.slice() : [];
            const oldCode = normStr(arge.code).toUpperCase();
            let found = arges.find(function (a) {
                return normStr(a.code).toUpperCase() === oldCode;
            });
            const clash = arges.some(function (a) {
                return (
                    a !== found &&
                    normStr(a.code).toUpperCase() === nextCode
                );
            });
            if (clash) {
                toast('Kürzel bereits vergeben.', 'warning');
                return;
            }
            if (!found) {
                found = { code: nextCode, name: nextName, subjects: subjects };
                arges.push(found);
            } else {
                found.code = nextCode;
                found.name = nextName;
                found.subjects = subjects;
            }
            found.headName = normStr(headNameEl.value);
            found.headEmail = normEmail(headEmailEl.value);
            cur.arges = arges;
            patchSchulOrganigrammRefLabel(cur, 'arge', oldCode || refKey, found.name, found.code);
            saveSettings(cur);
            arge = Object.assign({}, found);
            toast('ARGE gespeichert.', 'success');
            notify();
            if (closeAfter) closeOverlay();
            else {
                nameEl.value = found.name || '';
                codeEl.value = found.code || '';
                subjEl.value = (found.subjects || []).join(', ');
                headNameEl.value = found.headName || '';
                headEmailEl.value = found.headEmail || '';
                syncTeacherSelect();
                updateHeadHint();
            }
        }

        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-arge]').addEventListener('click', function () {
            persistArge(true);
        });
        overlay.querySelector('[data-schul-org-arge-clear-head]').addEventListener('click', function () {
            headNameEl.value = '';
            headEmailEl.value = '';
            teacherEl.value = '';
            updateHeadHint();
        });
        overlay.querySelector('[data-schul-org-arge-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'ARGE-Leitung in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    const mail = normEmail(picked.mail || picked.userPrincipalName || picked.email);
                    const name = normStr(picked.displayName || picked.name);
                    if (!mail || mail.indexOf('@') === -1) {
                        toast('Keine gültige E-Mail in der Auswahl.', 'warning');
                        return;
                    }
                    headEmailEl.value = mail;
                    if (name) headNameEl.value = name;
                    syncTeacherSelect();
                    updateHeadHint();
                    toast('Leitung aus Microsoft 365 übernommen.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelector('[data-schul-org-del-arge]').addEventListener('click', function () {
            void confirmDlg(
                'ARGE „' + (arge.name || arge.code) + '“ aus Stammdaten und Landkarte löschen?',
                {
                    title: 'ARGE löschen',
                    okText: 'Löschen',
                    cancelText: 'Abbrechen'
                }
            ).then(function (ok) {
                if (!ok) return;
                const cur = loadSettings();
                const code = normStr(arge.code).toUpperCase();
                const next = removeArgeEntityFromSettings(cur, code);
                Object.keys(next).forEach(function (k) {
                    cur[k] = next[k];
                });
                saveSettings(cur);
                toast('ARGE gelöscht.', 'success');
                closeOverlay();
                notify();
            });
        });

        nameEl.focus();
    }

    /**
     * @param {string} [parentSectionId]
     */
    async function createArgeUnderSection(parentSectionId) {
        let name = '';
        if (typeof window.ms365AppDialogPrompt === 'function') {
            name = await window.ms365AppDialogPrompt('Bezeichnung der neuen ARGE / Fachschaft', '', {
                title: 'ARGE anlegen',
                inputLabel: 'Bezeichnung'
            });
        } else {
            name = window.prompt('Bezeichnung der neuen ARGE', '') || '';
        }
        const label = normStr(name);
        if (!label) return;
        const cur = loadSettings();
        const exists = (cur.arges || []).some(function (a) {
            return normStr(a.name).toLowerCase() === label.toLowerCase();
        });
        if (exists) {
            toast('Diese ARGE gibt es bereits.', 'warning');
            return;
        }
        const used = new Set(
            (cur.arges || []).map(function (a) {
                return normStr(a.code).toUpperCase();
            })
        );
        let code = roleCodeFromName(label);
        let i = 2;
        while (used.has(code)) {
            code = (roleCodeFromName(label) + String(i)).slice(0, 24);
            i += 1;
        }
        const arge = {
            code: code,
            name: label,
            subjects: [],
            headName: '',
            headEmail: ''
        };
        cur.arges = (cur.arges || []).concat([arge]);
        const layout =
            cur.schulOrganigramm && Array.isArray(cur.schulOrganigramm.nodes)
                ? cur.schulOrganigramm
                : { version: 1, nodes: [] };
        const maxOrder = layout.nodes.reduce(function (m, n) {
            return Math.max(m, n.order || 0);
        }, 0);
        layout.nodes.push({
            id: 'arge_' + code,
            type: 'area',
            label: label,
            parentId: parentSectionId || 'sec_arge',
            order: maxOrder + 1,
            tone: 'arge',
            refKind: 'arge',
            refKey: code,
            meta: { nestPeople: true, audience: 'arge' }
        });
        cur.schulOrganigramm = layout;
        saveSettings(cur);
        toast('ARGE angelegt.', 'success');
        notify();
        openArgeEditor(code);
    }

    return { openArgeEditor, createArgeUnderSection };
}
