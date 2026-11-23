/**
 * Board list editor (SV/SGA collective seats).
 */
import {
    normalizeSchulOrganigramm,
    NODE,
    SCHUL_ORG_BOARD_REF,
    SCHUL_ORG_SV_SEATS,
    ensureSchulOrganigrammSgaStructure,
    applySchulOrganigrammToStammdaten
} from './schul-organigramm-logic.js';
import { sgaScopeLabel, normalizeSgaScope } from './schul-organigramm-edit-find.js';
import { writeSchoolData } from '../../shared/stammdaten-write.js';
import { pickEntraUser } from '../../shared/entra-user-picker.js';

/**
 * @param {Record<string, any>} ctx
 */
export function createBoardUi(ctx) {
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

    function openBoardListEditor(kind) {
        closeOverlay();
        const isCouncil = kind === 'council';
        const listKey = isCouncil ? 'studentCouncil' : 'sga';
        const title = isCouncil ? 'Schülervertretung' : 'SGA';
        const settings = loadSettings();
        /** @type {object[]} */
        let rows = (Array.isArray(settings[listKey]) ? settings[listKey] : []).map(function (r) {
            return Object.assign({}, r);
        });

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
            title: title,
            sub: isCouncil
                ? 'Alle Schülervertreter:innen – Name, E-Mail, Klasse.'
                : 'Schulgemeinschaftsausschuss nach Gruppen: Schüler, Lehrer, Eltern.',
            body:
                '<section class="schul-org-edit-panel">' +
                '<div class="schul-org-edit-people-actions" style="margin-bottom:10px">' +
                '<button type="button" class="btn btn-sm btn-success" data-schul-org-board-add>' +
                '<i class="bi bi-plus-lg" aria-hidden="true"></i> Person</button>' +
                '<button type="button" class="btn btn-sm btn-primary" data-schul-org-board-m365>' +
                '<i class="bi bi-microsoft" aria-hidden="true"></i> Aus Microsoft 365</button>' +
                (!isCouncil
                    ? '<select id="schulOrgBoardAddScope" class="schul-org-edit-inline-select" title="Gruppe">' +
                      '<option value="student">Schüler</option>' +
                      '<option value="teacher">Lehrer</option>' +
                      '<option value="external">Eltern</option>' +
                      '</select>'
                    : '') +
                '</div>' +
                '<div class="schul-org-inline-eup-host" data-schul-org-inline-eup hidden></div>' +
                '<div id="schulOrgBoardList" class="schul-org-edit-board-list"></div>' +
                '</section>',
            foot:
                '<button type="button" class="btn" data-schul-org-close>Schließen</button>' +
                '<button type="button" class="btn btn-primary" data-schul-org-save-board>Speichern</button>'
        });
        document.body.appendChild(overlay);
        setActiveOverlay(overlay);
        document.addEventListener('keydown', onEsc, true);

        const listHost = /** @type {HTMLElement} */ (overlay.querySelector('#schulOrgBoardList'));
        const scopeAddEl = /** @type {HTMLSelectElement|null} */ (
            overlay.querySelector('#schulOrgBoardAddScope')
        );
        const inlineEupHost = /** @type {HTMLElement|null} */ (
            overlay.querySelector('[data-schul-org-inline-eup]')
        );
        /** @type {{ api: { destroy: () => void, focus: () => void }|null }} */
        const inlineEupState = { api: null };

        function ensureBoardNodes(cur) {
            cur.schulOrganigramm = ensureSchulOrganigrammSgaStructure(cur.schulOrganigramm, cur);
        }

        function renderBoard() {
            listHost.replaceChildren();
            if (!rows.length) {
                const empty = document.createElement('p');
                empty.className = 'muted';
                empty.textContent = 'Noch keine Personen – „Person“ oder Microsoft 365 nutzen.';
                listHost.appendChild(empty);
                return;
            }
            /** @type {{ key: string, label: string, items: object[] }[]} */
            const groups = isCouncil
                ? [{ key: 'all', label: 'Schülervertretung (SGA)', items: rows }]
                : [
                      {
                          key: 'student',
                          label: 'Schülervertretung',
                          items: rows.filter(function (r) {
                              return (normalizeSgaScope(r.scope) || 'student') === 'student';
                          })
                      },
                      {
                          key: 'teacher',
                          label: 'Lehrer',
                          items: rows.filter(function (r) {
                              return normalizeSgaScope(r.scope) === 'teacher';
                          })
                      },
                      {
                          key: 'external',
                          label: 'Eltern',
                          items: rows.filter(function (r) {
                              return normalizeSgaScope(r.scope) === 'external';
                          })
                      }
                  ];
            groups.forEach(function (g) {
                const block = document.createElement('div');
                block.className = 'schul-org-edit-board-group';
                const h = document.createElement('h4');
                h.className = 'schul-org-edit-panel-title';
                h.textContent = g.label + ' (' + g.items.length + ')';
                block.appendChild(h);
                if (!g.items.length) {
                    const p = document.createElement('p');
                    p.className = 'muted';
                    p.textContent = 'leer';
                    block.appendChild(p);
                }
                g.items.forEach(function (row) {
                    const idx = rows.indexOf(row);
                    const seatHint =
                        isCouncil && SCHUL_ORG_SV_SEATS[idx]
                            ? SCHUL_ORG_SV_SEATS[idx].label
                            : g.key === 'student' && SCHUL_ORG_SV_SEATS[g.items.indexOf(row)]
                              ? SCHUL_ORG_SV_SEATS[g.items.indexOf(row)].label
                              : '';
                    const line = document.createElement('div');
                    line.className = 'schul-org-edit-board-row';
                    line.innerHTML =
                        (seatHint
                            ? '<span class="muted schul-org-edit-seat-label">' +
                              escapeHtml(seatHint) +
                              '</span>'
                            : '') +
                        '<input type="text" data-f="name" placeholder="Name" maxlength="120" />' +
                        '<input type="email" data-f="email" placeholder="E-Mail" maxlength="160" />' +
                        (isCouncil
                            ? '<input type="text" data-f="klasse" placeholder="Klasse" maxlength="24" list="schulOrgBoardKlassen" />'
                            : '<select data-f="scope">' +
                              '<option value="student">Schülervertretung</option>' +
                              '<option value="teacher">Lehrer</option>' +
                              '<option value="external">Eltern</option>' +
                              '</select>') +
                        '<button type="button" class="btn btn-sm" data-del title="Entfernen"><i class="bi bi-trash" aria-hidden="true"></i></button>';
                    const nameInp = /** @type {HTMLInputElement} */ (line.querySelector('[data-f="name"]'));
                    const emailInp = /** @type {HTMLInputElement} */ (line.querySelector('[data-f="email"]'));
                    nameInp.value = row.name || '';
                    emailInp.value = row.email || '';
                    nameInp.addEventListener('input', function () {
                        rows[idx].name = nameInp.value;
                    });
                    emailInp.addEventListener('input', function () {
                        rows[idx].email = emailInp.value;
                    });
                    if (isCouncil) {
                        const kInp = /** @type {HTMLInputElement} */ (line.querySelector('[data-f="klasse"]'));
                        kInp.value = row.klasse || '';
                        kInp.addEventListener('input', function () {
                            rows[idx].klasse = kInp.value;
                        });
                    } else {
                        const scEl = /** @type {HTMLSelectElement} */ (line.querySelector('[data-f="scope"]'));
                        scEl.value = normalizeSgaScope(row.scope) || 'student';
                        scEl.addEventListener('change', function () {
                            rows[idx].scope = scEl.value;
                            renderBoard();
                        });
                    }
                    line.querySelector('[data-del]').addEventListener('click', function () {
                        rows.splice(idx, 1);
                        renderBoard();
                    });
                    block.appendChild(line);
                });
                listHost.appendChild(block);
            });
            if (isCouncil && !overlay.querySelector('#schulOrgBoardKlassen')) {
                const dl = document.createElement('datalist');
                dl.id = 'schulOrgBoardKlassen';
                classOptions.forEach(function (c) {
                    const opt = document.createElement('option');
                    opt.value = c;
                    dl.appendChild(opt);
                });
                overlay.appendChild(dl);
            }
        }

        function addRow(preset) {
            const p = preset && typeof preset === 'object' ? preset : {};
            if (isCouncil) {
                rows.push({
                    name: normStr(p.name),
                    email: normEmail(p.email),
                    klasse: normStr(p.klasse).toUpperCase()
                });
            } else {
                rows.push({
                    name: normStr(p.name),
                    email: normEmail(p.email),
                    scope: normalizeSgaScope(p.scope) || (scopeAddEl && scopeAddEl.value) || 'student'
                });
            }
            renderBoard();
        }

        overlay.querySelector('[data-schul-org-board-add]').addEventListener('click', function () {
            addRow({});
        });
        overlay.querySelector('[data-schul-org-board-m365]').addEventListener('click', function () {
            void toggleInlineEntraUserPicker(inlineEupHost, inlineEupState, {
                hint: 'Person in Microsoft 365 suchen und antippen.',
                onPick: function (picked) {
                    addRow({
                        name: normStr(picked.displayName || picked.name),
                        email: normEmail(picked.mail || picked.userPrincipalName || picked.email),
                        scope: scopeAddEl ? scopeAddEl.value : 'student'
                    });
                    toast('Aus Microsoft 365 hinzugefügt.', 'success');
                }
            }).catch(function (e) {
                toast('M365: ' + (e && e.message ? e.message : String(e)), 'error');
            });
        });
        overlay.querySelectorAll('[data-schul-org-close]').forEach(function (btn) {
            btn.addEventListener('click', closeOverlay);
        });
        overlay.addEventListener('click', function (ev) {
            if (ev.target === overlay) closeOverlay();
        });
        overlay.querySelector('[data-schul-org-save-board]').addEventListener('click', function () {
            const cleaned = [];
            const seen = new Set();
            let badMail = '';
            rows.forEach(function (r) {
                const name = normStr(r.name);
                const email = normEmail(r.email);
                if (!name && !email) return;
                if (email && email.indexOf('@') === -1) {
                    badMail = email;
                    return;
                }
                const key =
                    email ||
                    name.toLowerCase() +
                        '|' +
                        (isCouncil ? normStr(r.klasse) : normalizeSgaScope(r.scope));
                if (seen.has(key)) return;
                seen.add(key);
                if (isCouncil) {
                    cleaned.push({ name: name, email: email, klasse: normStr(r.klasse).toUpperCase() });
                } else {
                    cleaned.push({
                        name: name,
                        email: email,
                        scope: normalizeSgaScope(r.scope) || 'student'
                    });
                }
            });
            if (badMail) {
                toast('Ungültige E-Mail: ' + badMail, 'warning');
                return;
            }
            const cur = loadSettings();
            cur[listKey] = cleaned;
            ensureBoardNodes(cur);
            saveSettings(cur);
            toast(title + ' gespeichert (' + cleaned.length + ' Personen).', 'success');
            closeOverlay();
            notify();
        });

        renderBoard();
    }

    return { openBoardListEditor };
}
