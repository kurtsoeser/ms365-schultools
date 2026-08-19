/**
 * Modal-UI: Gruppenmitglieder mit Lizenzprüfung in Stammdaten übernehmen.
 * @file
 */
import {
    applyAdminImportSelection,
    applyMembershipImportSelection,
    buildMembershipImportPreview
} from './membership-reconcile.js';
import { summarizeUserLicenses, teacherEmailOfUser } from './graph-licenses.js';
import { escapeHtml, normStr } from './utils/strings.js';

let modalRoot = null;

function gug() {
    const G = window.ms365GraphUnifiedGroups;
    if (!G) throw new Error('graph-unified-groups.js fehlt.');
    return G;
}

function patchDirectoryMatches(matches) {
    if (!matches || !Object.keys(matches).length) return;
    try {
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: matches });
        }
    } catch {
        /* ignore */
    }
}

function ensureModal() {
    if (modalRoot) return modalRoot;
    modalRoot = document.createElement('div');
    modalRoot.id = 'ms365MembershipImportModal';
    modalRoot.className = 'modal-overlay slg-import-modal';
    modalRoot.hidden = true;
    modalRoot.setAttribute('role', 'dialog');
    modalRoot.setAttribute('aria-modal', 'true');
    modalRoot.innerHTML =
        '<div class="modal-box slg-import-modal__box" tabindex="-1">' +
        '<div class="slg-import-modal__head">' +
        '<h3 class="slg-import-modal__title" id="ms365MembershipImportTitle">Stammdaten übernehmen</h3>' +
        '<button type="button" class="btn btn-sm" id="ms365MembershipImportClose" title="Schließen"><i class="bi bi-x-lg"></i></button>' +
        '</div>' +
        '<p class="muted slg-import-modal__status" id="ms365MembershipImportStatus"></p>' +
        '<p class="slg-import-modal__warn" id="ms365MembershipImportWarn" hidden></p>' +
        '<div class="slg-import-modal__toolbar">' +
        '<label class="slg-deviation-section__select-all">' +
        '<input type="checkbox" id="ms365MembershipImportSelectAll" checked />' +
        '<span>Alle importierbaren auswählen</span>' +
        '</label>' +
        '</div>' +
        '<div class="slg-import-table-wrap">' +
        '<table class="slg-deviation-table slg-import-table" aria-label="Import-Vorschau">' +
        '<thead id="ms365MembershipImportHead"></thead>' +
        '<tbody id="ms365MembershipImportBody"></tbody>' +
        '</table>' +
        '</div>' +
        '<div class="modal-actions slg-import-modal__actions">' +
        '<button type="button" class="btn" id="ms365MembershipImportCancel">Abbrechen</button>' +
        '<button type="button" class="btn btn-success" id="ms365MembershipImportApply">In Stammdaten übernehmen</button>' +
        '</div></div>';
    document.body.appendChild(modalRoot);
    return modalRoot;
}

function kindSpec(kind) {
    if (kind === 'schueler') {
        return {
            label: 'Schüler:innen',
            colKey: 'klasse',
            colLabel: 'Klasse',
            existingKey: 'students'
        };
    }
    if (kind === 'verwaltung') {
        return {
            label: 'Verwaltung',
            colKey: 'role',
            colLabel: 'Rolle',
            existingKey: 'admin'
        };
    }
    return {
        label: 'Lehrer:innen',
        colKey: 'code',
        colLabel: 'Kürzel',
        existingKey: 'teachers'
    };
}

function buildAdminImportPreviewRows(users, existingAdmin, defaultRole, skuLookup) {
    const existing = Array.isArray(existingAdmin) ? existingAdmin : [];
    const emailToExisting = new Map();
    existing.forEach(function (r) {
        const em = String((r && r.email) || '')
            .trim()
            .toLowerCase();
        if (em) emailToExisting.set(em, r);
    });
    const rows = [];
    (Array.isArray(users) ? users : []).forEach(function (u) {
        if (!u || !u.id) return;
        const email = teacherEmailOfUser(u);
        if (!email) return;
        const existingRow = emailToExisting.get(email);
        const sum = summarizeUserLicenses(u, skuLookup);
        rows.push({
            graphUserId: String(u.id),
            displayName: normStr(u.displayName),
            userPrincipalName: normStr(u.userPrincipalName),
            email: email,
            role: existingRow && existingRow.role ? existingRow.role : defaultRole || '',
            name: normStr(u.displayName),
            licenseLabel: sum.primaryLabel || '–',
            licenseWarning: false,
            alreadyInList: !!existingRow,
            selected: !existingRow
        });
    });
    rows.sort(function (a, b) {
        return String(a.name || '').localeCompare(String(b.name || ''), 'de', { sensitivity: 'base' });
    });
    return rows;
}

/**
 * @param {object} cfg
 * @param {'lehrer'|'schueler'} cfg.kind
 * @param {string[]} cfg.emails
 * @param {() => Promise<string>} cfg.getGraphToken
 * @param {() => object|null} cfg.loadSettings
 * @param {(settings: object) => object} cfg.saveSettings
 * @param {(msg: string) => void} [cfg.toast]
 * @param {(msg: string, opts?: object) => Promise<boolean>} [cfg.dlgConfirm]
 * @param {() => void|Promise<void>} [cfg.onApplied]
 * @param {(entry: object) => void} [cfg.logAction]
 */
export async function openMembershipImportDialog(cfg) {
    const kindRaw = cfg && cfg.kind ? String(cfg.kind) : 'lehrer';
    const kind =
        kindRaw === 'schueler' ? 'schueler' : kindRaw === 'verwaltung' ? 'verwaltung' : 'lehrer';
    const importOptions = (cfg && cfg.importOptions) || {};
    const emails = Array.isArray(cfg && cfg.emails) ? cfg.emails : [];
    const spec = kindSpec(kind);
    const toast =
        (cfg && cfg.toast) ||
        function (msg) {
            if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(msg);
            else window.alert(msg);
        };
    const dlgConfirm =
        (cfg && cfg.dlgConfirm) ||
        function (msg, opts) {
            if (typeof window.ms365AppDialogConfirm === 'function') {
                return window.ms365AppDialogConfirm(msg, opts || {});
            }
            return Promise.resolve(window.confirm(msg));
        };

    if (!emails.length) {
        toast('Keine Personen für den Import ausgewählt.');
        return { cancelled: true, added: 0 };
    }
    if (typeof cfg.loadSettings !== 'function' || typeof cfg.saveSettings !== 'function') {
        throw new Error('Stammdaten-API fehlt.');
    }

    ensureModal();
    const titleEl = document.getElementById('ms365MembershipImportTitle');
    const statusEl = document.getElementById('ms365MembershipImportStatus');
    const warnEl = document.getElementById('ms365MembershipImportWarn');
    const headEl = document.getElementById('ms365MembershipImportHead');
    const bodyEl = document.getElementById('ms365MembershipImportBody');
    const applyBtn = document.getElementById('ms365MembershipImportApply');
    const selectAllEl = document.getElementById('ms365MembershipImportSelectAll');

    /** @type {object[]} */
    let previewRows = [];
    let resolver = null;

    function closeModal(result) {
        modalRoot.hidden = true;
        modalRoot.classList.remove('open');
        document.removeEventListener('keydown', onKey, true);
        const fn = resolver;
        resolver = null;
        if (fn) fn(result || { cancelled: true, added: 0 });
    }

    function onKey(ev) {
        if (ev.key === 'Escape') {
            ev.preventDefault();
            closeModal({ cancelled: true, added: 0 });
        }
    }

    function renderHead() {
        if (!headEl) return;
        headEl.replaceChildren();
        const tr = document.createElement('tr');
        ['', spec.colLabel, 'Name', 'E-Mail', 'Lizenz', ''].forEach(function (label) {
            const th = document.createElement('th');
            th.className = label === '' ? 'col-check' : '';
            th.textContent = label;
            tr.appendChild(th);
        });
        headEl.appendChild(tr);
    }

    function updateStatus() {
        if (!statusEl) return;
        const selected = previewRows.filter(function (r) {
            return r.selected;
        }).length;
        const licensed = previewRows.filter(function (r) {
            return !r.licenseWarning;
        }).length;
        const warned = previewRows.filter(function (r) {
            return r.licenseWarning;
        }).length;
        statusEl.textContent =
            previewRows.length +
            ' Personen · ' +
            licensed +
            ' mit passender Education-Lizenz · ' +
            selected +
            ' ausgewählt zum Übernehmen.';
        if (warnEl) {
            if (warned) {
                warnEl.hidden = false;
                warnEl.innerHTML =
                    '<i class="bi bi-exclamation-triangle" aria-hidden="true"></i> ' +
                    warned +
                    ' ohne passende A1/A3/A5-Lizenz – standardmäßig abgewählt. Nur bewusst übernehmen (z. B. Service-Konten).';
            } else {
                warnEl.hidden = true;
                warnEl.textContent = '';
            }
        }
        if (applyBtn) applyBtn.disabled = selected === 0;
        if (selectAllEl) {
            const importable = previewRows.filter(function (r) {
                return !!r.email && !r.licenseWarning;
            });
            const sel = importable.filter(function (r) {
                return r.selected;
            });
            selectAllEl.checked = importable.length > 0 && sel.length === importable.length;
            selectAllEl.indeterminate = sel.length > 0 && sel.length < importable.length;
        }
    }

    function renderBody() {
        if (!bodyEl) return;
        bodyEl.replaceChildren();
        if (!previewRows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 6;
            td.className = 'muted';
            td.textContent = 'Keine importierbaren Personen gefunden.';
            tr.appendChild(td);
            bodyEl.appendChild(tr);
            updateStatus();
            return;
        }
        previewRows.forEach(function (row) {
            const tr = document.createElement('tr');
            if (row.licenseWarning) tr.classList.add('slg-import-row--warn');
            if (row.alreadyInList) tr.classList.add('slg-import-row--exists');

            const tdChk = document.createElement('td');
            tdChk.className = 'col-check';
            const chk = document.createElement('input');
            chk.type = 'checkbox';
            chk.checked = !!row.selected;
            chk.disabled = !row.email;
            chk.addEventListener('change', function () {
                row.selected = chk.checked;
                updateStatus();
            });
            tdChk.appendChild(chk);

            const tdEdit = document.createElement('td');
            const editInp = document.createElement('input');
            editInp.type = 'text';
            editInp.className = 'cell-editor';
            editInp.value = row[spec.colKey] || '';
            editInp.maxLength = spec.colKey === 'klasse' ? 16 : 12;
            editInp.setAttribute('aria-label', spec.colLabel);
            editInp.addEventListener('input', function () {
                row[spec.colKey] = editInp.value;
            });
            tdEdit.appendChild(editInp);

            const tdName = document.createElement('td');
            const nameInp = document.createElement('input');
            nameInp.type = 'text';
            nameInp.className = 'cell-editor';
            nameInp.value = row.name || '';
            nameInp.setAttribute('aria-label', 'Name');
            nameInp.addEventListener('input', function () {
                row.name = nameInp.value;
            });
            tdName.appendChild(nameInp);

            const tdEmail = document.createElement('td');
            tdEmail.textContent = row.email || '–';

            const tdLic = document.createElement('td');
            tdLic.innerHTML =
                '<span class="sw-lic-pill' +
                (row.licenseWarning ? ' sw-lic-pill--warn' : '') +
                '">' +
                escapeHtml(row.licenseLabel || '–') +
                '</span>';

            const tdHint = document.createElement('td');
            if (row.licenseWarning) {
                tdHint.className = 'slg-import-hint';
                tdHint.textContent = row.warningText || 'Prüfen';
            } else if (row.alreadyInList) {
                tdHint.className = 'muted';
                tdHint.textContent = 'Update';
            } else {
                tdHint.className = 'muted';
                tdHint.textContent = 'Neu';
            }

            tr.appendChild(tdChk);
            tr.appendChild(tdEdit);
            tr.appendChild(tdName);
            tr.appendChild(tdEmail);
            tr.appendChild(tdLic);
            tr.appendChild(tdHint);
            bodyEl.appendChild(tr);
        });
        updateStatus();
    }

    function bindOnce() {
        const closeBtn = document.getElementById('ms365MembershipImportClose');
        const cancelBtn = document.getElementById('ms365MembershipImportCancel');
        if (closeBtn && !closeBtn.dataset.bound) {
            closeBtn.dataset.bound = '1';
            closeBtn.addEventListener('click', function () {
                closeModal({ cancelled: true, added: 0 });
            });
        }
        if (cancelBtn && !cancelBtn.dataset.bound) {
            cancelBtn.dataset.bound = '1';
            cancelBtn.addEventListener('click', function () {
                closeModal({ cancelled: true, added: 0 });
            });
        }
        if (modalRoot && !modalRoot.dataset.bound) {
            modalRoot.dataset.bound = '1';
            modalRoot.addEventListener('click', function (ev) {
                if (ev.target === modalRoot) closeModal({ cancelled: true, added: 0 });
            });
        }
        if (selectAllEl && !selectAllEl.dataset.bound) {
            selectAllEl.dataset.bound = '1';
            selectAllEl.addEventListener('change', function () {
                const on = !!selectAllEl.checked;
                previewRows.forEach(function (r) {
                    if (r.email && !r.licenseWarning) r.selected = on;
                });
                renderBody();
            });
        }
        if (applyBtn && !applyBtn.dataset.bound) {
            applyBtn.dataset.bound = '1';
            applyBtn.addEventListener('click', function () {
                void applyImport();
            });
        }
    }

    async function applyImport() {
        const selected = previewRows.filter(function (r) {
            return r.selected;
        });
        if (!selected.length) {
            toast('Keine Zeilen ausgewählt.');
            return;
        }
        const warned = selected.filter(function (r) {
            return r.licenseWarning;
        });
        if (warned.length) {
            const ok = await dlgConfirm(
                warned.length +
                    (warned.length === 1 ? ' Person hat' : ' Personen haben') +
                    ' keine passende Education-Lizenz. Trotzdem in die Stammdaten übernehmen?',
                { title: 'Lizenz-Hinweis', danger: true }
            );
            if (!ok) return;
        }
        const ok = await dlgConfirm(
            selected.length +
                (selected.length === 1 ? ' Person' : ' Personen') +
                ' in die lokalen Stammdaten übernehmen?',
            { title: 'Import bestätigen' }
        );
        if (!ok) return;

        const settings = cfg.loadSettings();
        if (!settings) {
            toast('Stammdaten konnten nicht geladen werden.');
            return;
        }
        const existing = Array.isArray(settings[spec.existingKey]) ? settings[spec.existingKey] : [];
        let result;
        const next = Object.assign({}, settings);
        if (kind === 'verwaltung') {
            result = applyAdminImportSelection(existing, previewRows);
            next.admin = result.admin || existing;
        } else {
            result = applyMembershipImportSelection(kind, existing, previewRows);
            if (kind === 'lehrer') next.teachers = result.teachers || existing;
            else next.students = result.students || existing;
        }
        cfg.saveSettings(next);
        patchDirectoryMatches(result.directoryMatches);
        if (typeof cfg.logAction === 'function') {
            cfg.logAction({
                tool: 'slg',
                action: 'membership-import-local',
                target: kind,
                summary:
                    spec.label +
                    ': ' +
                    (result.added ? result.added.length : 0) +
                    ' neu, ' +
                    (result.updated ? result.updated.length : 0) +
                    ' aktualisiert'
            });
        }
        const added = result.added ? result.added.length : 0;
        toast(added + ' neu in Stammdaten übernommen.');
        if (typeof cfg.onApplied === 'function') {
            await cfg.onApplied({ added: added, updated: result.updated ? result.updated.length : 0 });
        }
        closeModal({ cancelled: false, added: added, updated: result.updated ? result.updated.length : 0 });
    }

    bindOnce();
    if (titleEl) titleEl.textContent = spec.label + ' – Stammdaten übernehmen';
    if (statusEl) statusEl.textContent = 'Lade Personen und Lizenzen aus Microsoft 365 …';
    if (warnEl) warnEl.hidden = true;
    renderHead();
    if (bodyEl) bodyEl.replaceChildren();
    if (applyBtn) applyBtn.disabled = true;
    modalRoot.hidden = false;
    modalRoot.classList.add('open');
    document.addEventListener('keydown', onKey, true);

    try {
        const G = gug();
        const token = await cfg.getGraphToken();
        let skuLookup = null;
        if (typeof G.fetchSubscribedSkus === 'function') {
            const sub = await G.fetchSubscribedSkus(token);
            if (sub && sub.ok && typeof G.skuLookupFromSubscribed === 'function') {
                skuLookup = G.skuLookupFromSubscribed(sub.skus);
            }
        }
        const users =
            typeof G.resolveUsersByEmailsForImport === 'function'
                ? await G.resolveUsersByEmailsForImport(token, emails)
                : [];
        const settings = cfg.loadSettings();
        const existing = settings && Array.isArray(settings[spec.existingKey]) ? settings[spec.existingKey] : [];
        if (kind === 'verwaltung') {
            let defaultRole = String(importOptions.defaultRole || '').trim();
            if (!defaultRole && typeof window.ms365AppDialogPrompt === 'function') {
                const v = await window.ms365AppDialogPrompt(
                    'Standard-Rolle für neue Verwaltungskontakte (z. B. Sekretariat).',
                    'Sekretariat',
                    { title: 'Rolle für Import', inputLabel: 'Rolle', okText: 'Weiter', cancelText: 'Abbrechen' }
                );
                if (v === null) {
                    closeModal({ cancelled: true, added: 0 });
                    return new Promise(function (resolve) {
                        resolver = resolve;
                    });
                }
                defaultRole = String(v || '').trim();
            }
            previewRows = buildAdminImportPreviewRows(users, existing, defaultRole, skuLookup);
        } else {
            previewRows = buildMembershipImportPreview(kind, users, existing, skuLookup, {
                activeOnly: true,
                families: ['a1', 'a3', 'a5']
            });
            if (kind === 'schueler' && importOptions.defaultClass) {
                const dc = String(importOptions.defaultClass).trim();
                previewRows.forEach(function (row) {
                    if (row && !String(row.klasse || '').trim()) row.klasse = dc;
                });
            }
        }
        renderBody();
    } catch (e) {
        if (statusEl) statusEl.textContent = '';
        if (bodyEl) {
            bodyEl.replaceChildren();
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 6;
            td.className = 'slg-deviation-panel__error';
            td.textContent = 'Import-Vorschau fehlgeschlagen: ' + (e.message || e);
            tr.appendChild(td);
            bodyEl.appendChild(tr);
        }
    }

    return new Promise(function (resolve) {
        resolver = resolve;
    });
}

const api = { openMembershipImportDialog: openMembershipImportDialog };

if (typeof window !== 'undefined') {
    window.ms365MembershipImportUi = api;
}

export default api;
