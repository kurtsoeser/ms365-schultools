/**
 * Wizard-UI für Kalender-Postfächer.
 */

import { getEl } from '../../shared/utils/dom.js';
import { compareDe } from '../../shared/utils/strings.js';
import { fetchAllPages, getGraphToken } from '../../shared/graph-client.js';
import { pickEntraUsers } from '../../shared/entra-user-picker.js';
import {
    CALENDAR_PRESETS,
    buildCalendarFromPreset,
    looksLikeCalendarMailbox
} from './kalender-postfaecher-presets.js';
import {
    buildCalendarPostfaecherCmd,
    buildCalendarPostfaecherScript
} from './kalender-postfaecher-script.js';
import { loadState, saveState } from './kalender-postfaecher-state.js';
import { classifyDirectorySharedMailboxCandidate } from '../postfaecher/postfaecher-load.js';

const GRAPH_SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All'
];

const USER_SELECT =
    'id,displayName,mail,userPrincipalName,mailNickname,accountEnabled,userType,assignedLicenses,jobTitle,department';

/**
 * Schuildomain aus Stammdaten / school-domain.js (ohne manuellen Override).
 * @returns {string}
 */
export function resolveSchoolDomain() {
    try {
        if (typeof window !== 'undefined' && typeof window.ms365GetSchoolDomainNoAt === 'function') {
            const d = String(window.ms365GetSchoolDomainNoAt() || '')
                .trim()
                .replace(/^@+/, '');
            if (d) return d;
        }
    } catch {
        /* ignore */
    }
    try {
        const api = typeof window !== 'undefined' ? window.ms365AppDataV2 : null;
        if (api && typeof api.getContainer === 'function') {
            const c = api.getContainer();
            const d = String((c && c.core && c.core.domain) || '')
                .trim()
                .replace(/^@+/, '');
            if (d) return d;
        }
    } catch {
        /* ignore */
    }
    try {
        const raw = typeof localStorage !== 'undefined' ? localStorage.getItem('ms365-tenant-settings-v1') : null;
        if (raw) {
            const o = JSON.parse(raw);
            const d = String((o && o.domain) || '')
                .trim()
                .replace(/^@+/, '');
            if (d) return d;
        }
    } catch {
        /* ignore */
    }
    return '';
}

function esc(s) {
    return String(s == null ? '' : s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

function linesToList(text) {
    return String(text || '')
        .split(/\r?\n/)
        .map((l) => l.trim())
        .filter(Boolean);
}

function listToLines(arr) {
    return (Array.isArray(arr) ? arr : []).join('\n');
}

/**
 * Ausgewählte Entra-Personen in eine UPN/E-Mail-Liste mergen (ohne Duplikate).
 * @param {string[]} existing
 * @param {Array<{ mail?: string }>} picked
 * @returns {string[]}
 */
export function mergePrincipalMails(existing, picked) {
    const out = [];
    const seen = new Set();
    function add(raw) {
        const mail = String(raw || '').trim();
        if (!mail || mail.indexOf('@') === -1) return;
        const key = mail.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(mail);
    }
    (Array.isArray(existing) ? existing : []).forEach(add);
    (Array.isArray(picked) ? picked : []).forEach(function (p) {
        add(p && p.mail);
    });
    return out;
}

function downloadText(filename, text, mime) {
    const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1500);
}

async function copyText(text) {
    try {
        await navigator.clipboard.writeText(text);
        return true;
    } catch {
        return false;
    }
}

/**
 * @param {{ getToken?: (scopes: string[]) => Promise<string> }} [deps]
 */
export function createKalenderPostfaecherUi(deps) {
    const state = loadState();
    state.domain = resolveSchoolDomain();
    let selectedId = state.items[0] ? state.items[0].id : '';
    let statusMsg = '';

    function persist() {
        saveState(state);
    }

    function setStatus(msg) {
        statusMsg = String(msg || '');
        const el = getEl('kpStatus');
        if (el) el.textContent = statusMsg;
    }

    function selectedItem() {
        return state.items.find((it) => it.id === selectedId) || null;
    }

    function renderPresetButtons() {
        const host = getEl('kpPresets');
        if (!host) return;
        host.innerHTML = CALENDAR_PRESETS.map(function (p) {
            return (
                '<button type="button" class="btn kp-preset" data-preset="' +
                esc(p.id) +
                '"><i class="bi bi-plus-lg"></i>' +
                esc(p.label) +
                '</button>'
            );
        }).join('');
        host.querySelectorAll('[data-preset]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const domain = resolveSchoolDomain();
                state.domain = domain;
                const item = buildCalendarFromPreset(btn.getAttribute('data-preset'), domain);
                state.items.push(item);
                selectedId = item.id;
                persist();
                render();
            });
        });
    }

    function renderInventory() {
        const host = getEl('kpInventory');
        if (!host) return;
        const rows = Array.isArray(state.inventory) ? state.inventory.slice() : [];
        rows.sort((a, b) => compareDe(a.name || a.upn || '', b.name || b.upn || ''));
        if (!rows.length) {
            host.innerHTML = '<p class="muted">Noch kein Inventar – „Aus Microsoft 365 laden“.</p>';
            return;
        }
        const calRows = rows.filter(looksLikeCalendarMailbox);
        const show = calRows.length ? calRows : rows;
        host.innerHTML =
            '<p class="muted">' +
            (calRows.length
                ? calRows.length + ' mögliche Kalender-Postfächer (Heuristik Name/Alias).'
                : rows.length + ' Shared-Mailbox-Kandidaten (keine Kalender-Namen erkannt).') +
            '</p><ul class="kp-inv-list">' +
            show
                .map(function (r) {
                    return (
                        '<li><button type="button" class="kp-inv-item" data-upn="' +
                        esc(r.upn || r.mail || '') +
                        '" data-name="' +
                        esc(r.name || '') +
                        '" data-alias="' +
                        esc(r.alias || '') +
                        '"><strong>' +
                        esc(r.name || r.upn || '—') +
                        '</strong><span class="muted">' +
                        esc(r.upn || r.mail || '') +
                        '</span></button></li>'
                    );
                })
                .join('') +
            '</ul>';
        host.querySelectorAll('.kp-inv-item').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const upn = btn.getAttribute('data-upn') || '';
                const name = btn.getAttribute('data-name') || upn;
                const alias =
                    btn.getAttribute('data-alias') ||
                    (upn.indexOf('@') !== -1 ? upn.split('@')[0] : '');
                const domain = resolveSchoolDomain();
                state.domain = domain;
                const item = buildCalendarFromPreset('custom', domain, { displayName: name });
                item.displayName = name;
                item.alias = alias || item.alias;
                item.primarySmtp = upn || item.primarySmtp;
                item.mode = 'permissionsOnly';
                item.publish = false;
                state.items.push(item);
                selectedId = item.id;
                persist();
                render();
            });
        });
    }

    function renderList() {
        const host = getEl('kpList');
        if (!host) return;
        if (!state.items.length) {
            host.innerHTML = '<p class="muted">Noch keine Kalender geplant – Vorlage wählen oder aus Inventar übernehmen.</p>';
            return;
        }
        host.innerHTML =
            '<ul class="kp-plan-list">' +
            state.items
                .map(function (it) {
                    const cur = it.id === selectedId ? ' aria-current="true"' : '';
                    return (
                        '<li><button type="button" class="kp-plan-item"' +
                        cur +
                        ' data-id="' +
                        esc(it.id) +
                        '"><strong>' +
                        esc(it.displayName) +
                        '</strong><span class="muted">' +
                        esc(it.primarySmtp || it.alias) +
                        ' · ' +
                        (it.mode === 'permissionsOnly' ? 'Rechte/Publish' : 'Anlegen') +
                        '</span></button></li>'
                    );
                })
                .join('') +
            '</ul>';
        host.querySelectorAll('.kp-plan-item').forEach(function (btn) {
            btn.addEventListener('click', function () {
                selectedId = btn.getAttribute('data-id') || '';
                renderEditor();
                renderList();
            });
        });
    }

    function readEditorIntoSelected() {
        const it = selectedItem();
        if (!it) return;
        const dn = getEl('kpDisplayName');
        const al = getEl('kpAlias');
        const smtp = getEl('kpSmtp');
        const desc = getEl('kpDesc');
        const mode = getEl('kpMode');
        const editors = getEl('kpEditors');
        const readers = getEl('kpReaders');
        const sendAs = getEl('kpSendAs');
        const publish = getEl('kpPublish');
        const note = getEl('kpPublishNote');
        if (dn) it.displayName = dn.value.trim();
        if (al) it.alias = al.value.trim();
        if (smtp) it.primarySmtp = smtp.value.trim();
        if (desc) it.description = desc.value.trim();
        if (mode) it.mode = mode.value === 'permissionsOnly' ? 'permissionsOnly' : 'create';
        if (editors) it.editors = linesToList(editors.value);
        if (readers) it.readers = linesToList(readers.value);
        if (sendAs) it.sendAs = !!sendAs.checked;
        if (publish) it.publish = !!publish.checked;
        if (note) it.publishUrlNote = note.value.trim();
        persist();
    }

    function renderEditor() {
        const host = getEl('kpEditor');
        if (!host) return;
        const it = selectedItem();
        if (!it) {
            host.innerHTML = '<p class="muted">Kalender in der Liste wählen.</p>';
            return;
        }
        host.innerHTML =
            '<div class="kp-form">' +
            '<div class="field"><label for="kpDisplayName">Anzeigename</label><input id="kpDisplayName" value="' +
            esc(it.displayName) +
            '"></div>' +
            '<div class="field"><label for="kpAlias">Alias</label><input id="kpAlias" value="' +
            esc(it.alias) +
            '"></div>' +
            '<div class="field"><label for="kpSmtp">Primäre SMTP</label><input id="kpSmtp" value="' +
            esc(it.primarySmtp) +
            '" placeholder="schulkalender@' +
            esc(resolveSchoolDomain() || 'ihre-schule.at') +
            '"></div>' +
            '<div class="field"><label for="kpMode">Modus</label><select id="kpMode">' +
            '<option value="create"' +
            (it.mode !== 'permissionsOnly' ? ' selected' : '') +
            '>Neu anlegen (Shared Mailbox)</option>' +
            '<option value="permissionsOnly"' +
            (it.mode === 'permissionsOnly' ? ' selected' : '') +
            '>Nur Rechte / Publish nachziehen</option>' +
            '</select></div>' +
            '<div class="field kp-full"><label for="kpDesc">Beschreibung</label><textarea id="kpDesc" rows="2">' +
            esc(it.description) +
            '</textarea></div>' +
            '<div class="field kp-full"><label for="kpEditors">Eintragen (Benutzer, eine UPN/E-Mail pro Zeile)</label>' +
            '<div class="kp-pick-row">' +
            '<button type="button" class="btn" id="kpPickEditors"><i class="bi bi-person-bounding-box"></i>Aus Microsoft 365 wählen</button>' +
            '</div>' +
            '<textarea id="kpEditors" rows="4" placeholder="person@' +
            esc(resolveSchoolDomain() || 'schule.at') +
            '">' +
            esc(listToLines(it.editors)) +
            '</textarea>' +
            '<p class="muted">FullAccess + Kalender-Editor' +
            (it.sendAs !== false ? ' + SendAs' : '') +
            '. M365-Gruppen gehen hier nicht – bitte Benutzer wählen.</p></div>' +
            '<div class="field kp-full"><label for="kpReaders">Nur lesen (Benutzer, eine pro Zeile)</label>' +
            '<div class="kp-pick-row">' +
            '<button type="button" class="btn" id="kpPickReaders"><i class="bi bi-person-bounding-box"></i>Aus Microsoft 365 wählen</button>' +
            '</div>' +
            '<textarea id="kpReaders" rows="3" placeholder="kollege@' +
            esc(resolveSchoolDomain() || 'schule.at') +
            '">' +
            esc(listToLines(it.readers)) +
            '</textarea>' +
            '<p class="muted">Kalender-Ordner: Reviewer (ohne FullAccess). Ebenfalls Benutzer, keine M365-Gruppe.</p></div>' +
            '<label class="kp-check"><input type="checkbox" id="kpSendAs"' +
            (it.sendAs !== false ? ' checked' : '') +
            '> SendAs für Eintragen-Personen</label>' +
            '<label class="kp-check"><input type="checkbox" id="kpPublish"' +
            (it.publish ? ' checked' : '') +
            '> Internet-Kalender veröffentlichen (iCal-URL)</label>' +
            '<div class="field kp-full"><label for="kpPublishNote">Publish-URL (nach Skriptlauf hier einfügen)</label>' +
            '<input id="kpPublishNote" value="' +
            esc(it.publishUrlNote || '') +
            '" placeholder="https://…outlook.office365.com/owa/calendar/…"></div>' +
            '<div class="kp-editor-actions">' +
            '<button type="button" class="btn" id="kpSaveItem"><i class="bi bi-check2"></i>Übernehmen</button>' +
            '<button type="button" class="btn btn-danger" id="kpRemoveItem"><i class="bi bi-trash"></i>Entfernen</button>' +
            '</div></div>';

        const bind = () => {
            readEditorIntoSelected();
            renderScript();
            renderList();
        };
        ['kpDisplayName', 'kpAlias', 'kpSmtp', 'kpDesc', 'kpMode', 'kpEditors', 'kpReaders', 'kpPublishNote'].forEach(
            function (id) {
                const el = getEl(id);
                if (el) el.addEventListener('change', bind);
            }
        );
        ['kpSendAs', 'kpPublish'].forEach(function (id) {
            const el = getEl(id);
            if (el) el.addEventListener('change', bind);
        });
        const saveBtn = getEl('kpSaveItem');
        if (saveBtn) {
            saveBtn.addEventListener('click', function () {
                bind();
                setStatus('Kalender übernommen.');
            });
        }
        const rm = getEl('kpRemoveItem');
        if (rm) {
            rm.addEventListener('click', function () {
                state.items = state.items.filter((x) => x.id !== it.id);
                selectedId = state.items[0] ? state.items[0].id : '';
                persist();
                render();
            });
        }

        async function pickInto(role) {
            const current = selectedItem();
            if (!current) return;
            readEditorIntoSelected();
            const isEditors = role === 'editors';
            try {
                const picked = await pickEntraUsers({
                    title: isEditors
                        ? 'Personen mit Eintragen-Recht wählen'
                        : 'Personen mit Leserecht wählen',
                    hint:
                        'Benutzer aus Microsoft 365 suchen und anhaken. ' +
                        'M365-Gruppen sind für Postfach-/Kalenderrechte ungeeignet – bitte Personen wählen.'
                });
                if (!picked || !picked.length) return;
                if (isEditors) {
                    current.editors = mergePrincipalMails(current.editors, picked);
                } else {
                    current.readers = mergePrincipalMails(current.readers, picked);
                }
                persist();
                renderEditor();
                renderScript();
                renderList();
                setStatus(
                    picked.length +
                        (isEditors ? ' Person(en) zu Eintragen' : ' Person(en) zu Nur lesen') +
                        ' hinzugefügt.'
                );
            } catch (e) {
                setStatus('Picker: ' + (e && e.message ? e.message : String(e)));
            }
        }

        const pickEd = getEl('kpPickEditors');
        if (pickEd) pickEd.addEventListener('click', () => void pickInto('editors'));
        const pickRd = getEl('kpPickReaders');
        if (pickRd) pickRd.addEventListener('click', () => void pickInto('readers'));
    }

    function renderScript() {
        const ta = getEl('kpScript');
        if (!ta) return;
        readEditorIntoSelected();
        if (!state.items.length) {
            ta.value = '# Noch keine Kalender geplant.';
            return;
        }
        ta.value = buildCalendarPostfaecherScript(state.items, {
            suggestedAdminUpn: state.suggestedAdminUpn
        });
    }

    function renderMeta() {
        state.domain = resolveSchoolDomain();
        const domainEl = getEl('kpDomainDisplay');
        if (domainEl) {
            if (state.domain) {
                domainEl.innerHTML =
                    '<strong>' +
                    esc(state.domain) +
                    '</strong> <span class="muted">· aus Stammdaten</span>';
            } else {
                domainEl.innerHTML =
                    '<span class="muted">Noch nicht gesetzt – bitte in den </span>' +
                    '<a href="../tenant.html">Stammdaten</a>' +
                    '<span class="muted"> die E-Mail-Domain der Schule eintragen.</span>';
            }
        }
        const admin = getEl('kpAdmin');
        if (admin && admin.value !== state.suggestedAdminUpn) admin.value = state.suggestedAdminUpn;
    }

    function render() {
        renderMeta();
        renderPresetButtons();
        renderInventory();
        renderList();
        renderEditor();
        renderScript();
        setStatus(statusMsg || (state.items.length ? state.items.length + ' Kalender geplant.' : 'Bereit.'));
    }

    /**
     * Wie Postfächer: Advanced Query mit ConsistencyLevel, sonst einfacher Fallback.
     * @param {string} token
     * @param {string} filterExpr
     */
    async function fetchUsersAdvanced(token, filterExpr) {
        const path =
            '/users?$count=true&$filter=' +
            encodeURIComponent(filterExpr) +
            '&$select=' +
            encodeURIComponent(USER_SELECT) +
            '&$top=999';
        const res = await fetchAllPages(token, path, {
            maxPages: 20,
            extraHeaders: { ConsistencyLevel: 'eventual' }
        });
        return res.items || [];
    }

    async function loadInventoryFromGraph() {
        setStatus('Lade Shared-Mailbox-Kandidaten …');
        try {
            state.domain = resolveSchoolDomain();
            let token;
            if (deps && typeof deps.getToken === 'function') {
                token = await deps.getToken(GRAPH_SCOPES);
            } else {
                try {
                    token = await getGraphToken(GRAPH_SCOPES);
                } catch {
                    if (typeof window.ms365AuthAcquireToken === 'function') {
                        token = await window.ms365AuthAcquireToken(GRAPH_SCOPES);
                    } else {
                        throw new Error('Bitte oben rechts anmelden.');
                    }
                }
            }

            /** @type {object[]} */
            const buckets = [];
            /** @type {string[]} */
            const errors = [];

            try {
                setStatus('Verzeichnis: deaktivierte Konten …');
                buckets.push(
                    ...(await fetchUsersAdvanced(token, 'accountEnabled eq false and mail ne null'))
                );
            } catch (e) {
                errors.push('Pass1: ' + (e && e.message ? e.message : String(e)));
            }

            if (!buckets.length) {
                try {
                    setStatus('Fallback: deaktivierte Konten (einfache Abfrage) …');
                    const simplePath =
                        '/users?$filter=' +
                        encodeURIComponent('accountEnabled eq false') +
                        '&$select=' +
                        encodeURIComponent(USER_SELECT) +
                        '&$top=999';
                    const res = await fetchAllPages(token, simplePath, { maxPages: 20 });
                    (res.items || []).forEach(function (u) {
                        if (u && (u.mail || u.userPrincipalName)) buckets.push(u);
                    });
                } catch (e) {
                    errors.push('Fallback: ' + (e && e.message ? e.message : String(e)));
                }
            }

            if (!buckets.length && errors.length) {
                throw new Error(errors.join(' | '));
            }

            const byId = new Map();
            buckets.forEach(function (u) {
                const id = String((u && u.id) || '');
                if (id) byId.set(id, u);
            });

            const rows = [];
            byId.forEach(function (u) {
                const row = classifyDirectorySharedMailboxCandidate(u);
                if (!row) return;
                rows.push({
                    id: row.id || u.id || '',
                    name: row.name || u.displayName || '',
                    mail: row.mail || u.mail || '',
                    upn: row.upn || u.userPrincipalName || '',
                    alias: row.alias || u.mailNickname || '',
                    highConfidence: !!row.highConfidence
                });
            });
            state.inventory = rows;
            persist();
            setStatus(rows.length + ' Kandidaten geladen.');
            render();
        } catch (e) {
            setStatus('Fehler: ' + (e && e.message ? e.message : String(e)));
        }
    }

    function bindChrome() {
        const admin = getEl('kpAdmin');
        if (admin) {
            admin.addEventListener('change', function () {
                state.suggestedAdminUpn = admin.value.trim();
                persist();
            });
        }
        const loadBtn = getEl('kpLoadInventory');
        if (loadBtn) loadBtn.addEventListener('click', () => loadInventoryFromGraph());
        const copyBtn = getEl('kpCopyScript');
        if (copyBtn) {
            copyBtn.addEventListener('click', async function () {
                renderScript();
                const ta = getEl('kpScript');
                const ok = await copyText(ta ? ta.value : '');
                setStatus(ok ? 'Skript kopiert.' : 'Kopieren fehlgeschlagen – manuell markieren.');
            });
        }
        const dlPs = getEl('kpDownloadPs1');
        if (dlPs) {
            dlPs.addEventListener('click', function () {
                renderScript();
                const ta = getEl('kpScript');
                downloadText('kalender-postfaecher.ps1', ta ? ta.value : '', 'text/plain;charset=utf-8');
                setStatus('PS1 heruntergeladen.');
            });
        }
        const dlCmd = getEl('kpDownloadCmd');
        if (dlCmd) {
            dlCmd.addEventListener('click', function () {
                readEditorIntoSelected();
                const cmd = buildCalendarPostfaecherCmd(state.items, {
                    suggestedAdminUpn: state.suggestedAdminUpn
                });
                downloadText('kalender-postfaecher.cmd', cmd, 'application/octet-stream');
                setStatus('CMD heruntergeladen – doppelklicken und als Exchange-Admin anmelden.');
            });
        }
        const clearBtn = getEl('kpClearPlan');
        if (clearBtn) {
            clearBtn.addEventListener('click', function () {
                if (!state.items.length) return;
                if (!window.confirm('Geplante Kalender wirklich leeren?')) return;
                state.items = [];
                selectedId = '';
                persist();
                render();
            });
        }
    }

    bindChrome();
    render();

    return {
        render,
        getState: () => state,
        loadInventoryFromGraph
    };
}
