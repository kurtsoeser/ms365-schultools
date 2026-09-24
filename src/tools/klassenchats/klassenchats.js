/**
 * KlassenChats-Wizard: Gruppenchats für Lehrkräfte einer Klasse.
 */
import {
    calcYearPrefix,
    defaultChatNamePattern,
    normalizeChatNamePattern,
    chatTokenLabel,
    buildChatTopicFromPattern,
    buildKvByKlasse,
    buildClassChatPlans,
    buildManualClassChatPlan,
    summarizePlans,
    existingByKlasseFromState,
    upsertClassChatItem,
    normalizeClassChatsState
} from './klassenchats-logic.js';
import { provisionClassChat } from './klassenchats-graph.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}

function escapeHtml(s) {
    return String(s || '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
}

function G() {
    const api = window.ms365GraphUnifiedGroups;
    if (!api) throw new Error('Graph nicht geladen.');
    return api;
}

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Chat.Create',
    'https://graph.microsoft.com/Chat.ReadWrite'
];

async function getToken() {
    if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
        return window.ms365AuthAcquireTokenPopup(SCOPES);
    }
    return G().getGraphToken();
}

/** @type {ReturnType<typeof defaultChatNamePattern>} */
let namePattern = defaultChatNamePattern();
let step = 1;
/** @type {'belegung'|'manual'} */
let mode = 'belegung';
/** @type {ReturnType<typeof buildClassChatPlans>} */
let plans = [];
/** Manuelle Warteschlange (mehrere Chats vor dem Anlegen). */
/** @type {Array<{ klasse: string, memberEmails: string[], includeKv: boolean }>} */
let manualQueue = [];
/** Aus Tenant-Suche übernommene Personen: email → { email, label, id } */
/** @type {Map<string, { email: string, label: string, id: string }>} */
let tenantPicked = new Map();
/** @type {object|null} */
let belegung = null;
let yearPrefix = calcYearPrefix();
let busy = false;

function appData() {
    return window.ms365AppDataV2 || null;
}

function loadTenantClasses() {
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            return (s && Array.isArray(s.classes) ? s.classes : []) || [];
        }
    } catch {
        /* ignore */
    }
    try {
        const api = appData();
        if (!api || typeof api.getYearBucket !== 'function') return [];
        const { bucket } = api.getYearBucket();
        return (bucket && Array.isArray(bucket.classes) ? bucket.classes : []) || [];
    } catch {
        return [];
    }
}

function loadTenantTeachers() {
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') {
            const s = window.ms365TenantSettingsLoad();
            return (s && Array.isArray(s.teachers) ? s.teachers : []) || [];
        }
    } catch {
        /* ignore */
    }
    try {
        const api = appData();
        if (!api || typeof api.getContainer !== 'function') return [];
        const c = api.getContainer();
        return (c && c.core && Array.isArray(c.core.teachers) ? c.core.teachers : []) || [];
    } catch {
        return [];
    }
}

function loadBelegung() {
    const api = appData();
    if (!api || typeof api.getUnterrichtsbelegung !== 'function') return null;
    return api.getUnterrichtsbelegung();
}

function loadClassChatsState() {
    const api = appData();
    if (!api || typeof api.getClassChats !== 'function') return null;
    return api.getClassChats();
}

function saveClassChatsState(state) {
    const api = appData();
    if (!api || typeof api.setClassChats !== 'function') return null;
    return api.setClassChats(state);
}

function persistPatternMeta(state) {
    const base = normalizeClassChatsState(state) || {
        updatedAt: new Date().toISOString(),
        yearPrefix,
        namePattern,
        items: []
    };
    return saveClassChatsState({
        ...base,
        yearPrefix,
        namePattern,
        updatedAt: new Date().toISOString()
    });
}

function setMode(next) {
    mode = next === 'manual' ? 'manual' : 'belegung';
    document.querySelectorAll('[data-kc-mode]').forEach((btn) => {
        btn.classList.toggle('is-selected', btn.getAttribute('data-kc-mode') === mode);
    });
    const belegungPanel = $('kcPanelBelegung');
    const manualPanel = $('kcPanelManual');
    if (belegungPanel) belegungPanel.hidden = mode !== 'belegung';
    if (manualPanel) manualPanel.hidden = mode !== 'manual';
    updateManualQueueHint();
    if (mode === 'manual') {
        updateManualKvHint();
        updateNamePreview();
    }
}

function setStep(n) {
    step = Math.max(1, Math.min(4, Number(n) || 1));
    document.querySelectorAll('[data-kc-goto]').forEach((btn) => {
        btn.classList.toggle('is-active', Number(btn.getAttribute('data-kc-goto')) === step);
    });
    document.querySelectorAll('[data-kc-step]').forEach((el) => {
        el.classList.toggle('is-active', Number(el.getAttribute('data-kc-step')) === step);
    });
    if (step === 3) renderPreview();
}

function refreshClassDatalist() {
    const list = $('kcClassList');
    if (!list) return;
    const classes = loadTenantClasses()
        .map((c) => String(c.code || '').trim())
        .filter(Boolean)
        .sort((a, b) => a.localeCompare(b, 'de'));
    list.innerHTML = classes.map((c) => '<option value="' + escapeHtml(c) + '"></option>').join('');
}

function renderTeacherList() {
    const host = $('kcTeacherList');
    if (!host) return;
    const teachers = loadTenantTeachers()
        .filter((t) => t && t.email && String(t.email).indexOf('@') !== -1)
        .slice()
        .sort((a, b) => {
            const an = String(a.name || a.code || a.email);
            const bn = String(b.name || b.code || b.email);
            return an.localeCompare(bn, 'de');
        });
    if (!teachers.length) {
        host.innerHTML = '<p class="muted" style="margin:6px;">Keine Lehrkräfte in den Stammdaten – E-Mails unten eintragen.</p>';
        return;
    }
    host.innerHTML = teachers
        .map((t, i) => {
            const em = String(t.email).trim().toLowerCase();
            const label =
                (t.code ? String(t.code) + ' – ' : '') +
                (t.name ? String(t.name) + ' · ' : '') +
                em;
            const search = (String(t.code || '') + ' ' + String(t.name || '') + ' ' + em).toLowerCase();
            return (
                '<label class="kc-teacher-item" data-kc-teacher-search="' +
                escapeHtml(search) +
                '">' +
                '<input type="checkbox" data-kc-teacher-email="' +
                escapeHtml(em) +
                '" id="kcTeach' +
                i +
                '">' +
                '<span>' +
                escapeHtml(label) +
                '</span></label>'
            );
        })
        .join('');
}

function selectedTeacherEmails() {
    return Array.from(document.querySelectorAll('[data-kc-teacher-email]:checked')).map((el) =>
        String(el.getAttribute('data-kc-teacher-email') || '')
    );
}

function tenantPickedEmails() {
    return Array.from(tenantPicked.keys());
}

function userEmailFromGraph(u) {
    if (!u || typeof u !== 'object') return '';
    const mail = String(u.mail || '').trim().toLowerCase();
    if (mail && mail.indexOf('@') !== -1) return mail;
    const upn = String(u.userPrincipalName || '').trim().toLowerCase();
    if (upn && upn.indexOf('@') !== -1) return upn;
    return '';
}

function renderTenantPicked() {
    const wrap = $('kcTenantPickedWrap');
    const host = $('kcTenantPicked');
    if (!wrap || !host) return;
    if (!tenantPicked.size) {
        wrap.hidden = true;
        host.replaceChildren();
        return;
    }
    wrap.hidden = false;
    host.replaceChildren();
    tenantPicked.forEach((p) => {
        const chip = document.createElement('span');
        chip.className = 'kc-picked-chip';
        const txt = document.createElement('span');
        txt.textContent = p.label || p.email;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.setAttribute('aria-label', 'Entfernen');
        btn.textContent = '✕';
        btn.addEventListener('click', () => {
            tenantPicked.delete(p.email);
            renderTenantPicked();
        });
        chip.append(txt, btn);
        host.appendChild(chip);
    });
}

function fillTenantSearchChecklist(users) {
    const host = $('kcTenantSearchResults');
    const applyRow = $('kcTenantApplyRow');
    if (!host) return;
    host.replaceChildren();
    const list = Array.isArray(users) ? users : [];
    if (!list.length) {
        host.hidden = false;
        const p = document.createElement('div');
        p.className = 'muted';
        p.style.padding = '8px';
        p.textContent = '(keine Treffer)';
        host.appendChild(p);
        if (applyRow) applyRow.hidden = true;
        return;
    }

    host.hidden = false;
    if (applyRow) applyRow.hidden = false;

    const selAllRow = document.createElement('label');
    selAllRow.className = 'kc-user-checklist__selectall';
    const selAllCb = document.createElement('input');
    selAllCb.type = 'checkbox';
    selAllCb.setAttribute('aria-label', 'Alle auswählen');
    const selAllTxt = document.createElement('span');
    selAllTxt.textContent = 'Alle auswählen (' + list.length + ')';
    selAllRow.append(selAllCb, selAllTxt);
    host.appendChild(selAllRow);

    const api = G();
    const items = [];
    list.forEach((u) => {
        const email = userEmailFromGraph(u);
        const id = u && u.id ? String(u.id) : '';
        if (!email && !id) return;
        const label = document.createElement('label');
        label.className = 'kc-user-checklist__item';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = id;
        cb.dataset.kcUserEmail = email;
        cb.dataset.kcUserLabel = typeof api.personLabel === 'function' ? api.personLabel(u) : email || id;
        if (email && tenantPicked.has(email)) {
            cb.checked = true;
            label.classList.add('is-checked');
        }
        const txt = document.createElement('span');
        txt.textContent = cb.dataset.kcUserLabel || email || id;
        label.append(cb, txt);
        cb.addEventListener('change', () => {
            label.classList.toggle('is-checked', !!cb.checked);
            syncSelectAll();
        });
        host.appendChild(label);
        items.push({ cb, label });
    });

    function syncSelectAll() {
        const all = host.querySelectorAll('.kc-user-checklist__item input[type="checkbox"]');
        const checkedCount = host.querySelectorAll(
            '.kc-user-checklist__item input[type="checkbox"]:checked'
        ).length;
        selAllCb.indeterminate = checkedCount > 0 && checkedCount < all.length;
        selAllCb.checked = all.length > 0 && checkedCount === all.length;
    }

    selAllCb.addEventListener('change', () => {
        items.forEach((it) => {
            it.cb.checked = selAllCb.checked;
            it.label.classList.toggle('is-checked', selAllCb.checked);
        });
        syncSelectAll();
    });
    syncSelectAll();
}

async function runTenantUserSearch() {
    const inp = $('kcTenantSearch');
    const q = inp ? String(inp.value || '').trim() : '';
    if (!q) {
        toast('Bitte einen Suchbegriff eingeben (Name oder E-Mail).');
        return;
    }
    const btn = $('kcBtnTenantSearch');
    if (btn) btn.disabled = true;
    try {
        const token = await getToken();
        const users = await G().searchUsers(token, q);
        fillTenantSearchChecklist(users);
        toast('Suche: ' + users.length + ' Treffer.');
    } catch (e) {
        toast('Tenant-Suche: ' + ((e && e.message) || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

function applyTenantSelection() {
    const host = $('kcTenantSearchResults');
    const checked = host
        ? Array.from(host.querySelectorAll('.kc-user-checklist__item input[type="checkbox"]:checked'))
        : [];
    if (!checked.length) {
        toast('Bitte mindestens einen Treffer auswählen.');
        return;
    }
    let added = 0;
    checked.forEach((cb) => {
        const email = String(cb.dataset.kcUserEmail || '').trim().toLowerCase();
        const label = String(cb.dataset.kcUserLabel || email).trim();
        const id = String(cb.value || '').trim();
        if (!email || email.indexOf('@') === -1) return;
        if (!tenantPicked.has(email)) added += 1;
        tenantPicked.set(email, { email, label, id });
    });
    renderTenantPicked();
    toast(
        added
            ? added + ' Person(en) übernommen.'
            : 'Auswahl war bereits übernommen.'
    );
}

function updateManualKvHint() {
    const el = $('kcManualKvHint');
    if (!el) return;
    const klasse = ($('kcManualKlasse') && $('kcManualKlasse').value.trim()) || '';
    if (!klasse) {
        el.textContent = 'Klasse wählen – dann wird der KV aus den Stammdaten angezeigt.';
        return;
    }
    const kv = buildKvByKlasse(loadTenantClasses());
    const hit = kv.get(klasse) || kv.get(klasse.toUpperCase());
    if (hit) {
        el.textContent = 'KV: ' + (hit.name ? hit.name + ' · ' : '') + hit.email;
    } else {
        el.textContent = 'Für „' + klasse + '“ ist in den Stammdaten kein KV hinterlegt.';
    }
}

function updateManualQueueHint() {
    const el = $('kcManualQueueHint');
    const btnClear = $('kcBtnClearManualQueue');
    if (btnClear) btnClear.hidden = manualQueue.length === 0;
    if (!el) return;
    if (mode !== 'manual') {
        el.textContent = '';
        return;
    }
    if (!manualQueue.length) {
        el.textContent =
            'Noch keine manuellen Chats in der Liste. Formular ausfüllen und „Zur Liste hinzufügen“ – oder nur das Formular nutzen (wird in der Vorschau als einzelner Chat genommen).';
        return;
    }
    el.textContent =
        manualQueue.length +
        ' manueller Chat' +
        (manualQueue.length === 1 ? '' : 's') +
        ' in der Liste: ' +
        manualQueue.map((q) => q.klasse).join(', ');
}

function addCurrentManualToQueue() {
    const klasse = ($('kcManualKlasse') && $('kcManualKlasse').value.trim()) || '';
    const includeKv = !$('kcManualIncludeKv') || !!$('kcManualIncludeKv').checked;
    const membersText = ($('kcManualEmails') && $('kcManualEmails').value) || '';
    const emails = selectedTeacherEmails().concat(tenantPickedEmails());
    const plan = buildManualClassChatPlan({
        klasse,
        memberEmails: emails,
        membersText,
        includeKv,
        yearPrefix: ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || yearPrefix,
        namePattern: getPatternFromBuilder(),
        kvByKlasse: buildKvByKlasse(loadTenantClasses()),
        teacherDirectory: loadTenantTeachers(),
        existingByKlasse: existingByKlasseFromState(loadClassChatsState())
    });
    if (!plan.klasse) {
        toast('Bitte eine Klasse angeben.');
        return;
    }
    if (plan.memberEmails.length < 2) {
        toast('Mindestens 2 Mitglieder mit E-Mail nötig (inkl. KV falls aktiv).');
        return;
    }
    const idx = manualQueue.findIndex((q) => q.klasse.toUpperCase() === plan.klasse.toUpperCase());
    const entry = {
        klasse: plan.klasse,
        memberEmails: plan.memberEmails.slice(),
        includeKv
    };
    if (idx >= 0) manualQueue[idx] = entry;
    else manualQueue.push(entry);
    updateManualQueueHint();
    toast('„' + plan.klasse + '“ zur Liste hinzugefügt (' + plan.memberEmails.length + ' Mitglieder).');
}

function buildPlansFromManualFormOrQueue() {
    yearPrefix = ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || calcYearPrefix();
    namePattern = getPatternFromBuilder();
    const kv = buildKvByKlasse(loadTenantClasses());
    const existing = existingByKlasseFromState(loadClassChatsState());
    const teachers = loadTenantTeachers();
    const optsBase = {
        yearPrefix,
        namePattern,
        kvByKlasse: kv,
        teacherDirectory: teachers,
        existingByKlasse: existing
    };

    if (manualQueue.length) {
        return manualQueue.map((q) =>
            buildManualClassChatPlan({
                ...optsBase,
                klasse: q.klasse,
                memberEmails: q.memberEmails,
                includeKv: q.includeKv !== false
            })
        );
    }

    // Einzelner Chat aus dem Formular (ohne explizit „hinzufügen“)
    return [
        buildManualClassChatPlan({
            ...optsBase,
            klasse: ($('kcManualKlasse') && $('kcManualKlasse').value.trim()) || '',
            memberEmails: selectedTeacherEmails().concat(tenantPickedEmails()),
            membersText: ($('kcManualEmails') && $('kcManualEmails').value) || '',
            includeKv: !$('kcManualIncludeKv') || !!$('kcManualIncludeKv').checked
        })
    ];
}

function refreshSourcePanel() {
    belegung = loadBelegung();
    const status = $('kcSourceStatus');
    const hint = $('kcSourceHint');
    if (!status) return;
    if (!belegung || !Array.isArray(belegung.rows) || !belegung.rows.length) {
        status.innerHTML =
            '<span class="kc-status kc-status--warn">Keine Unterrichtsbelegung gefunden.</span>';
        if (hint) {
            hint.innerHTML =
                'Bitte zuerst unter <a href="kursteams.html">Unterrichtsteams</a> die Team-Namen generieren – oder den Modus <strong>Manuell</strong> nutzen.';
        }
        return;
    }
    const when = belegung.updatedAt
        ? (() => {
              try {
                  return new Date(belegung.updatedAt).toLocaleString('de-AT');
              } catch {
                  return belegung.updatedAt;
              }
          })()
        : '–';
    status.innerHTML =
        '<span class="kc-status kc-status--ok">' +
        escapeHtml(String(belegung.rows.length)) +
        ' Einträge · ' +
        escapeHtml(String(belegung.classCount || '–')) +
        ' Klassen · ' +
        escapeHtml(String(belegung.teacherCount || '–')) +
        ' Lehrkräfte' +
        (belegung.yearPrefix ? ' · ' + escapeHtml(belegung.yearPrefix) : '') +
        '</span>';
    if (hint) {
        hint.textContent =
            'Stand: ' + when + ' · KV aus den Stammdaten wird immer ergänzt · mind. 2 Lehrkräfte pro Chat.';
    }
    if (belegung.yearPrefix && $('kcYearPrefix') && !$('kcYearPrefix').dataset.userEdited) {
        $('kcYearPrefix').value = belegung.yearPrefix;
        yearPrefix = belegung.yearPrefix;
        updateNamePreview();
    }
}

function rebuildPlans() {
    yearPrefix = ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || calcYearPrefix();
    namePattern = getPatternFromBuilder();
    if (mode === 'manual') {
        plans = buildPlansFromManualFormOrQueue();
        return plans;
    }
    belegung = loadBelegung();
    const kv = buildKvByKlasse(loadTenantClasses());
    const existing = existingByKlasseFromState(loadClassChatsState());
    plans = buildClassChatPlans(belegung, kv, {
        yearPrefix,
        namePattern,
        existingByKlasse: existing
    });
    return plans;
}

function statusLabel(p) {
    if (p.status === 'skip') return 'Überspringen';
    if (p.status === 'neu') return 'Neu';
    if (p.status === 'sync') return 'Sync (Name)';
    if (p.status === 'ok') return 'Vorhanden';
    return p.status;
}

function renderPreview() {
    rebuildPlans();
    const sum = summarizePlans(plans);
    const sumEl = $('kcPreviewSummary');
    if (sumEl) {
        sumEl.textContent =
            sum.eligible +
            ' Chats anlegbar · ' +
            sum.neu +
            ' neu · ' +
            sum.sync +
            ' mit vorhandener ID · ' +
            sum.skip +
            ' übersprungen';
    }
    const tbody = $('kcPreviewBody');
    if (!tbody) return;
    if (!plans.length) {
        tbody.innerHTML =
            '<tr><td colspan="5" class="muted">' +
            (mode === 'manual'
                ? 'Kein manueller Chat – Klasse und mindestens 2 Mitglieder angeben.'
                : 'Keine Klassen aus der Belegung – bitte Quelle prüfen oder Manuell wählen.') +
            '</td></tr>';
        return;
    }
    tbody.innerHTML = plans
        .map((p) => {
            const members = p.members
                .map((m) => {
                    const tag = m.role === 'kv' ? ' (KV)' : '';
                    return escapeHtml((m.code || m.email) + tag);
                })
                .join(', ');
            const cls = p.eligible ? '' : ' class="kc-row-skip"';
            return (
                '<tr' +
                cls +
                '>' +
                '<td><strong>' +
                escapeHtml(p.klasse) +
                '</strong></td>' +
                '<td>' +
                escapeHtml(p.topic) +
                '</td>' +
                '<td>' +
                p.memberEmails.length +
                '<div class="kc-member-list muted">' +
                members +
                '</div></td>' +
                '<td>' +
                escapeHtml(statusLabel(p)) +
                (p.skipReason ? '<div class="muted">' + escapeHtml(p.skipReason) + '</div>' : '') +
                '</td>' +
                '<td class="muted" style="font-size:0.8em;">' +
                (p.chatId ? escapeHtml(p.chatId.slice(0, 8)) + '…' : '–') +
                '</td>' +
                '</tr>'
            );
        })
        .join('');
}

/* ── Name builder (wie Kursteams / Spielwiesen) ── */

function getPatternFromBuilder() {
    const zone = $('kcNameBuilder');
    if (!zone) return normalizeChatNamePattern(namePattern);
    const tokens = [];
    zone.querySelectorAll('[data-token-type]').forEach((el) => {
        const type = String(el.getAttribute('data-token-type') || '');
        if (type === 'text') {
            tokens.push({ type: 'text', value: String(el.getAttribute('data-token-value') || '') });
        } else tokens.push({ type });
    });
    return normalizeChatNamePattern(tokens);
}

function updateNamePreview() {
    const el = $('kcNamePreview');
    if (!el) return;
    const yp = ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || yearPrefix || 'SJ26';
    const klassePreview =
        (mode === 'manual' && $('kcManualKlasse') && $('kcManualKlasse').value.trim()) || '1A';
    const preview = buildChatTopicFromPattern(getPatternFromBuilder(), {
        yearPrefix: yp,
        klasse: klassePreview
    });
    el.textContent = 'Vorschau: ' + preview;
}

function addChip(zone, token) {
    const chip = document.createElement('span');
    chip.className = 'name-chip';
    chip.draggable = true;
    chip.setAttribute('data-token-type', token.type);
    if (token.type === 'text') chip.setAttribute('data-token-value', String(token.value ?? ''));

    const txt = document.createElement('span');
    txt.textContent = token.type === 'text' ? chatTokenLabel(token) : chatTokenLabel(token);

    const x = document.createElement('button');
    x.type = 'button';
    x.className = 'chip-x';
    x.textContent = '✕';
    x.title = 'Baustein entfernen';
    x.addEventListener('click', () => {
        chip.remove();
        namePattern = getPatternFromBuilder();
        updateNamePreview();
    });

    chip.append(txt, x);
    zone.appendChild(chip);
}

function renderNameBuilder() {
    const zone = $('kcNameBuilder');
    if (!zone) return;
    zone.innerHTML = '';
    normalizeChatNamePattern(namePattern).forEach((t) => addChip(zone, t));
    updateNamePreview();
}

function wireNameBuilder() {
    const zone = $('kcNameBuilder');
    if (!zone || zone.dataset.wired) return;
    zone.dataset.wired = '1';

    let dragEl = null;
    zone.addEventListener('dragstart', (e) => {
        const target = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
        if (!target) return;
        dragEl = target;
        target.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
    });
    zone.addEventListener('dragend', () => {
        if (dragEl) dragEl.classList.remove('dragging');
        dragEl = null;
    });
    zone.addEventListener('dragover', (e) => {
        e.preventDefault();
        const over = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
        if (!dragEl || !over || over === dragEl) return;
        const rect = over.getBoundingClientRect();
        if (e.clientX > rect.left + rect.width / 2) over.after(dragEl);
        else over.before(dragEl);
    });
    zone.addEventListener('drop', (e) => {
        e.preventDefault();
        namePattern = getPatternFromBuilder();
        updateNamePreview();
    });

    document.querySelectorAll('[data-kc-token]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const type = btn.getAttribute('data-kc-token');
            if (type === 'yearPrefix' || type === 'klasse') addChip(zone, { type });
            namePattern = getPatternFromBuilder();
            updateNamePreview();
        });
    });

    const btnSep = $('kcNameAddSep');
    if (btnSep) {
        btnSep.addEventListener('click', () => {
            const v = $('kcNameSepValue') ? $('kcNameSepValue').value : ' | ';
            addChip(zone, { type: 'text', value: String(v ?? '') });
            namePattern = getPatternFromBuilder();
            updateNamePreview();
        });
    }
    const btnText = $('kcNameAddText');
    if (btnText) {
        btnText.addEventListener('click', () => {
            const v = $('kcNameTextValue') ? $('kcNameTextValue').value : '';
            addChip(zone, { type: 'text', value: String(v ?? '') });
            namePattern = getPatternFromBuilder();
            updateNamePreview();
        });
    }
    const btnReset = $('kcNameResetDefault');
    if (btnReset) {
        btnReset.addEventListener('click', () => {
            namePattern = defaultChatNamePattern();
            renderNameBuilder();
        });
    }
}

function log(msg) {
    const el = $('kcLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
    el.scrollTop = el.scrollHeight;
}

async function runProvision() {
    if (busy) return;
    rebuildPlans();
    const eligible = plans.filter((p) => p.eligible);
    if (!eligible.length) {
        toast('Keine anlegbaren Chats – Belegung, Formular oder Mitglieder prüfen.');
        return;
    }

    busy = true;
    const btn = $('kcBtnRun');
    if (btn) btn.disabled = true;
    const logEl = $('kcLog');
    if (logEl) logEl.textContent = '';

    try {
        log('Anmeldung bei Microsoft …');
        const token = await getToken();
        const me = await G().graphJson('GET', '/me?$select=id,mail,userPrincipalName', token, undefined);
        const meId = me && me.id ? String(me.id) : '';
        log('Angemeldet. ' + eligible.length + ' Chat(s) werden bearbeitet …');

        let state = loadClassChatsState();
        let ok = 0;
        let fail = 0;

        for (let i = 0; i < eligible.length; i++) {
            const plan = eligible[i];
            log('[' + (i + 1) + '/' + eligible.length + '] ' + plan.klasse + ' → ' + plan.topic);
            try {
                const res = await provisionClassChat(G(), token, plan, meId);
                state = upsertClassChatItem(state, {
                    klasse: plan.klasse,
                    topic: res.topic,
                    chatId: res.chatId,
                    memberEmails: res.memberEmails,
                    webUrl: res.webUrl,
                    yearPrefix,
                    namePattern
                });
                saveClassChatsState(state);
                ok += 1;
                const bits = [];
                if (res.created) bits.push('angelegt');
                else bits.push('aktualisiert');
                if (res.topicUpdated) bits.push('Topic geändert');
                if (res.added) bits.push('+' + res.added + ' Mitglied(er)');
                if (res.missingEmails && res.missingEmails.length) {
                    bits.push('ohne Graph: ' + res.missingEmails.join(', '));
                }
                log('  OK – ' + bits.join(', '));
            } catch (e) {
                fail += 1;
                log('  Fehler: ' + (e && e.message ? e.message : String(e)));
            }
        }

        persistPatternMeta(loadClassChatsState());
        log('Fertig: ' + ok + ' ok, ' + fail + ' Fehler.');
        toast('KlassenChats: ' + ok + ' ok' + (fail ? ', ' + fail + ' Fehler' : '') + '.');
        renderPreview();
    } catch (e) {
        log('Abbruch: ' + (e && e.message ? e.message : String(e)));
        toast('Anmeldung oder Graph fehlgeschlagen.');
    } finally {
        busy = false;
        if (btn) btn.disabled = false;
    }
}

function boot() {
    const saved = loadClassChatsState();
    if (saved && saved.namePattern) namePattern = normalizeChatNamePattern(saved.namePattern);
    if (saved && saved.yearPrefix) yearPrefix = saved.yearPrefix;

    const yp = $('kcYearPrefix');
    if (yp) {
        yp.value = yearPrefix || calcYearPrefix();
        yp.addEventListener('input', () => {
            yp.dataset.userEdited = '1';
            yearPrefix = yp.value.trim();
            updateNamePreview();
        });
    }

    wireNameBuilder();
    renderNameBuilder();
    refreshClassDatalist();
    renderTeacherList();
    refreshSourcePanel();

    // Ohne Belegung: Manuell vorwählen
    belegung = loadBelegung();
    if (!belegung || !Array.isArray(belegung.rows) || !belegung.rows.length) {
        setMode('manual');
    } else {
        setMode('belegung');
    }

    document.querySelectorAll('[data-kc-mode]').forEach((btn) => {
        btn.addEventListener('click', () => setMode(btn.getAttribute('data-kc-mode')));
    });

    const klasseInp = $('kcManualKlasse');
    if (klasseInp) {
        klasseInp.addEventListener('input', () => {
            updateManualKvHint();
            updateNamePreview();
        });
        klasseInp.addEventListener('change', () => {
            updateManualKvHint();
            updateNamePreview();
        });
    }

    const filter = $('kcTeacherFilter');
    if (filter) {
        filter.addEventListener('input', () => {
            const q = String(filter.value || '')
                .trim()
                .toLowerCase();
            document.querySelectorAll('.kc-teacher-item').forEach((row) => {
                const hay = String(row.getAttribute('data-kc-teacher-search') || '');
                row.hidden = !!(q && hay.indexOf(q) === -1);
            });
        });
    }

    const btnAdd = $('kcBtnAddManual');
    if (btnAdd) btnAdd.addEventListener('click', () => addCurrentManualToQueue());
    const btnClear = $('kcBtnClearManualQueue');
    if (btnClear) {
        btnClear.addEventListener('click', () => {
            manualQueue = [];
            updateManualQueueHint();
            toast('Manuelle Liste geleert.');
        });
    }

    const btnTenantSearch = $('kcBtnTenantSearch');
    if (btnTenantSearch) btnTenantSearch.addEventListener('click', () => runTenantUserSearch());
    const tenantSearch = $('kcTenantSearch');
    if (tenantSearch) {
        tenantSearch.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                runTenantUserSearch();
            }
        });
    }
    const btnTenantApply = $('kcBtnTenantApply');
    if (btnTenantApply) btnTenantApply.addEventListener('click', () => applyTenantSelection());
    renderTenantPicked();

    document.querySelectorAll('[data-kc-goto]').forEach((btn) => {
        btn.addEventListener('click', () => setStep(btn.getAttribute('data-kc-goto')));
    });
    document.querySelectorAll('[data-kc-next]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const n = Number(btn.getAttribute('data-kc-next'));
            if (n === 3 || n === 4) rebuildPlans();
            if (mode === 'manual' && n === 3) {
                const eligible = plans.filter((p) => p.eligible);
                if (!eligible.length) {
                    toast('Manuell: Klasse und mind. 2 E-Mails nötig (oder zur Liste hinzufügen).');
                    return;
                }
            }
            if (n >= 2) persistPatternMeta(loadClassChatsState());
            setStep(n);
        });
    });
    document.querySelectorAll('[data-kc-prev]').forEach((btn) => {
        btn.addEventListener('click', () => setStep(btn.getAttribute('data-kc-prev')));
    });

    const btnReload = $('kcBtnReloadSource');
    if (btnReload) {
        btnReload.addEventListener('click', () => {
            refreshSourcePanel();
            refreshClassDatalist();
            renderTeacherList();
        });
    }

    const btnRun = $('kcBtnRun');
    if (btnRun) btnRun.addEventListener('click', () => runProvision());

    setStep(1);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
