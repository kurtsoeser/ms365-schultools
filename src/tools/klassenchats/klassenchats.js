/**
 * KlassenChats-Wizard: Gruppenchats oder private Kanäle für Lehrkräfte einer Klasse.
 * Transport: shared/graph-client.js (R22 / b11).
 */
import { graphJson as sharedGraphJson } from '../../shared/graph-client.js';
import { escapeHtml } from '../../shared/utils/strings.js';
import {
    calcYearPrefix,
    defaultChatNamePattern,
    normalizeChatNamePattern,
    chatTokenLabel,
    buildChatTopicFromPattern,
    displayNameForTarget,
    buildKvByKlasse,
    buildClassChatPlans,
    buildManualClassChatPlan,
    existingByKlasseFromState,
    upsertClassChatItem,
    normalizeClassChatsState,
    normalizeExtraMemberEmails,
    normalizeTargetKind,
    parseMemberEmailsText,
    teamsHrefForPlan,
    targetKindLabel
} from './klassenchats-logic.js';
import {
    formatKlassenchatsGraphError,
    prepareTeacherTeamForChannels,
    provisionClassResource
} from './klassenchats-graph.js';

function $(id) {
    return document.getElementById(id);
}

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}


function G() {
    const api = window.ms365GraphUnifiedGroups;
    if (!api) throw new Error('Graph nicht geladen.');
    return api;
}

const SCOPES_CHAT = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Chat.Create',
    'https://graph.microsoft.com/Chat.ReadWrite'
];

const SCOPES_CHANNEL = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/Channel.Create',
    'https://graph.microsoft.com/Channel.ReadBasic.All',
    'https://graph.microsoft.com/ChannelSettings.ReadWrite.All',
    'https://graph.microsoft.com/ChannelMember.ReadWrite.All'
];

function scopesForTarget() {
    return targetKind === 'privateChannel' ? SCOPES_CHANNEL : SCOPES_CHAT;
}

async function getToken() {
    if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
        return window.ms365AuthAcquireTokenPopup(scopesForTarget());
    }
    return G().getGraphToken();
}

/** @type {ReturnType<typeof defaultChatNamePattern>} */
let namePattern = defaultChatNamePattern();
let step = 1;
/** @type {'belegung'|'manual'} */
let mode = 'belegung';
/** @type {'groupChat'|'privateChannel'} */
let targetKind = 'groupChat';
/** @type {{ id: string, displayName: string, mail: string }|null} */
let lehrerTeam = null;
/** @type {ReturnType<typeof buildClassChatPlans>} */
let plans = [];
/** Klassen, die die IT in dieser Sitzung nicht anlegen will. Schlüssel = Klasse in Großbuchstaben. */
const omittedKlassen = new Set();
/** Höchster Schritt, den „Weiter“ schon freigegeben hat. */
let maxReached = 1;
/** Manuelle Warteschlange (mehrere Chats vor dem Anlegen). */
/** @type {Array<{ klasse: string, memberEmails: string[], includeKv: boolean }>} */
let manualQueue = [];
/** Aus Tenant-Suche übernommene Personen: email → { email, label, id } */
/** @type {Map<string, { email: string, label: string, id: string }>} */
let tenantPicked = new Map();
/** Globale Zusatzpersonen (Direktion etc.): email → { email, label, id } */
/** @type {Map<string, { email: string, label: string, id: string }>} */
let extraPicked = new Map();
/** @type {object|null} */
let belegung = null;
let yearPrefix = calcYearPrefix();
let busy = false;
/** @type {'confirm'|'exec'|'result'} */
let runPhase = 'confirm';
/** @type {Array<{ klasse: string, topic: string, ok: boolean, created?: boolean, updated?: boolean, error?: string, detail?: string }>} */
let lastRunResults = [];

function guidLike(s) {
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(
        String(s || '').trim()
    );
}

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
    const saved = api.setClassChats(state);
    const n = saved && Array.isArray(saved.items) ? saved.items.length : 0;
    const want = state && Array.isArray(state.items) ? state.items.length : 0;
    if (want > 0 && n === 0) {
        toast('Warnung: Klassenchats konnten nicht dauerhaft gespeichert werden (Browser-Speicher voll?).');
    }
    return saved;
}

function getExtraMemberEmails() {
    const fromText = parseMemberEmailsText(($('kcExtraEmails') && $('kcExtraEmails').value) || '');
    const fromPicked = Array.from(extraPicked.keys());
    return normalizeExtraMemberEmails(fromText.concat(fromPicked));
}

function renderExtraPicked() {
    const wrap = $('kcExtraPickedWrap');
    const host = $('kcExtraPicked');
    if (!wrap || !host) return;
    if (!extraPicked.size) {
        wrap.hidden = true;
        host.replaceChildren();
        return;
    }
    wrap.hidden = false;
    host.replaceChildren();
    extraPicked.forEach((p) => {
        const chip = document.createElement('span');
        chip.className = 'kc-picked-chip';
        const txt = document.createElement('span');
        txt.textContent = p.label || p.email;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.setAttribute('aria-label', 'Entfernen');
        btn.textContent = '✕';
        btn.addEventListener('click', () => {
            extraPicked.delete(p.email);
            renderExtraPicked();
            persistPatternMeta(loadClassChatsState());
            if (step >= 3) renderPreview();
        });
        chip.append(txt, btn);
        host.appendChild(chip);
    });
}

function fillExtraSearchChecklist(users) {
    const host = $('kcExtraSearchResults');
    const applyRow = $('kcExtraApplyRow');
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
    const api = G();
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
        cb.dataset.kcUserLabel =
            typeof api.personLabel === 'function' ? api.personLabel(u) : email || id;
        if (email && extraPicked.has(email)) {
            cb.checked = true;
            label.classList.add('is-checked');
        }
        const txt = document.createElement('span');
        txt.textContent = cb.dataset.kcUserLabel || email || id;
        label.append(cb, txt);
        cb.addEventListener('change', () => {
            label.classList.toggle('is-checked', !!cb.checked);
        });
        host.appendChild(label);
    });
}

async function runExtraUserSearch() {
    const inp = $('kcExtraSearch');
    const q = inp ? String(inp.value || '').trim() : '';
    if (!q) {
        toast('Bitte einen Suchbegriff eingeben (Name oder E-Mail).');
        return;
    }
    const btn = $('kcBtnExtraSearch');
    if (btn) btn.disabled = true;
    try {
        const token = await getToken();
        const users = await G().searchUsers(token, q);
        fillExtraSearchChecklist(users);
        toast('Zusatzsuche: ' + users.length + ' Treffer.');
    } catch (e) {
        toast('Zusatzsuche: ' + ((e && e.message) || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

function applyExtraSelection() {
    const host = $('kcExtraSearchResults');
    const checked = host
        ? Array.from(host.querySelectorAll('.kc-user-checklist__item input[type="checkbox"]:checked'))
        : [];
    if (!checked.length) {
        toast('Bitte mindestens einen Treffer auswählen.');
        return;
    }
    let added = 0;
    checked.forEach((cb) => {
        const email = String(cb.dataset.kcUserEmail || '')
            .trim()
            .toLowerCase();
        const label = String(cb.dataset.kcUserLabel || email).trim();
        const id = String(cb.value || '').trim();
        if (!email || email.indexOf('@') === -1) return;
        if (!extraPicked.has(email)) added += 1;
        extraPicked.set(email, { email, label, id });
    });
    renderExtraPicked();
    persistPatternMeta(loadClassChatsState());
    if (step >= 3) renderPreview();
    toast(added ? added + ' Zusatzperson(en) übernommen.' : 'Auswahl war bereits in der Zusatzliste.');
}

function persistPatternMeta(state) {
    const extras = getExtraMemberEmails();
    const base = normalizeClassChatsState(state) || {
        updatedAt: new Date().toISOString(),
        yearPrefix,
        namePattern,
        targetKind,
        lehrerTeamId: lehrerTeam ? lehrerTeam.id : '',
        lehrerTeamDisplayName: lehrerTeam ? lehrerTeam.displayName : '',
        extraMemberEmails: extras,
        items: []
    };
    return saveClassChatsState({
        ...base,
        yearPrefix,
        namePattern,
        targetKind,
        lehrerTeamId: lehrerTeam ? lehrerTeam.id : '',
        lehrerTeamDisplayName: lehrerTeam ? lehrerTeam.displayName : '',
        extraMemberEmails: extras,
        updatedAt: new Date().toISOString()
    });
}

function setMode(next) {
    const prev = mode;
    mode = next === 'manual' ? 'manual' : 'belegung';
    document.querySelectorAll('[data-kc-mode]').forEach((btn) => {
        const on = btn.getAttribute('data-kc-mode') === mode;
        btn.classList.toggle('is-selected', on);
        btn.setAttribute('aria-checked', on ? 'true' : 'false');
    });
    const manualPanel = $('kcPanelManual');
    if (manualPanel) manualPanel.hidden = mode !== 'manual';
    const reload = $('kcBtnReloadSource');
    if (reload) reload.hidden = mode !== 'belegung';
    updateManualQueueHint();
    if (prev !== mode) {
        lockAhead();
        if (step >= 3) toast('Durch die geänderte Quelle wird die Vorschau neu berechnet.');
    }
    renderLists();
    if (mode === 'manual') updateManualKvHint();
}

function setTargetKind(next) {
    targetKind = normalizeTargetKind(next);
    document.querySelectorAll('[data-kc-target]').forEach((btn) => {
        const on = btn.getAttribute('data-kc-target') === targetKind;
        btn.classList.toggle('is-selected', on);
        btn.setAttribute('aria-checked', on ? 'true' : 'false');
    });
    const teamPanel = $('kcPanelLehrerTeam');
    if (teamPanel) teamPanel.hidden = targetKind !== 'privateChannel';
    lockAhead();
    renderLists();
    renderLehrerTeamPicked();
}

function updateTargetUiCopy() {
    const isChannel = targetKind === 'privateChannel';
    const lead = $('kcPreviewLead');
    if (lead) {
        lead.textContent = isChannel
            ? 'Angehakte Klassen werden als private Kanäle angelegt. Nicht anlegbare Zeilen bleiben deaktiviert.'
            : 'Angehakte Klassen werden als Gruppenchats angelegt. Nicht anlegbare Zeilen bleiben deaktiviert.';
    }
    document.querySelectorAll('[data-kc-col-name]').forEach((el) => {
        el.textContent = isChannel ? 'Kanalname' : 'Chat-Name';
    });
    const runHint = $('kcRunHint');
    if (runHint) {
        runHint.textContent = 'Kontrolliert ausführen und live sehen, was passiert.';
    }
    updateRunButton();
    paintConfirmPanel();
}

function renderLehrerTeamPicked() {
    const el = $('kcTeamPicked');
    if (!el) return;
    if (!lehrerTeam || !lehrerTeam.id) {
        el.innerHTML = 'Noch kein Lehrerteam gewählt.';
        el.classList.add('muted');
        return;
    }
    el.classList.remove('muted');
    const mail = lehrerTeam.mail ? ' · ' + escapeHtml(lehrerTeam.mail) : '';
    el.innerHTML =
        'Gewählt: <strong>' +
        escapeHtml(lehrerTeam.displayName || lehrerTeam.id) +
        '</strong>' +
        mail +
        ' <button type="button" class="btn btn-small" id="kcBtnClearTeam" style="margin-left:8px;">Entfernen</button>';
    const btn = $('kcBtnClearTeam');
    if (btn) {
        btn.addEventListener('click', () => {
            lehrerTeam = null;
            lockAhead();
            renderLehrerTeamPicked();
            persistPatternMeta(loadClassChatsState());
        });
    }
}

function pickLehrerTeam(team) {
    if (!team || !team.id) return;
    lehrerTeam = {
        id: String(team.id),
        displayName: String(team.displayName || ''),
        mail: String(team.mail || '')
    };
    renderLehrerTeamPicked();
    persistPatternMeta(loadClassChatsState());
    toast('Lehrerteam „' + (lehrerTeam.displayName || lehrerTeam.id) + '“ gewählt.');
}

async function runTeamSearch() {
    const inp = $('kcTeamSearch');
    const q = inp ? String(inp.value || '').trim() : '';
    if (!q) {
        toast('Bitte einen Suchbegriff für das Lehrerteam eingeben.');
        return;
    }
    const btn = $('kcBtnTeamSearch');
    if (btn) btn.disabled = true;
    try {
        const token = await getToken();
        const api = G();
        let list = [];
        if (guidLike(q) && typeof api.fetchGroup === 'function') {
            const g = await api.fetchGroup(token, q);
            if (g && g.id) list = [g];
        } else if (typeof api.searchUnifiedGroups === 'function') {
            list = (await api.searchUnifiedGroups(token, q)) || [];
        }
        fillTeamSearchResults(list);
        toast('Team-Suche: ' + list.length + ' Treffer.');
    } catch (e) {
        toast('Team-Suche: ' + ((e && e.message) || e));
    } finally {
        if (btn) btn.disabled = false;
    }
}

function fillTeamSearchResults(groups) {
    const host = $('kcTeamSearchResults');
    if (!host) return;
    host.replaceChildren();
    const list = Array.isArray(groups) ? groups : [];
    if (!list.length) {
        host.hidden = false;
        const p = document.createElement('div');
        p.className = 'muted';
        p.style.padding = '8px';
        p.textContent = '(keine Treffer)';
        host.appendChild(p);
        return;
    }
    host.hidden = false;
    list.forEach((g) => {
        if (!g || !g.id) return;
        const row = document.createElement('button');
        row.type = 'button';
        row.className = 'kc-user-checklist__item';
        row.style.cssText =
            'display:block;width:100%;text-align:left;border:0;background:transparent;padding:8px;cursor:pointer;font:inherit;color:inherit;';
        const label = (g.displayName || g.mail || g.id) + (g.mail ? ' · ' + g.mail : '');
        row.textContent = label;
        row.addEventListener('click', () => {
            pickLehrerTeam({
                id: g.id,
                displayName: g.displayName || '',
                mail: g.mail || ''
            });
            host.hidden = true;
        });
        host.appendChild(row);
    });
}

function lockAhead() {
    maxReached = 1;
    updateStepPills();
}

function updateStepPills() {
    document.querySelectorAll('.pw-steps [data-kc-goto], .kc-steps [data-kc-goto]').forEach((btn) => {
        const n = Number(btn.getAttribute('data-kc-goto'));
        const locked = n > maxReached;
        btn.disabled = locked;
        btn.classList.toggle('is-locked', locked);
        btn.setAttribute('aria-disabled', locked ? 'true' : 'false');
        const active = n === step;
        btn.classList.toggle('is-active', active);
        btn.classList.toggle('is-done', n < step && n <= maxReached);
        if (active) btn.setAttribute('aria-current', 'step');
        else btn.removeAttribute('aria-current');
    });
}

const HERO_BY_STEP = {
    1: {
        title: 'KlassenChats anlegen',
        sub: 'Quelle und Ziel wählen – danach Namen, Vorschau und Anlegen in Teams.'
    },
    2: {
        title: 'Namen festlegen',
        sub: 'Lege das Namensschema für die KlassenChats fest.'
    },
    3: {
        title: 'Vorschau prüfen',
        sub: 'Was würde passieren? Angehakte Klassen werden angelegt.'
    },
    4: {
        title: 'In Teams anlegen',
        sub: 'Kontrolliert ausführen und live sehen, was tatsächlich passiert.'
    }
};

function updateHero() {
    const conf = HERO_BY_STEP[step] || HERO_BY_STEP[1];
    const title = $('kcHeroTitle');
    const sub = $('kcHeroSub');
    if (title) title.textContent = conf.title;
    if (sub) sub.textContent = conf.sub;
}

function focusStepHeading() {
    const el = $('kcStep' + step + 'Heading');
    if (el && typeof el.focus === 'function') {
        try {
            el.focus({ preventScroll: false });
        } catch {
            el.focus();
        }
    }
}

function isTooFew(p) {
    return String((p && p.skipReason) || '').indexOf('Weniger als 2') === 0;
}

function actionablePlans() {
    return plans.filter((p) => p && p.eligible && !p.omitted && !isTooFew(p));
}

function blockReason(targetStep) {
    if (targetStep <= step) return '';
    if (targetStep > maxReached + 1) return 'Bitte der Reihe nach weitergehen.';
    if (targetStep > 1 && targetKind === 'privateChannel' && !(lehrerTeam && lehrerTeam.id)) {
        return 'Für private Kanäle bitte zuerst ein Lehrerteam wählen.';
    }
    if (targetStep >= 3 && mode === 'manual') {
        currentPlans();
        if (!actionablePlans().length) {
            return 'Manuell: Klasse und mind. 2 E-Mails nötig (oder zur Liste hinzufügen).';
        }
    }
    return '';
}

function goToStep(n) {
    const target = Math.max(1, Math.min(4, Number(n) || 1));
    if (target > step) {
        const reason = blockReason(target);
        if (reason) {
            toast(reason);
            return;
        }
        maxReached = Math.max(maxReached, target);
    }
    setStep(target);
}

function setStep(n) {
    step = Math.max(1, Math.min(4, Number(n) || 1));
    updateStepPills();
    document.querySelectorAll('[data-kc-step]').forEach((el) => {
        el.classList.toggle('is-active', Number(el.getAttribute('data-kc-step')) === step);
    });
    if (step === 2) updateNamePreview();
    else if (step === 4) {
        renderLists();
        if (runPhase !== 'exec' && runPhase !== 'result') showStep4Phase('confirm');
        paintConfirmPanel();
    } else {
        if (step < 4 && runPhase !== 'exec') {
            runPhase = 'confirm';
            showStep4Phase('confirm');
        }
        renderLists();
    }
    updateHero();
    focusStepHeading();
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
    yearPrefix = ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || yearPrefix;
    namePattern = getPatternFromBuilder();
    const plan = buildManualClassChatPlan({
        ...planOptsBase(),
        klasse,
        memberEmails: emails,
        membersText,
        includeKv,
        kvByKlasse: buildKvByKlasse(loadTenantClasses()),
        teacherDirectory: loadTenantTeachers()
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
    renderLists();
    toast('„' + plan.klasse + '“ zur Liste hinzugefügt (' + plan.memberEmails.length + ' Mitglieder).');
}

function planOptsBase() {
    const teamId = lehrerTeam && lehrerTeam.id ? lehrerTeam.id : '';
    return {
        yearPrefix,
        namePattern,
        targetKind,
        teamId,
        extraMemberEmails: getExtraMemberEmails(),
        teacherDirectory: loadTenantTeachers(),
        existingByKlasse: existingByKlasseFromState(loadClassChatsState(), {
            targetKind,
            teamId
        })
    };
}

function buildPlansFromManualFormOrQueue() {
    yearPrefix = ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || calcYearPrefix();
    namePattern = getPatternFromBuilder();
    const kv = buildKvByKlasse(loadTenantClasses());
    const teachers = loadTenantTeachers();
    const optsBase = {
        ...planOptsBase(),
        kvByKlasse: kv,
        teacherDirectory: teachers
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

function applyOmissions() {
    plans = plans.map((p) => {
        const key = String((p && p.klasse) || '').toUpperCase();
        if (!key || !omittedKlassen.has(key)) return p;
        return Object.assign({}, p, { omitted: true });
    });
}

function currentPlans() {
    rebuildPlans();
    applyOmissions();
    return plans;
}

function deCount(n, one, many) {
    return String(n) + ' ' + (Number(n) === 1 ? one : many);
}

function formatStand(iso) {
    if (!iso) return '–';
    try {
        return new Date(iso).toLocaleString('de-AT', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric',
            hour: '2-digit',
            minute: '2-digit'
        });
    } catch {
        return String(iso);
    }
}

function paintCheckList(host, items) {
    if (!host) return;
    if (!items || !items.length) {
        host.hidden = true;
        host.innerHTML = '';
        return;
    }
    host.hidden = false;
    host.innerHTML = items
        .map((it) => {
            const mark = it.ok ? '✓' : '⚠';
            return '<li>' + mark + ' ' + escapeHtml(it.text) + '</li>';
        })
        .join('');
}

function paintSourceBanner() {
    const status = $('kcSourceStatus');
    const hint = $('kcSourceHint');
    const checks = $('kcSourceChecks');
    const banner = $('kcSourceBanner');
    const reload = $('kcBtnReloadSource');
    if (reload) reload.hidden = mode !== 'belegung';
    if (!status) return;

    if (mode === 'manual') {
        const ready = actionablePlans().length;
        if (banner) banner.className = 'pw-msg ' + (ready ? 'pw-msg--ok' : 'pw-msg--warn');
        status.textContent = ready
            ? deCount(ready, 'Klasse', 'Klassen') + ' manuell bereit'
            : 'Noch keine Klasse gewählt';
        paintCheckList(
            checks,
            ready
                ? [{ ok: true, text: deCount(ready, 'Klasse kann', 'Klassen können') + ' angelegt werden' }]
                : [{ ok: false, text: 'Klasse und mindestens zwei Lehrkräfte mit E-Mail festlegen' }]
        );
        if (hint) hint.textContent = 'Manuelle Auswahl · Quelle jederzeit änderbar.';
        return;
    }

    belegung = loadBelegung();
    if (!belegung || !Array.isArray(belegung.rows) || !belegung.rows.length) {
        if (banner) banner.className = 'pw-msg pw-msg--warn';
        status.textContent = 'Keine Unterrichtsbelegung gefunden';
        paintCheckList(checks, [
            {
                ok: false,
                text: 'Bitte zuerst unter Unterrichtsteams die Belegung laden – oder Manuell wählen'
            }
        ]);
        if (hint) {
            hint.innerHTML =
                'Verwandtes Tool: <a href="kursteams.html">Unterrichtsteams ↗</a>';
        }
        return;
    }

    const vorhanden = plans.filter((p) => p.status === 'ok').length;
    const abweichung = plans.filter((p) => p.status === 'sync').length;
    const neu = plans.filter((p) => p.status === 'neu').length;
    const zuWenig = plans.filter((p) => isTooFew(p)).length;
    const warn = zuWenig > 0 || abweichung > 0;
    if (banner) banner.className = 'pw-msg ' + (warn ? 'pw-msg--warn' : 'pw-msg--ok');
    status.textContent = deCount(plans.length, 'Klasse erkannt', 'Klassen erkannt');

    const items = [];
    if (neu) items.push({ ok: true, text: deCount(neu, 'kann neu angelegt werden', 'können neu angelegt werden') });
    if (vorhanden) {
        items.push({
            ok: true,
            text:
                vorhanden === 1
                    ? '1 bestehender KlassenChat gefunden'
                    : vorhanden + ' bestehende KlassenChats gefunden'
        });
    } else {
        items.push({ ok: true, text: 'Keine bestehenden KlassenChats gefunden' });
    }
    if (abweichung) {
        items.push({
            ok: false,
            text:
                abweichung === 1
                    ? '1 Name weicht vom Schema ab'
                    : abweichung + ' Namen weichen vom Schema ab'
        });
    }
    if (zuWenig) {
        items.push({
            ok: false,
            text:
                deCount(zuWenig, 'Klasse kann', 'Klassen können') +
                ' noch nicht angelegt werden – zu wenige Lehrkräfte'
        });
    } else if (plans.length) {
        items.push({ ok: true, text: 'Alle Klassen haben ausreichend Lehrkräfte' });
    }
    paintCheckList(checks, items);
    if (hint) {
        hint.textContent = 'Stand: ' + formatStand(belegung.updatedAt) + ' · Klassenvorstand wird ergänzt.';
    }
}

function refreshSourcePanel() {
    renderLists();
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
    plans = buildClassChatPlans(belegung, kv, {
        ...planOptsBase(),
        knownClassCodes: loadTenantClasses(),
        teacherDirectory: loadTenantTeachers()
    });
    return plans;
}

function statusLabel(p) {
    if (p && p._runStatus === 'created') return 'Erstellt';
    if (p && p._runStatus === 'updated') return 'Aktualisiert';
    if (p && p._runStatus === 'failed') return 'Fehlgeschlagen';
    if (isTooFew(p) || (p && !p.eligible) || p.status === 'skip') return 'Nicht anlegbar';
    if (p.status === 'neu') return 'Bereit';
    if (p.status === 'sync') return 'Prüfen';
    if (p.status === 'ok') return 'Bereits vorhanden';
    return p.status || 'Prüfen';
}

function statusPillHtml(p) {
    const label = statusLabel(p);
    let tone = '';
    if (isTooFew(p) || (p && !p.eligible) || p.status === 'skip' || p._runStatus === 'failed') {
        tone = 'fb-pill--bad';
    } else if (p.status === 'sync') tone = 'fb-pill--warn';
    else if (p.status === 'ok' || p._runStatus === 'updated') tone = 'fb-pill--info';
    else if (p.status === 'neu' || p._runStatus === 'created') tone = 'fb-pill--ok';
    return '<span class="fb-pill ' + tone + '">' + escapeHtml(label) + '</span>';
}

function shortPersonLabel(m) {
    if (!m) return '';
    const name = String(m.name || '').trim();
    if (name) {
        const parts = name.split(/\s+/).filter(Boolean);
        if (parts.length >= 2) {
            return parts[0] + ' ' + parts[parts.length - 1].charAt(0) + '.';
        }
        return name;
    }
    if (m.code) return String(m.code);
    const em = String(m.email || '');
    return em ? em.split('@')[0] : '';
}

function membersCellHtml(p) {
    const members = Array.isArray(p.members) ? p.members : [];
    const n = p.memberEmails ? p.memberEmails.length : members.length;
    const kv = members.find((m) => m && m.role === 'kv');
    const others = members.filter((m) => m && m.role !== 'kv');
    const show = others.slice(0, 2).map(shortPersonLabel).filter(Boolean);
    const rest = Math.max(0, others.length - show.length);
    const title = members
        .map((m) => {
            const tag = m.role === 'kv' ? ' (KV)' : m.role === 'extra' ? ' (zusätzlich)' : '';
            return (m.name || m.code || m.email) + tag;
        })
        .join(', ');
    let line = show.join(' · ');
    if (rest > 0) line += (line ? ' · ' : '') + '+' + rest;
    return (
        '<div class="pw-members" title="' +
        escapeHtml(title) +
        '">' +
        '<strong>' +
        n +
        ' Lehrkräfte</strong>' +
        (kv
            ? '<div>KV: ' + escapeHtml(shortPersonLabel(kv) || kv.email || '–') + '</div>'
            : '') +
        (line ? '<div>' + escapeHtml(line) + '</div>' : '') +
        '</div>'
    );
}

function planIsProblem(p) {
    return !!(p && (isTooFew(p) || !p.eligible || p.status === 'sync' || p.status === 'skip'));
}

function planRowHtml(p) {
    const tooFew = isTooFew(p);
    const key = String(p.klasse || '');
    const included = !tooFew && !!key && !p.omitted && p.eligible;
    const href = teamsHrefForPlan(p);
    const teamsCell = href
        ? '<a href="' +
          escapeHtml(href) +
          '" target="_blank" rel="noopener noreferrer">In Teams öffnen</a>'
        : '<span class="muted">–</span>';
    const rowClass = [
        tooFew || !p.eligible ? 'kc-row-skip' : '',
        planIsProblem(p) ? 'pw-row-problem' : ''
    ]
        .filter(Boolean)
        .join(' ');
    const disabledReason = tooFew
        ? p.skipReason || 'Zu wenige Lehrkräfte'
        : !p.eligible
          ? p.skipReason || 'Nicht anlegbar'
          : '';
    return (
        '<tr' +
        (rowClass ? ' class="' + rowClass + '"' : '') +
        '>' +
        '<td><input type="checkbox" data-kc-include="' +
        escapeHtml(key) +
        '"' +
        (included ? ' checked' : '') +
        (tooFew || !key || !p.eligible ? ' disabled' : '') +
        (disabledReason ? ' title="' + escapeHtml(disabledReason) + '"' : '') +
        ' aria-label="Klasse ' +
        escapeHtml(key || 'ohne Namen') +
        ' mit anlegen"></td>' +
        '<td><strong>' +
        escapeHtml(key || '–') +
        '</strong></td>' +
        '<td>' +
        escapeHtml(p.topic || '') +
        '</td>' +
        '<td>' +
        membersCellHtml(p) +
        '</td>' +
        '<td>' +
        statusPillHtml(p) +
        (p.omitted ? '<div class="muted">wird nicht angelegt</div>' : '') +
        (disabledReason ? '<div class="muted">' + escapeHtml(disabledReason) + '</div>' : '') +
        '</td>' +
        '<td>' +
        teamsCell +
        '</td>' +
        '</tr>'
    );
}

function planCardHtml(p) {
    const tooFew = isTooFew(p);
    const key = String(p.klasse || '');
    const included = !tooFew && !!key && !p.omitted && p.eligible;
    return (
        '<article class="pw-card' +
        (planIsProblem(p) ? ' pw-row-problem' : '') +
        '">' +
        '<label class="kc-checkbox" style="margin:0 0 8px;">' +
        '<input type="checkbox" data-kc-include="' +
        escapeHtml(key) +
        '"' +
        (included ? ' checked' : '') +
        (tooFew || !key || !p.eligible ? ' disabled' : '') +
        '>' +
        '<span><strong>' +
        escapeHtml(key || '–') +
        '</strong></span></label>' +
        '<div>' +
        escapeHtml(p.topic || '') +
        '</div>' +
        membersCellHtml(p) +
        '<div style="margin-top:8px;">' +
        statusPillHtml(p) +
        '</div></article>'
    );
}

function paintPlanTable(tbody) {
    if (!tbody) return;
    if (!plans.length) {
        tbody.innerHTML =
            '<tr><td colspan="6" class="kc-empty-cell">' +
            (mode === 'manual'
                ? 'Kein manueller Eintrag – Klasse und mindestens 2 Mitglieder angeben.'
                : 'Keine Klassen aus der Belegung – bitte Quelle prüfen oder Manuell wählen.') +
            '</td></tr>';
        return;
    }
    tbody.innerHTML = plans.map((p) => planRowHtml(p)).join('');
}

function paintPreviewCards() {
    const host = $('kcPreviewCards');
    if (!host) return;
    if (!plans.length) {
        host.innerHTML =
            '<p class="muted">' +
            (mode === 'manual'
                ? 'Kein manueller Eintrag.'
                : 'Keine Klassen aus der Belegung.') +
            '</p>';
        return;
    }
    host.innerHTML = plans.map((p) => planCardHtml(p)).join('');
}

function syncSelectAll() {
    const box = $('kcSelectAllPreview');
    if (!box) return;
    const boxes = Array.from(document.querySelectorAll('#kcPreviewBody [data-kc-include]:not(:disabled)'));
    if (!boxes.length) {
        box.checked = false;
        box.indeterminate = false;
        box.disabled = true;
        return;
    }
    box.disabled = false;
    const on = boxes.filter((b) => b.checked).length;
    box.checked = on === boxes.length;
    box.indeterminate = on > 0 && on < boxes.length;
}

function paintSummaries() {
    const ready = actionablePlans().length;
    const total = plans.length;
    const vorhanden = plans.filter((p) => p.status === 'ok').length;
    const abweichung = plans.filter((p) => p.status === 'sync').length;
    const zuWenig = plans.filter((p) => isTooFew(p)).length;
    const notReady = total - ready;
    const box = $('kcPreviewSummaryBox');
    const title = $('kcPreviewSummaryTitle');
    const list = $('kcPreviewSummary');
    const channel = targetKind === 'privateChannel';
    const noun = channel ? 'KlassenChats (Kanäle)' : 'KlassenChats';
    const hasWarn = zuWenig > 0 || abweichung > 0 || (notReady > 0 && ready < total);

    if (box) box.className = 'pw-msg ' + (hasWarn && ready < total ? 'pw-msg--warn' : 'pw-msg--ok');
    if (title) {
        title.textContent =
            ready === total && total > 0
                ? 'Bereit zum Anlegen'
                : ready > 0
                  ? ready + ' von ' + total + ' können angelegt werden'
                  : total
                    ? 'Noch nicht anlegbar'
                    : 'Keine Klassen in der Vorschau';
    }
    const items = [];
    if (ready) {
        items.push({
            ok: true,
            text:
                ready +
                ' ' +
                noun +
                (ready === 1 ? ' wird' : ' werden') +
                (vorhanden && ready === vorhanden ? ' abgeglichen' : ' neu angelegt')
        });
    }
    if (!abweichung) items.push({ ok: true, text: 'Keine Namenskonflikte' });
    else {
        items.push({
            ok: false,
            text: abweichung === 1 ? '1 Name weicht ab' : abweichung + ' Namen weichen ab'
        });
    }
    if (!zuWenig) items.push({ ok: true, text: 'Alle Klassen haben ausreichend Lehrkräfte' });
    else {
        items.push({
            ok: false,
            text:
                zuWenig === 1
                    ? '1 Klasse hat zu wenige Lehrkräfte'
                    : zuWenig + ' Klassen haben zu wenige Lehrkräfte'
        });
    }
    if (vorhanden) {
        items.push({
            ok: true,
            text:
                vorhanden === 1
                    ? '1 Chat existiert bereits'
                    : vorhanden + ' Chats existieren bereits'
        });
    }
    paintCheckList(list, items);
}

function namingSchemaLabel() {
    const yp = ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || yearPrefix || calcYearPrefix();
    const sample = displayNameForTarget(
        buildChatTopicFromPattern(getPatternFromBuilder(), { yearPrefix: yp, klasse: '{Klasse}' }),
        targetKind
    );
    return sample || (yp + ' | {Klasse} | Lehrer');
}

function paintConfirmPanel() {
    const ready = actionablePlans();
    const n = ready.length;
    const channel = targetKind === 'privateChannel';
    const headline = $('kcConfirmHeadline');
    if (headline) {
        headline.textContent = n
            ? n + ' KlassenChats werden erstellt'
            : 'Keine KlassenChats ausgewählt';
    }
    const target = $('kcConfirmTarget');
    if (target) {
        target.textContent =
            'Microsoft Teams · ' + (channel ? 'Privater Kanal' : 'Gruppenchat');
    }
    const classes = $('kcConfirmClasses');
    if (classes) {
        classes.textContent = n
            ? ready.map((p) => p.klasse).filter(Boolean).join(', ')
            : '–';
    }
    const members = $('kcConfirmMembers');
    if (members) {
        const extras = getExtraMemberEmails().length;
        members.textContent =
            (mode === 'manual' ? 'manuelle Auswahl' : 'aus Unterrichtsbelegung + Klassenvorstand') +
            (extras ? ' + ' + extras + ' zusätzliche Person(en)' : '');
    }
    const naming = $('kcConfirmNaming');
    if (naming) naming.textContent = namingSchemaLabel();
    const note = $('kcConfirmNote');
    if (note) {
        note.textContent = channel
            ? 'Es werden private Kanäle im Lehrerteam angelegt. Bestehende Kanäle werden nicht dupliziert.'
            : 'Es werden Gruppenchats in Microsoft Teams angelegt. Bestehende Chats werden nicht dupliziert.';
    }
}

function updateRunButton() {
    const btn = $('kcBtnRun');
    if (!btn) return;
    const n = actionablePlans().length;
    const channel = targetKind === 'privateChannel';
    const noun = n === 1 ? 'KlassenChat' : 'KlassenChats';
    const icon = channel ? 'bi-lock' : 'bi-chat-dots';
    const label = n ? n + ' ' + noun + ' anlegen' : 'Keine KlassenChats anlegen';
    btn.disabled = !n || busy || runPhase === 'exec';
    btn.className = 'fb-btn';
    btn.innerHTML = '<i class="bi ' + icon + '"></i>' + label;
}

function showStep4Phase(phase) {
    runPhase = phase;
    const confirm = $('kcConfirmPanel');
    const exec = $('kcExecPanel');
    const result = $('kcResultPanel');
    if (confirm) confirm.hidden = phase !== 'confirm';
    if (exec) exec.hidden = phase !== 'exec';
    if (result) result.hidden = phase !== 'result';
}

function renderLists() {
    belegung = loadBelegung();
    if (belegung && belegung.yearPrefix && $('kcYearPrefix') && !$('kcYearPrefix').dataset.userEdited) {
        $('kcYearPrefix').value = belegung.yearPrefix;
        yearPrefix = belegung.yearPrefix;
    }
    currentPlans();
    updateTargetUiCopy();
    paintSourceBanner();
    paintPlanTable($('kcPreviewBody'));
    paintPreviewCards();
    paintSummaries();
    syncSelectAll();
    updateNamePreview();
    updateRunButton();
    paintConfirmPanel();
    updateHero();
}

function renderPreview() {
    renderLists();
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
    const line = $('kcNameLine');
    const hint = $('kcNameHint');
    const examples = $('kcNameExamples');
    if (!line) return;
    const classList = plans
        .map((p) => String(p.klasse || '').trim())
        .filter(Boolean);
    const klasse =
        classList[0] ||
        (($('kcManualKlasse') && $('kcManualKlasse').value.trim()) || '');
    if (!klasse) {
        line.textContent = 'Noch keine Klasse';
        if (hint) hint.textContent = 'Der Name erscheint, sobald eine Belegung oder eine Klasse da ist.';
        if (examples) {
            examples.hidden = true;
            examples.innerHTML = '';
        }
        return;
    }
    const yp = ($('kcYearPrefix') && $('kcYearPrefix').value.trim()) || yearPrefix || calcYearPrefix();
    const pattern = getPatternFromBuilder();
    const preview = displayNameForTarget(
        buildChatTopicFromPattern(pattern, { yearPrefix: yp, klasse }),
        targetKind
    );
    line.textContent = preview || '–';
    if (hint) {
        hint.textContent =
            'Die weiteren Klassen folgen demselben Schema.' +
            (targetKind === 'privateChannel' ? ' Im Kanal wird | zu ·.' : '');
    }
    if (examples) {
        const more = classList.filter((c) => c !== klasse).slice(0, 2);
        if (!more.length) {
            examples.hidden = true;
            examples.innerHTML = '';
        } else {
            examples.hidden = false;
            examples.innerHTML = more
                .map((c) => {
                    const name = displayNameForTarget(
                        buildChatTopicFromPattern(pattern, { yearPrefix: yp, klasse: c }),
                        targetKind
                    );
                    return '<li>' + escapeHtml(name || c) + '</li>';
                })
                .join('');
        }
    }
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

function paintExecList(rows) {
    const list = $('kcExecList');
    const meta = $('kcExecMeta');
    if (!list) return;
    list.innerHTML = rows
        .map((r) => {
            let cls = 'is-wait';
            let mark = '○ Wartet';
            if (r.state === 'run') {
                cls = 'is-run';
                mark = '● Wird erstellt …';
            } else if (r.state === 'ok') {
                cls = 'is-ok';
                mark = '✓ ' + (r.created ? 'Erstellt' : 'Aktualisiert');
            } else if (r.state === 'fail') {
                cls = 'is-fail';
                mark = '✕ Fehlgeschlagen';
            }
            return (
                '<li class="' +
                cls +
                '"><span><strong>' +
                escapeHtml(r.klasse) +
                '</strong>' +
                (r.message
                    ? '<div class="muted">' + escapeHtml(r.message) + '</div>'
                    : '') +
                '</span><span>' +
                mark +
                '</span></li>'
            );
        })
        .join('');
    if (meta) {
        const done = rows.filter((r) => r.state === 'ok' || r.state === 'fail').length;
        meta.textContent = done + ' von ' + rows.length + ' abgeschlossen';
    }
}

function paintResultPanel(results) {
    const ok = results.filter((r) => r.ok).length;
    const fail = results.filter((r) => !r.ok).length;
    const title = $('kcResultTitle');
    const sub = $('kcResultSub');
    const list = $('kcResultList');
    const retry = $('kcBtnRetryFailed');
    if (title) {
        title.textContent =
            fail === 0
                ? 'Fertig ✓'
                : ok === 0
                  ? 'Nicht angelegt'
                  : ok + ' von ' + results.length + ' KlassenChats erstellt';
    }
    if (sub) {
        if (fail === 0) {
            sub.textContent =
                results.length === 1
                    ? '1 KlassenChat wurde erfolgreich angelegt.'
                    : results.length + ' KlassenChats wurden erfolgreich angelegt.';
        } else if (ok === 0) {
            sub.textContent = 'Keine KlassenChats konnten erstellt werden.';
        } else {
            sub.textContent =
                fail === 1
                    ? '1 KlassenChat konnte nicht erstellt werden.'
                    : fail + ' KlassenChats konnten nicht erstellt werden.';
        }
    }
    if (list) {
        list.innerHTML = results
            .map((r) => {
                return (
                    '<li><span><strong>' +
                    escapeHtml(r.klasse) +
                    '</strong><div class="muted">' +
                    escapeHtml(r.topic || '') +
                    (r.error ? ' · ' + escapeHtml(r.error) : '') +
                    '</div>' +
                    (r.detail
                        ? '<details class="kc-tech-details"><summary>Technische Details</summary><pre class="kc-log">' +
                          escapeHtml(r.detail) +
                          '</pre></details>'
                        : '') +
                    '</span><span>' +
                    (r.ok ? (r.created ? '✓ Erstellt' : '✓ Aktualisiert') : '✕ Fehlgeschlagen') +
                    '</span></li>'
                );
            })
            .join('');
    }
    if (retry) retry.hidden = fail === 0;
}

function userFacingProvisionError(e, klasse) {
    const detail = formatKlassenchatsGraphError(e);
    const lower = String(detail || '').toLowerCase();
    let reason = '';
    if (lower.indexOf('permission') !== -1 || lower.indexOf('403') !== -1 || lower.indexOf('denied') !== -1) {
        reason = 'Die Berechtigung für die Zielressource fehlt.';
    } else if (lower.indexOf('not found') !== -1 || lower.indexOf('404') !== -1) {
        reason = 'Die Zielressource wurde nicht gefunden.';
    } else if (lower.indexOf('throttl') !== -1 || lower.indexOf('429') !== -1) {
        reason = 'Microsoft Teams ist gerade überlastet. Bitte kurz warten und erneut versuchen.';
    }
    return {
        error: 'Der KlassenChat für ' + klasse + ' konnte nicht erstellt werden.' + (reason ? ' ' + reason : ''),
        detail
    };
}

async function runProvision(opts) {
    const onlyFailed = !!(opts && opts.onlyFailed);
    if (busy) return;
    rebuildPlans();
    applyOmissions();
    if (targetKind === 'privateChannel' && !(lehrerTeam && lehrerTeam.id)) {
        toast('Bitte zuerst ein Lehrerteam wählen (Schritt 1).');
        return;
    }
    let eligible = actionablePlans();
    if (onlyFailed) {
        const failedKeys = new Set(
            lastRunResults.filter((r) => !r.ok).map((r) => String(r.klasse || '').toUpperCase())
        );
        eligible = eligible.filter((p) => failedKeys.has(String(p.klasse || '').toUpperCase()));
    }
    if (!eligible.length) {
        toast(
            targetKind === 'privateChannel'
                ? 'Keine anlegbaren Kanäle – Team, Belegung oder Mitglieder prüfen.'
                : 'Keine anlegbaren Chats – Belegung, Formular oder Mitglieder prüfen.'
        );
        return;
    }

    busy = true;
    showStep4Phase('exec');
    updateRunButton();
    const back = $('kcBtnBackFromConfirm');
    if (back) back.disabled = true;
    const logEl = $('kcLog');
    if (logEl) logEl.textContent = '';

    /** @type {Array<{ klasse: string, topic: string, state: string, created?: boolean, message?: string }>} */
    const execRows = eligible.map((p) => ({
        klasse: p.klasse,
        topic: p.topic,
        state: 'wait'
    }));
    paintExecList(execRows);
    lastRunResults = [];

    try {
        log('Anmeldung bei Microsoft …');
        const token = await getToken();
        const me = await sharedGraphJson('GET', '/me?$select=id,mail,userPrincipalName', token, undefined);
        const meId = me && me.id ? String(me.id) : '';
        log(
            'Angemeldet · Ziel: ' +
                targetKindLabel(targetKind) +
                (lehrerTeam ? ' · Team: ' + lehrerTeam.displayName : '') +
                '. ' +
                eligible.length +
                ' KlassenChat(s) werden bearbeitet …'
        );

        let state = loadClassChatsState();

        if (targetKind === 'privateChannel' && lehrerTeam && lehrerTeam.id) {
            const allEmails = [];
            eligible.forEach((p) => {
                (p.memberEmails || []).forEach((em) => allEmails.push(em));
            });
            log('Lehrerteam vorbereiten: fehlende Lehrkräfte aufnehmen …');
            try {
                const prep = await prepareTeacherTeamForChannels(
                    G(),
                    token,
                    lehrerTeam.id,
                    allEmails,
                    meId
                );
                if (prep.added) {
                    log(
                        '  ' +
                            prep.added +
                            ' Person(en) ins Team aufgenommen – warte auf Teams-Sync …'
                    );
                } else {
                    log('  Alle geplanten Personen sind bereits im Team (oder schon bekannt).');
                }
                if (prep.missingEmails && prep.missingEmails.length) {
                    log('  Nicht in Graph gefunden: ' + prep.missingEmails.join(', '));
                }
                eligible.forEach((p) => {
                    p._teamRosterSettled = true;
                });
            } catch (e) {
                log(
                    '  Hinweis Team-Vorbereitung: ' +
                        formatKlassenchatsGraphError(e) +
                        ' – weiter mit Einzelversuchen …'
                );
            }
        }

        for (let i = 0; i < eligible.length; i++) {
            const plan = eligible[i];
            execRows[i].state = 'run';
            paintExecList(execRows);
            log('[' + (i + 1) + '/' + eligible.length + '] ' + plan.klasse + ' → ' + plan.topic);
            try {
                const res = await provisionClassResource(G(), token, plan, meId);
                state = upsertClassChatItem(state, {
                    klasse: plan.klasse,
                    kind: res.kind,
                    topic: res.topic,
                    chatId: res.chatId,
                    channelId: res.channelId,
                    teamId: res.teamId,
                    memberEmails: res.memberEmails,
                    webUrl: res.webUrl,
                    yearPrefix,
                    namePattern,
                    targetKind,
                    lehrerTeamId: lehrerTeam ? lehrerTeam.id : '',
                    lehrerTeamDisplayName: lehrerTeam ? lehrerTeam.displayName : ''
                });
                saveClassChatsState(state);
                execRows[i].state = 'ok';
                execRows[i].created = !!res.created;
                paintExecList(execRows);
                lastRunResults.push({
                    klasse: plan.klasse,
                    topic: res.topic || plan.topic,
                    ok: true,
                    created: !!res.created,
                    updated: !res.created
                });
                const bits = [];
                if (res.created) bits.push('angelegt');
                else bits.push('aktualisiert');
                if (res.topicUpdated) bits.push('Name geändert');
                if (res.added) bits.push('+' + res.added + ' Mitglied(er)');
                if (res.teamMembersAdded) bits.push('+' + res.teamMembersAdded + ' ins Team');
                if (res.missingEmails && res.missingEmails.length) {
                    bits.push('ohne Graph: ' + res.missingEmails.join(', '));
                }
                log('  OK – ' + bits.join(', '));
            } catch (e) {
                const faced = userFacingProvisionError(e, plan.klasse);
                execRows[i].state = 'fail';
                execRows[i].message = faced.error;
                paintExecList(execRows);
                lastRunResults.push({
                    klasse: plan.klasse,
                    topic: plan.topic,
                    ok: false,
                    error: faced.error,
                    detail: faced.detail
                });
                log('  Fehler: ' + faced.detail);
            }
        }

        persistPatternMeta(loadClassChatsState());
        const ok = lastRunResults.filter((r) => r.ok).length;
        const fail = lastRunResults.filter((r) => !r.ok).length;
        log('Fertig: ' + ok + ' ok, ' + fail + ' Fehler.');
        paintResultPanel(lastRunResults);
        showStep4Phase('result');
        toast(
            fail === 0
                ? ok + ' KlassenChats erfolgreich angelegt.'
                : ok + ' von ' + lastRunResults.length + ' KlassenChats erstellt.'
        );
        renderPreview();
    } catch (e) {
        log('Abbruch: ' + (e && e.message ? e.message : String(e)));
        toast('Anmeldung oder Graph fehlgeschlagen.');
        showStep4Phase('confirm');
    } finally {
        busy = false;
        if (back) back.disabled = false;
        updateRunButton();
    }
}

function boot() {
    const saved = loadClassChatsState();
    if (saved && saved.namePattern) namePattern = normalizeChatNamePattern(saved.namePattern);
    if (saved && saved.yearPrefix) yearPrefix = saved.yearPrefix;
    if (saved && saved.targetKind) targetKind = normalizeTargetKind(saved.targetKind);
    if (saved && saved.lehrerTeamId) {
        lehrerTeam = {
            id: String(saved.lehrerTeamId),
            displayName: String(saved.lehrerTeamDisplayName || ''),
            mail: ''
        };
    }
    if (saved && Array.isArray(saved.extraMemberEmails) && saved.extraMemberEmails.length) {
        const ta = $('kcExtraEmails');
        if (ta) ta.value = saved.extraMemberEmails.join('\n');
    }

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
    const nameAdjust = $('kcNameAdjust');
    if (nameAdjust) {
        const same =
            JSON.stringify(normalizeChatNamePattern(namePattern)) ===
            JSON.stringify(defaultChatNamePattern());
        if (!same) nameAdjust.open = true;
    }
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
    setTargetKind(targetKind);

    document.querySelectorAll('[data-kc-mode]').forEach((btn) => {
        btn.addEventListener('click', () => setMode(btn.getAttribute('data-kc-mode')));
    });
    document.querySelectorAll('[data-kc-target]').forEach((btn) => {
        btn.addEventListener('click', () => {
            setTargetKind(btn.getAttribute('data-kc-target'));
            persistPatternMeta(loadClassChatsState());
        });
    });

    const btnTeamSearch = $('kcBtnTeamSearch');
    if (btnTeamSearch) btnTeamSearch.addEventListener('click', () => runTeamSearch());
    const teamSearch = $('kcTeamSearch');
    if (teamSearch) {
        teamSearch.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                runTeamSearch();
            }
        });
    }

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
            renderLists();
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

    const btnExtraSearch = $('kcBtnExtraSearch');
    if (btnExtraSearch) btnExtraSearch.addEventListener('click', () => runExtraUserSearch());
    const extraSearch = $('kcExtraSearch');
    if (extraSearch) {
        extraSearch.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                e.preventDefault();
                runExtraUserSearch();
            }
        });
    }
    const btnExtraApply = $('kcBtnExtraApply');
    if (btnExtraApply) btnExtraApply.addEventListener('click', () => applyExtraSelection());
    const extraEmails = $('kcExtraEmails');
    if (extraEmails) {
        extraEmails.addEventListener('change', () => {
            persistPatternMeta(loadClassChatsState());
            if (step >= 3) renderPreview();
        });
    }
    renderExtraPicked();

    document.querySelectorAll('[data-kc-goto]').forEach((btn) => {
        btn.addEventListener('click', () => {
            if (btn.disabled) return;
            goToStep(btn.getAttribute('data-kc-goto'));
        });
    });
    document.querySelectorAll('[data-kc-next]').forEach((btn) => {
        btn.addEventListener('click', () => {
            const n = Number(btn.getAttribute('data-kc-next'));
            if (n >= 2) persistPatternMeta(loadClassChatsState());
            goToStep(n);
        });
    });
    document.querySelectorAll('[data-kc-prev]').forEach((btn) => {
        btn.addEventListener('click', () => goToStep(btn.getAttribute('data-kc-prev')));
    });

    function onIncludeChange(ev) {
        const box = ev.target && ev.target.closest ? ev.target.closest('[data-kc-include]') : null;
        if (!box) return;
        const key = String(box.getAttribute('data-kc-include') || '').toUpperCase();
        if (!key) return;
        if (box.checked) omittedKlassen.delete(key);
        else omittedKlassen.add(key);
        renderLists();
    }
    const previewBody = $('kcPreviewBody');
    if (previewBody) previewBody.addEventListener('change', onIncludeChange);
    const previewCards = $('kcPreviewCards');
    if (previewCards) previewCards.addEventListener('change', onIncludeChange);

    const selectAll = $('kcSelectAllPreview');
    if (selectAll) {
        selectAll.addEventListener('change', () => {
            const boxes = Array.from(
                document.querySelectorAll('#kcPreviewBody [data-kc-include]:not(:disabled)')
            );
            boxes.forEach((box) => {
                const key = String(box.getAttribute('data-kc-include') || '').toUpperCase();
                if (!key) return;
                if (selectAll.checked) omittedKlassen.delete(key);
                else omittedKlassen.add(key);
            });
            renderLists();
        });
    }

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

    const btnRetry = $('kcBtnRetryFailed');
    if (btnRetry) {
        btnRetry.addEventListener('click', () => {
            showStep4Phase('confirm');
            runProvision({ onlyFailed: true });
        });
    }
    const btnRestart = $('kcBtnRestart');
    if (btnRestart) {
        btnRestart.addEventListener('click', () => {
            lastRunResults = [];
            runPhase = 'confirm';
            showStep4Phase('confirm');
            goToStep(1);
        });
    }
    const btnShowPreview = $('kcBtnShowPreview');
    if (btnShowPreview) {
        btnShowPreview.addEventListener('click', () => {
            maxReached = Math.max(maxReached, 3);
            goToStep(3);
        });
    }

    showStep4Phase('confirm');
    setStep(1);
}

if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', boot);
} else {
    boot();
}
