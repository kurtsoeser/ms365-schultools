import { escapeHtml } from '../../shared/utils/strings.js';
/**
 * Spielwiesen-Wizard: Demo-Klasse, Lehrer-Bulk, Einzel-Team.
 */
import { createEducationClassTeam } from '../../shared/education-class-team.js';
import {
    MAX_DEMO_STUDENTS,
    DEMO_CLASS_CODE,
    buildSpielwiesenPlan,
    buildBulkTeacherPlans,
    buildDemoStudentPlan,
    isDemoStudentUser,
    filterDemoStudents,
    validateDemoPool,
    generateDemoPassword,
    defaultSpielNamePattern,
    normalizeSpielNamePattern,
    spielTokenLabel,
    buildSpielDisplayName,
    buildSpielMailNickname,
    loadSpielNamePattern,
    saveSpielNamePattern
} from './spielwiesen-logic.js';
import {
    licensesFromAssigned,
    summarizeLicenses,
    studentUserPlanSkuIds
} from '../../shared/graph-licenses.js';

function $(id) {
    return document.getElementById(id);
}

/** @type {ReturnType<typeof defaultSpielNamePattern>} */
let namePattern = defaultSpielNamePattern();

function toast(m) {
    if (typeof window.ms365ToastOrAlert === 'function') window.ms365ToastOrAlert(m);
    else window.alert(m);
}


function log(msg) {
    const el = $('spLog');
    if (!el) return;
    el.textContent += (el.textContent ? '\n' : '') + String(msg || '');
    el.scrollTop = el.scrollHeight;
}

const SCOPES = [
    'https://graph.microsoft.com/User.Read',
    'https://graph.microsoft.com/User.Read.All',
    'https://graph.microsoft.com/User.ReadWrite.All',
    'https://graph.microsoft.com/Group.ReadWrite.All',
    'https://graph.microsoft.com/Organization.Read.All',
    'https://graph.microsoft.com/Team.Create',
    'https://graph.microsoft.com/EduRoster.ReadWrite'
];

/** @type {'teacher'|'single'} */
let mode = 'teacher';
let step = 1;
let busy = false;
/** @type {'confirm'|'exec'|'result'} */
let runPhase = 'confirm';
/** @type {Array<{ label: string, ok: boolean, created?: boolean, skipped?: boolean, error?: string }>} */
let lastRunResults = [];
/** @type {Array<{ id: string, displayName: string, userPrincipalName: string, mailNickname?: string, assignedLicenses?: any[], selected: boolean }>} */
let demoPool = [];
/** @type {Array<{ code: string, name: string, email: string, selected: boolean }>} */
let teachers = [];
/** @type {Map<string, string>|null} */
let skuLookup = null;

function G() {
    const api = window.ms365GraphUnifiedGroups;
    if (!api) throw new Error('Graph nicht geladen.');
    return api;
}

async function getToken() {
    if (typeof window.ms365AuthAcquireTokenPopup === 'function') {
        return window.ms365AuthAcquireTokenPopup(SCOPES);
    }
    return G().getGraphToken();
}

const USER_SELECT =
    'id,displayName,givenName,surname,mail,userPrincipalName,mailNickname,department,jobTitle,assignedLicenses,accountEnabled,usageLocation';

async function graphJson(method, path, token, body, extraHeaders) {
    return G().graphJson(method, path, token, body, extraHeaders);
}

function poolKey() {
    return 'ms365-spielwiesen-demo-pool-v1';
}

function savePoolIds() {
    try {
        localStorage.setItem(
            poolKey(),
            JSON.stringify({
                ids: demoPool.map(function (u) {
                    return u.id;
                }),
                upns: demoPool.map(function (u) {
                    return u.userPrincipalName;
                })
            })
        );
    } catch {
        /* ignore */
    }
}

function getDomain() {
    const fromInput = String(($('spDomain') && $('spDomain').value) || '')
        .trim()
        .replace(/^@+/, '');
    if (fromInput) return fromInput.toLowerCase();
    if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
        return String(window.ms365GetSchoolDomainNoAt() || '')
            .replace(/^@+/, '')
            .toLowerCase();
    }
    try {
        const s = window.ms365TenantSettingsLoad && window.ms365TenantSettingsLoad();
        return String((s && s.domain) || '')
            .replace(/^@+/, '')
            .toLowerCase();
    } catch {
        return '';
    }
}

function getYear() {
    return String(($('spYear') && $('spYear').value) || '').trim() || String(new Date().getFullYear());
}

function asDemo() {
    return !!($('spAsDemo') && $('spAsDemo').checked);
}

function getLabel() {
    return String(($('spLabel') && $('spLabel').value) || '').trim();
}

function namingCtx(extra) {
    const base = {
        asDemo: asDemo(),
        year: getYear(),
        label: getLabel(),
        lehrer: '',
        lehrerName: ''
    };
    if (extra && typeof extra === 'object') {
        Object.keys(extra).forEach(function (k) {
            base[k] = extra[k];
        });
    }
    return base;
}

function planOpts(extra) {
    const o = {
        year: getYear(),
        asDemo: asDemo(),
        label: getLabel(),
        pattern: namePattern
    };
    if (extra && typeof extra === 'object') {
        Object.keys(extra).forEach(function (k) {
            o[k] = extra[k];
        });
    }
    return o;
}

function getPatternFromBuilder() {
    const zone = $('spNameBuilder');
    if (!zone) return normalizeSpielNamePattern(namePattern);
    const tokens = [];
    zone.querySelectorAll('[data-token-type]').forEach(function (el) {
        const type = String(el.getAttribute('data-token-type') || '');
        if (type === 'text') {
            tokens.push({ type: 'text', value: String(el.getAttribute('data-token-value') || '') });
        } else if (type) {
            tokens.push({ type: type });
        }
    });
    return normalizeSpielNamePattern(tokens);
}

function refreshNamePreview() {
    const el = $('spNamePreview');
    if (!el) return;
    const sampleTeacher = namingCtx({ lehrer: 'MU', lehrerName: 'Muster' });
    const sampleSingle = namingCtx({ lehrer: '', lehrerName: '' });
    const tName = buildSpielDisplayName(namePattern, sampleTeacher);
    const tNick = buildSpielMailNickname(namePattern, sampleTeacher);
    const sName = buildSpielDisplayName(namePattern, sampleSingle);
    const sNick = buildSpielMailNickname(namePattern, sampleSingle);
    el.innerHTML =
        '<p><strong>Lehrer-Bulk (Beispiel MU):</strong> ' +
        escapeHtml(tName) +
        ' · <code>' +
        escapeHtml(tNick) +
        '</code></p><p><strong>Einzel-Team:</strong> ' +
        escapeHtml(sName) +
        ' · <code>' +
        escapeHtml(sNick) +
        '</code></p>';
}

function addNameChip(zone, token) {
    const chip = document.createElement('span');
    chip.className = 'name-chip';
    chip.draggable = true;
    chip.setAttribute('data-token-type', token.type);
    if (token.type === 'text') chip.setAttribute('data-token-value', String(token.value ?? ''));

    const txt = document.createElement('span');
    if (token.type === 'text') {
        const v = String(token.value ?? '');
        txt.textContent = v === '' ? '(leer)' : v;
    } else {
        txt.textContent = spielTokenLabel(token);
    }

    const x = document.createElement('button');
    x.type = 'button';
    x.className = 'chip-x';
    x.textContent = '✕';
    x.title = 'Baustein entfernen';
    x.addEventListener('click', function () {
        chip.remove();
        namePattern = getPatternFromBuilder();
        onNamingChanged();
    });

    chip.append(txt, x);
    zone.appendChild(chip);
}

function renderNameBuilder() {
    const zone = $('spNameBuilder');
    if (!zone) return;
    zone.replaceChildren();
    normalizeSpielNamePattern(namePattern).forEach(function (t) {
        addNameChip(zone, t);
    });
    refreshNamePreview();
}

function wireNameBuilder() {
    const zone = $('spNameBuilder');
    if (!zone || zone.dataset.wired === '1') return;
    zone.dataset.wired = '1';

    let dragEl = null;
    zone.addEventListener('dragstart', function (e) {
        const target = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
        if (!target) return;
        dragEl = target;
        target.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
    });
    zone.addEventListener('dragend', function () {
        if (dragEl) dragEl.classList.remove('dragging');
        dragEl = null;
        namePattern = getPatternFromBuilder();
        onNamingChanged();
    });
    zone.addEventListener('dragover', function (e) {
        e.preventDefault();
        const over = e.target && e.target.closest ? e.target.closest('.name-chip') : null;
        if (!dragEl || !over || over === dragEl) return;
        const rect = over.getBoundingClientRect();
        const after = e.clientX > rect.left + rect.width / 2;
        if (after) over.after(dragEl);
        else over.before(dragEl);
    });
    zone.addEventListener('drop', function (e) {
        e.preventDefault();
        namePattern = getPatternFromBuilder();
        onNamingChanged();
    });

    document.querySelectorAll('[data-sp-token]').forEach(function (btn) {
        btn.addEventListener('click', function () {
            const type = String(btn.getAttribute('data-sp-token') || '').trim();
            if (!type) return;
            addNameChip(zone, { type: type });
            namePattern = getPatternFromBuilder();
            onNamingChanged();
        });
    });

    const btnSep = $('spNameAddSep');
    if (btnSep) {
        btnSep.addEventListener('click', function () {
            const v = ($('spNameSepValue') && $('spNameSepValue').value) || ' | ';
            addNameChip(zone, { type: 'text', value: String(v) });
            namePattern = getPatternFromBuilder();
            onNamingChanged();
        });
    }
    const btnText = $('spNameAddText');
    if (btnText) {
        btnText.addEventListener('click', function () {
            const v = ($('spNameTextValue') && $('spNameTextValue').value) || '';
            addNameChip(zone, { type: 'text', value: String(v) });
            namePattern = getPatternFromBuilder();
            onNamingChanged();
        });
    }
    const btnReset = $('spNameResetDefault');
    if (btnReset) {
        btnReset.addEventListener('click', function () {
            namePattern = defaultSpielNamePattern();
            renderNameBuilder();
            onNamingChanged();
        });
    }
    const btnSave = $('spNameSave');
    if (btnSave) {
        btnSave.addEventListener('click', function () {
            namePattern = getPatternFromBuilder();
            saveSpielNamePattern(namePattern);
            toast('Namensschema gespeichert.');
        });
    }
}

function onNamingChanged() {
    refreshNamePreview();
    refreshSinglePreview();
    renderTeachers();
    refreshRunSummary();
}

const HERO_BY_STEP = {
    1: {
        title: 'Spielwiesen anlegen',
        sub: 'Zweck und Namen wählen – danach Demo-Schüler, Empfänger und Anlegen.'
    },
    2: {
        title: 'Demo-Schüler vorbereiten',
        sub: 'Gemeinsamer Pool für alle Spielwiesen-Teams (max. 10 Konten).'
    },
    3: {
        title: 'Empfänger wählen',
        sub: 'Lehrkräfte oder ein Einzel-Team – Namen folgen dem Muster aus Schritt 1.'
    },
    4: {
        title: 'In Teams anlegen',
        sub: 'Kursteams (EDU_Class) erzeugen; bestehende Alias werden übersprungen.'
    }
};

function updateHero() {
    const conf = HERO_BY_STEP[step] || HERO_BY_STEP[1];
    const title = $('spHeroTitle');
    const sub = $('spHeroSub');
    if (title) title.textContent = conf.title;
    if (sub) sub.textContent = conf.sub;
}

function showStep4Phase(phase) {
    runPhase = phase;
    const confirm = $('spConfirmPanel');
    const exec = $('spExecPanel');
    const result = $('spResultPanel');
    if (confirm) confirm.hidden = phase !== 'confirm';
    if (exec) exec.hidden = phase !== 'exec';
    if (result) result.hidden = phase !== 'result';
}

function setStep(n) {
    step = Math.max(1, Math.min(4, Number(n) || 1));
    document.querySelectorAll('[data-sp-step]').forEach(function (el) {
        el.classList.toggle('is-active', Number(el.getAttribute('data-sp-step')) === step);
    });
    document.querySelectorAll('.pw-steps [data-sp-goto], .sp-steps [data-sp-goto]').forEach(function (el) {
        const num = Number(el.getAttribute('data-sp-goto'));
        const active = num === step;
        el.classList.toggle('is-active', active);
        el.classList.toggle('is-done', num < step);
        if (active) el.setAttribute('aria-current', 'step');
        else el.removeAttribute('aria-current');
    });
    updateHero();
    if (step === 3) refreshRecvUi();
    if (step === 4) {
        if (runPhase !== 'result' && runPhase !== 'exec') showStep4Phase('confirm');
        refreshRunSummary();
    }
    const heading = $('spStep' + step + 'Heading');
    if (heading && typeof heading.focus === 'function') {
        try {
            heading.focus({ preventScroll: false });
        } catch {
            heading.focus();
        }
    }
}

function setMode(m) {
    mode = m === 'single' ? 'single' : 'teacher';
    document.querySelectorAll('[data-sp-mode]').forEach(function (el) {
        const on = el.getAttribute('data-sp-mode') === mode;
        el.classList.toggle('is-selected', on);
        el.setAttribute('aria-checked', on ? 'true' : 'false');
    });
    refreshRecvUi();
    refreshSinglePreview();
    refreshRunSummary();
}

function refreshRecvUi() {
    const t = $('spRecvTeacher');
    const s = $('spRecvSingle');
    const hint = $('spRecvHint');
    if (t) t.hidden = mode !== 'teacher';
    if (s) s.hidden = mode !== 'single';
    if (hint) {
                hint.textContent =
            mode === 'teacher'
                ? 'Lehrkräfte aus den Stammdaten – ein Team pro Auswahl; Namen aus den Bausteinen in Schritt 1.'
                : 'Einzelnes Schilf-/Fach-Team – Name aus Bausteinen (Thema in Schritt 1).';
    }
}

function refreshSinglePreview() {
    const plan = buildSpielwiesenPlan(planOpts());
    const prev = $('spPreview');
    if (prev) {
        prev.innerHTML =
            '<p><strong>' +
            escapeHtml(plan.displayName) +
            '</strong></p><p><code>' +
            escapeHtml(plan.mailNickname) +
            '</code></p><p class="muted">' +
            escapeHtml(plan.description) +
            '</p>';
    }
    renderNotebook(plan.notebookChecklist);
    refreshNamePreview();
}

function renderNotebook(steps) {
    const list = $('spNotebookList');
    if (!list) return;
    list.replaceChildren();
    (steps || []).forEach(function (text, i) {
        const li = document.createElement('li');
        li.innerHTML =
            '<label style="display:flex;gap:8px;align-items:flex-start;"><input type="checkbox" data-sp-nb="' +
            i +
            '"><span>' +
            escapeHtml(text) +
            '</span></label>';
        list.appendChild(li);
    });
    restoreNb();
}

function nbKey() {
    return 'ms365-spielwiesen-notebook-v1';
}

function restoreNb() {
    let state = {};
    try {
        state = JSON.parse(localStorage.getItem(nbKey()) || '{}') || {};
    } catch {
        state = {};
    }
    document.querySelectorAll('[data-sp-nb]').forEach(function (el) {
        el.checked = !!state[el.getAttribute('data-sp-nb')];
        if (el.dataset.bound === '1') return;
        el.dataset.bound = '1';
        el.addEventListener('change', function () {
            const s = {};
            document.querySelectorAll('[data-sp-nb]').forEach(function (x) {
                s[x.getAttribute('data-sp-nb')] = !!x.checked;
            });
            try {
                localStorage.setItem(nbKey(), JSON.stringify(s));
            } catch {
                /* ignore */
            }
        });
    });
}

function licenseLabel(user) {
    const list = licensesFromAssigned(user && user.assignedLicenses, skuLookup);
    const sum = summarizeLicenses(list);
    if (sum.hasStudentUserPlan) {
        const a1 = list.find(function (l) {
            return l.audience === 'student' && l.family === 'a1';
        });
        return a1 ? a1.shortLabel : 'Schüler-Plan';
    }
    if (sum.hasAny) return list.map(function (l) {
        return l.shortLabel;
    }).join(', ') || 'Lizenz';
    return 'keine';
}

function selectedDemo() {
    return demoPool.filter(function (u) {
        return u.selected;
    });
}

function renderDemoTable() {
    const body = $('spDemoBody');
    const status = $('spDemoStatus');
    if (status) {
        const v = validateDemoPool(selectedDemo());
        status.textContent =
            'Pool: ' +
            demoPool.length +
            ' gefunden, ' +
            selectedDemo().length +
            ' ausgewählt (max. ' +
            MAX_DEMO_STUDENTS +
            ', Klasse ' +
            DEMO_CLASS_CODE +
            ')' +
            (v.ok ? '' : ' – ' + v.issues.join('; '));
    }
    if (!body) return;
    body.replaceChildren();
    if (!demoPool.length) {
        body.innerHTML = '<tr><td colspan="4" class="muted">Keine Demo-Schüler – suchen oder anlegen.</td></tr>';
        return;
    }
    demoPool.forEach(function (u) {
        const tr = document.createElement('tr');
        const tdC = document.createElement('td');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = !!u.selected;
        cb.addEventListener('change', function () {
            u.selected = cb.checked;
            const sel = selectedDemo();
            if (sel.length > MAX_DEMO_STUDENTS) {
                u.selected = false;
                cb.checked = false;
                toast('Maximal ' + MAX_DEMO_STUDENTS + ' Demo-Schüler.');
            }
            renderDemoTable();
            savePoolIds();
        });
        tdC.appendChild(cb);
        tr.appendChild(tdC);
        const tdN = document.createElement('td');
        tdN.textContent = u.displayName || '';
        tr.appendChild(tdN);
        const tdU = document.createElement('td');
        tdU.innerHTML = '<code style="font-size:0.85em;">' + escapeHtml(u.userPrincipalName || '') + '</code>';
        tr.appendChild(tdU);
        const tdL = document.createElement('td');
        tdL.textContent = licenseLabel(u);
        tdL.className = licenseLabel(u) === 'keine' ? 'muted' : '';
        tr.appendChild(tdL);
        body.appendChild(tr);
    });
}

function mergeDemoUsers(users) {
    const byId = new Map(
        demoPool.map(function (u) {
            return [u.id, u];
        })
    );
    (users || []).forEach(function (u) {
        if (!u || !u.id) return;
        const prev = byId.get(u.id);
        byId.set(u.id, {
            id: u.id,
            displayName: u.displayName || '',
            userPrincipalName: u.userPrincipalName || '',
            mailNickname: u.mailNickname || '',
            assignedLicenses: u.assignedLicenses || [],
            selected: prev ? prev.selected : true
        });
    });
    demoPool = Array.from(byId.values()).sort(function (a, b) {
        return String(a.displayName).localeCompare(String(b.displayName), 'de');
    });
    let selCount = 0;
    demoPool.forEach(function (u) {
        if (u.selected) {
            selCount++;
            if (selCount > MAX_DEMO_STUDENTS) u.selected = false;
        }
    });
    savePoolIds();
    renderDemoTable();
}

async function ensureSkuLookup(token) {
    if (skuLookup) return skuLookup;
    try {
        const res = await G().fetchSubscribedSkus(token);
        skuLookup = G().skuLookupFromSubscribed(res && res.skus);
    } catch {
        skuLookup = new Map();
    }
    return skuLookup;
}

async function refreshUserRow(token, id) {
    const path =
        '/users/' + encodeURIComponent(id) + '?$select=' + encodeURIComponent(USER_SELECT);
    return graphJson('GET', path, token);
}

async function findDemoStudents() {
    log('Suche Demo-Schüler …');
    const token = await getToken();
    await ensureSkuLookup(token);
    const found = [];
    const queries = ['DEMO Schüler', 'demo.schueler', 'DEMO'];
    for (let i = 0; i < queries.length; i++) {
        try {
            const hits = await G().searchUsers(token, queries[i]);
            hits.forEach(function (u) {
                found.push(u);
            });
        } catch (e) {
            log('Suche „' + queries[i] + '“: ' + (e.message || e));
        }
    }
    // Department DEMO
    try {
        const path =
            '/users?$filter=' +
            encodeURIComponent("department eq '" + DEMO_CLASS_CODE + "'") +
            '&$select=' +
            encodeURIComponent(USER_SELECT) +
            '&$top=25';
        const data = await graphJson('GET', path, token);
        ((data && data.value) || []).forEach(function (u) {
            found.push(u);
        });
    } catch {
        /* optional */
    }
    const demoOnly = filterDemoStudents(found);
    const map = new Map();
    for (let i = 0; i < demoOnly.length; i++) {
        const u = demoOnly[i];
        if (!u || !u.id || map.has(u.id)) continue;
        try {
            const fresh = await refreshUserRow(token, u.id);
            map.set(u.id, fresh);
        } catch {
            map.set(u.id, u);
        }
    }
    mergeDemoUsers(Array.from(map.values()));
    log(map.size + ' Demo-Schüler gefunden.');
    toast(map.size + ' Demo-Schüler gefunden.');
}

async function findStudentA1SkuId(token) {
    await ensureSkuLookup(token);
    const res = await G().fetchSubscribedSkus(token);
    const list = (res && res.skus) || (Array.isArray(res) ? res : []) || [];
    const studentIds = new Set(studentUserPlanSkuIds().map(function (id) {
        return String(id).toLowerCase();
    }));
    let a1 = null;
    let anyStudent = null;
    list.forEach(function (s) {
        const id = String((s && s.skuId) || '').toLowerCase();
        if (!id || !studentIds.has(id)) return;
        const part = String((s && s.skuPartNumber) || '').toUpperCase();
        const info = licensesFromAssigned([{ skuId: id, skuPartNumber: part }], skuLookup)[0];
        if (!anyStudent) anyStudent = id;
        if (info && info.family === 'a1') a1 = id;
    });
    return a1 || anyStudent;
}

async function assignStudentLicense(token, userId) {
    const skuId = await findStudentA1SkuId(token);
    if (!skuId) {
        log('Keine Schüler-SKU (A1/A3/A5) im Tenant gefunden – Lizenz übersprungen.');
        return false;
    }
    try {
        await G().graphJson('PATCH', '/users/' + encodeURIComponent(userId), token, {
            usageLocation: 'AT'
        });
    } catch (e) {
        log('usageLocation: ' + (e.message || e));
    }
    await G().graphJson('POST', '/users/' + encodeURIComponent(userId) + '/assignLicense', token, {
        addLicenses: [{ skuId: skuId, disabledPlans: [] }],
        removeLicenses: []
    });
    return true;
}

async function ensureDemoStudents() {
    const domain = getDomain();
    if (!domain) throw new Error('Schul-Domain fehlt (Schritt 2).');
    const count = Math.max(1, Math.min(MAX_DEMO_STUDENTS, Number(($('spDemoCount') && $('spDemoCount').value) || 5)));
    let password = String(($('spDemoPassword') && $('spDemoPassword').value) || '').trim();
    if (!password) {
        password = generateDemoPassword();
        if ($('spDemoPassword')) $('spDemoPassword').value = password;
    }
    const assignA1 = !!($('spAssignA1') && $('spAssignA1').checked);
    const token = await getToken();
    await ensureSkuLookup(token);

    if (!demoPool.length) {
        try {
            await findDemoStudents();
        } catch {
            /* continue create */
        }
    }

    const byUpn = new Map(
        demoPool.map(function (u) {
            return [String(u.userPrincipalName || '').toLowerCase(), u];
        })
    );

    for (let i = 1; i <= count; i++) {
        const plan = buildDemoStudentPlan({ index: i, domain: domain });
        if (!plan.ok) throw new Error(plan.issues.join(', '));
        const existing = byUpn.get(plan.userPrincipalName.toLowerCase());
        if (existing) {
            log('Vorhanden: ' + plan.userPrincipalName);
            continue;
        }
        log('Lege an: ' + plan.userPrincipalName + ' …');
        try {
            const body = {
                accountEnabled: true,
                displayName: plan.displayName,
                givenName: plan.givenName,
                surname: plan.surname,
                mailNickname: plan.mailNickname,
                userPrincipalName: plan.userPrincipalName,
                department: plan.department,
                jobTitle: plan.jobTitle,
                usageLocation: 'AT',
                passwordProfile: {
                    password: password,
                    forceChangePasswordNextSignIn: true
                }
            };
            const created = await graphJson('POST', '/users', token, body);
            if (assignA1 && created && created.id) {
                try {
                    await assignStudentLicense(token, created.id);
                    log('  → A1/Schüler-Lizenz zugewiesen.');
                } catch (e) {
                    log('  → Lizenz: ' + (e.message || e));
                }
            }
            try {
                const fresh = created && created.id ? await refreshUserRow(token, created.id) : created;
                mergeDemoUsers([fresh || created]);
            } catch {
                mergeDemoUsers([created]);
            }
        } catch (e) {
            const msg = String(e.message || e);
            if (/already exists|ObjectConflict|another object/i.test(msg)) {
                log('  → existiert bereits, suche …');
                try {
                    const hits = await G().searchUsers(token, plan.userPrincipalName);
                    mergeDemoUsers(hits.filter(isDemoStudentUser));
                } catch {
                    /* ignore */
                }
            } else {
                log('  → FEHLER: ' + msg);
            }
        }
        await G().sleep(200);
    }
    toast('Demo-Schüler-Pool aktualisiert.');
    renderDemoTable();
}

function loadTeachers() {
    let list = [];
    try {
        const s = window.ms365TenantSettingsLoad && window.ms365TenantSettingsLoad();
        list = Array.isArray(s && s.teachers) ? s.teachers : [];
    } catch {
        list = [];
    }
    const prev = new Map(
        teachers.map(function (t) {
            return [t.code, t.selected];
        })
    );
    teachers = list
        .map(function (t) {
            const code = String((t && t.code) || '')
                .trim()
                .toUpperCase();
            if (!code) return null;
            return {
                code: code,
                name: String((t && t.name) || '').trim(),
                email: String((t && t.email) || '')
                    .trim()
                    .toLowerCase(),
                selected: prev.has(code) ? prev.get(code) : !!(t && t.email)
            };
        })
        .filter(Boolean)
        .sort(function (a, b) {
            return a.code.localeCompare(b.code, 'de');
        });
    renderTeachers();
    toast(teachers.length + ' Lehrkräfte geladen.');
}

function renderTeachers() {
    const body = $('spTeacherBody');
    if (!body) return;
    body.replaceChildren();
    if (!teachers.length) {
        body.innerHTML =
            '<tr><td colspan="5" class="muted">Keine Lehrkräfte in den Stammdaten – unter Einrichtung/Schul-Einstellungen pflegen.</td></tr>';
        return;
    }
    const bulk = buildBulkTeacherPlans(
        planOpts({
            teachers: teachers,
            selectedCodes: teachers.map(function (t) {
                return t.code;
            })
        })
    );
    const planByCode = new Map(
        bulk.plans.map(function (p) {
            return [p.code, p];
        })
    );
    teachers.forEach(function (t) {
        const tr = document.createElement('tr');
        const tdC = document.createElement('td');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = !!t.selected;
        cb.disabled = !t.email;
        cb.addEventListener('change', function () {
            t.selected = cb.checked;
            refreshRunSummary();
        });
        tdC.appendChild(cb);
        tr.appendChild(tdC);
        const tdK = document.createElement('td');
        tdK.textContent = t.code;
        tr.appendChild(tdK);
        const tdN = document.createElement('td');
        tdN.textContent = t.name || '—';
        tr.appendChild(tdN);
        const tdE = document.createElement('td');
        tdE.textContent = t.email || '(E-Mail fehlt)';
        if (!t.email) tdE.className = 'muted';
        tr.appendChild(tdE);
        const tdP = document.createElement('td');
        const p = planByCode.get(t.code);
        tdP.innerHTML = p
            ? '<span style="font-size:0.85em;">' +
              escapeHtml(p.displayName) +
              '</span><br><code style="font-size:0.75em;">' +
              escapeHtml(p.mailNickname) +
              '</code>'
            : '';
        tr.appendChild(tdP);
        body.appendChild(tr);
    });
}

function refreshRunSummary() {
    const demos = selectedDemo();
    const headline = $('spConfirmHeadline');
    const teams = $('spConfirmTeams');
    const demosEl = $('spConfirmDemos');
    const naming = $('spConfirmNaming');
    const btn = $('spBtnRun');

    if (mode === 'single') {
        const plan = buildSpielwiesenPlan(planOpts());
        if (headline) headline.textContent = '1 Spielwiesen-Team wird erstellt';
        if (teams) teams.textContent = plan.displayName || '–';
        if (demosEl) {
            demosEl.textContent =
                demos.length + ' Demo-Schüler (Klasse ' + DEMO_CLASS_CODE + ')';
        }
        if (naming) naming.textContent = plan.displayName || '–';
        if (btn) {
            btn.disabled = busy || !plan.ok;
            btn.innerHTML = '<i class="bi bi-play-fill"></i>1 Spielwiese anlegen';
        }
        return;
    }

    const selected = teachers.filter(function (t) {
        return t.selected && t.email;
    });
    const bulk = buildBulkTeacherPlans(
        planOpts({
            teachers: selected,
            selectedCodes: selected.map(function (t) {
                return t.code;
            })
        })
    );
    const n = bulk.plans.length;
    const sample = bulk.plans[0];
    if (headline) {
        headline.textContent = n
            ? n + ' Spielwiesen-Teams werden erstellt'
            : 'Keine Lehrkräfte ausgewählt';
    }
    if (teams) {
        teams.textContent = n
            ? selected
                  .slice(0, 8)
                  .map(function (t) {
                      return t.code || t.name;
                  })
                  .join(', ') + (selected.length > 8 ? ' …' : '')
            : '–';
    }
    if (demosEl) {
        demosEl.textContent =
            demos.length + ' Demo-Schüler (Klasse ' + DEMO_CLASS_CODE + ')';
    }
    if (naming) {
        naming.textContent = sample
            ? sample.displayName.replace(String(sample.code || ''), '{Kürzel}')
            : 'DEMO | {Thema} | {Kürzel}';
    }
    if (btn) {
        btn.disabled = busy || !n || !bulk.ok;
        btn.innerHTML =
            '<i class="bi bi-play-fill"></i>' +
            (n ? n + ' Spielwiesen anlegen' : 'Keine Spielwiesen anlegen');
    }
}

function paintExecList(rows) {
    const list = $('spExecList');
    const meta = $('spExecMeta');
    if (!list) return;
    list.innerHTML = rows
        .map(function (r) {
            let cls = 'is-wait';
            let mark = '○ Wartet';
            if (r.state === 'run') {
                cls = 'is-run';
                mark = '● Wird erstellt …';
            } else if (r.state === 'ok') {
                cls = 'is-ok';
                mark = r.skipped ? '✓ Bereits vorhanden' : '✓ Erstellt';
            } else if (r.state === 'fail') {
                cls = 'is-fail';
                mark = '✕ Fehlgeschlagen';
            }
            return (
                '<li class="' +
                cls +
                '"><span><strong>' +
                escapeHtml(r.label) +
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
        const done = rows.filter(function (r) {
            return r.state === 'ok' || r.state === 'fail';
        }).length;
        meta.textContent = done + ' von ' + rows.length + ' abgeschlossen';
    }
}

function paintResultPanel(results) {
    const ok = results.filter(function (r) {
        return r.ok;
    }).length;
    const fail = results.filter(function (r) {
        return !r.ok;
    }).length;
    const title = $('spResultTitle');
    const sub = $('spResultSub');
    const list = $('spResultList');
    const retry = $('spBtnRetryFailed');
    if (title) {
        title.textContent =
            fail === 0 ? 'Fertig ✓' : ok === 0 ? 'Nicht angelegt' : ok + ' von ' + results.length + ' Teams erstellt';
    }
    if (sub) {
        if (fail === 0) {
            sub.textContent =
                results.length === 1
                    ? '1 Spielwiese wurde erfolgreich angelegt.'
                    : results.length + ' Spielwiesen wurden erfolgreich angelegt.';
        } else if (ok === 0) {
            sub.textContent = 'Keine Spielwiesen konnten erstellt werden.';
        } else {
            sub.textContent =
                fail === 1
                    ? '1 Team konnte nicht erstellt werden.'
                    : fail + ' Teams konnten nicht erstellt werden.';
        }
    }
    if (list) {
        list.innerHTML = results
            .map(function (r) {
                return (
                    '<li><span><strong>' +
                    escapeHtml(r.label) +
                    '</strong>' +
                    (r.error
                        ? '<div class="muted">' + escapeHtml(r.error) + '</div>'
                        : '') +
                    '</span><span>' +
                    (r.ok
                        ? r.skipped
                            ? '✓ Bereits vorhanden'
                            : '✓ Erstellt'
                        : '✕ Fehlgeschlagen') +
                    '</span></li>'
                );
            })
            .join('');
    }
    if (retry) retry.hidden = fail === 0;
}

async function findGroupByNickname(token, nick) {
    const esc = String(nick || '').replace(/'/g, "''");
    const path =
        '/groups?$filter=' +
        encodeURIComponent("mailNickname eq '" + esc + "'") +
        '&$select=id,displayName,mailNickname&$top=5';
    const data = await graphJson('GET', path, token);
    const list = (data && data.value) || [];
    return list[0] || null;
}

async function createSpielTeam(token, plan, ownerUserId) {
    let group = await findGroupByNickname(token, plan.mailNickname);
    if (group) {
        log('Übersprungen (existiert): ' + plan.mailNickname);
        return { group: group, created: false };
    }

    let ownerId = ownerUserId;
    if (!ownerId) {
        const me = await graphJson('GET', '/me?$select=id,userPrincipalName', token);
        ownerId = me && me.id ? me.id : '';
        if (ownerId) log('Besitzer: angemeldeter Benutzer (' + (me.userPrincipalName || me.id) + ')');
    }
    if (!ownerId) throw new Error('Kein Besitzer – Kursteam (EDU_Class) braucht mind. einen Owner.');

    log('Lege Kursteam an (educationClass / EDU_Class): ' + plan.displayName + ' …');
    const created = await createEducationClassTeam({
        graphJson: function (method, path, tok, body) {
            return G().graphJson(method, path, tok, body);
        },
        graphRequest: function (method, path, tok, body) {
            return G().graphRequest(method, path, tok, body);
        },
        sleep: function (ms) {
            return G().sleep(ms);
        },
        getToken: getToken,
        token: token,
        log: log,
        displayName: plan.displayName,
        mailNickname: plan.mailNickname,
        description: plan.description,
        classCode: plan.mailNickname || plan.displayName,
        ownerId: ownerId
    });
    log('Methode: ' + created.method + ' → ' + created.groupId);
    return {
        group: { id: created.groupId, displayName: plan.displayName, mailNickname: plan.mailNickname },
        created: true,
        method: created.method
    };
}

async function addDemoMembers(token, groupId) {
    if (!($('spAddDemoMembers') && $('spAddDemoMembers').checked)) return;
    const demos = selectedDemo();
    for (let i = 0; i < demos.length; i++) {
        try {
            await G().graphAddMember(token, groupId, demos[i].id);
        } catch (e) {
            if (G().isDuplicateMemberError && G().isDuplicateMemberError(e)) continue;
            log('  Mitglied ' + demos[i].userPrincipalName + ': ' + (e.message || e));
        }
        await G().sleep(120);
    }
    log('  Demo-Schüler: ' + demos.length + ' Mitgliedschaft(en) versucht.');
}

async function runCreate(opts) {
    const onlyFailed = !!(opts && opts.onlyFailed);
    if (busy) return;

    const demos = selectedDemo();
    const addMembers = !!($('spAddDemoMembers') && $('spAddDemoMembers').checked);
    if (addMembers) {
        const v = validateDemoPool(demos);
        if (!v.ok) throw new Error(v.issues.join(', '));
    }

    /** @type {Array<{ label: string, plan: any, email?: string }>} */
    let jobs = [];
    if (mode === 'single') {
        const plan = buildSpielwiesenPlan(planOpts());
        if (!plan.ok) throw new Error(plan.issues.join(', '));
        jobs = [{ label: plan.displayName, plan: plan, email: null }];
    } else {
        const selected = teachers.filter(function (t) {
            return t.selected && t.email;
        });
        const bulk = buildBulkTeacherPlans(
            planOpts({
                teachers: selected,
                selectedCodes: selected.map(function (t) {
                    return t.code;
                })
            })
        );
        if (!bulk.ok) throw new Error(bulk.issues.join('; '));
        jobs = bulk.plans.map(function (plan) {
            return {
                label: plan.displayName || plan.code,
                plan: plan,
                email: plan.email
            };
        });
    }

    if (onlyFailed) {
        const failed = new Set(
            lastRunResults
                .filter(function (r) {
                    return !r.ok;
                })
                .map(function (r) {
                    return String(r.label || '');
                })
        );
        jobs = jobs.filter(function (j) {
            return failed.has(j.label);
        });
    }
    if (!jobs.length) throw new Error('Keine Spielwiesen zum Anlegen ausgewählt.');

    busy = true;
    showStep4Phase('exec');
    refreshRunSummary();
    const back = $('spBtnBackFromConfirm');
    if (back) back.disabled = true;
    const logEl = $('spLog');
    if (logEl) logEl.textContent = '';

    const execRows = jobs.map(function (j) {
        return { label: j.label, state: 'wait' };
    });
    paintExecList(execRows);
    lastRunResults = [];

    try {
        const token = await getToken();
        for (let i = 0; i < jobs.length; i++) {
            const job = jobs[i];
            execRows[i].state = 'run';
            paintExecList(execRows);
            try {
                let ownerId = null;
                if (job.email) {
                    try {
                        const owner = await G().resolveUserByEmail(token, job.email);
                        ownerId = owner && owner.id ? owner.id : null;
                    } catch (e) {
                        log(
                            job.plan.code +
                                ': Besitzer nicht gefunden (' +
                                job.email +
                                ') – ' +
                                (e.message || e)
                        );
                    }
                    if (!ownerId) {
                        throw new Error('Besitzer nicht gefunden: ' + job.email);
                    }
                }
                const res = await createSpielTeam(token, job.plan, ownerId);
                if (res.group && res.group.id) await addDemoMembers(token, res.group.id);
                execRows[i].state = 'ok';
                execRows[i].skipped = !res.created;
                paintExecList(execRows);
                lastRunResults.push({
                    label: job.label,
                    ok: true,
                    created: !!res.created,
                    skipped: !res.created
                });
            } catch (e) {
                const msg = e && e.message ? e.message : String(e);
                execRows[i].state = 'fail';
                execRows[i].message = 'Das Team konnte nicht erstellt werden.';
                paintExecList(execRows);
                lastRunResults.push({
                    label: job.label,
                    ok: false,
                    error: msg
                });
                log(job.label + ' FEHLER: ' + msg);
            }
            await G().sleep(400);
        }

        const ok = lastRunResults.filter(function (r) {
            return r.ok;
        }).length;
        const fail = lastRunResults.filter(function (r) {
            return !r.ok;
        }).length;
        const skip = lastRunResults.filter(function (r) {
            return r.ok && r.skipped;
        }).length;
        log('Fertig: ' + ok + ' ok (' + skip + ' übersprungen), ' + fail + ' Fehler.');
        paintResultPanel(lastRunResults);
        showStep4Phase('result');
        toast(
            fail === 0
                ? ok + ' Spielwiesen erfolgreich.'
                : ok + ' von ' + lastRunResults.length + ' Spielwiesen erstellt.'
        );
    } catch (e) {
        showStep4Phase('confirm');
        throw e;
    } finally {
        busy = false;
        if (back) back.disabled = false;
        refreshRunSummary();
    }
}

function fillDefaults() {
    if ($('spYear') && !$('spYear').value) $('spYear').value = String(new Date().getFullYear());
    if ($('spDomain') && !$('spDomain').value) {
        const d = getDomain();
        if (d) $('spDomain').value = d;
    }
    if ($('spDemoPassword') && !$('spDemoPassword').value) {
        $('spDemoPassword').value = generateDemoPassword();
    }
}

function boot() {
    fillDefaults();
    namePattern = loadSpielNamePattern();
    wireNameBuilder();
    renderNameBuilder();
    setMode('teacher');
    showStep4Phase('confirm');
    setStep(1);
    refreshSinglePreview();

    document.querySelectorAll('[data-sp-mode]').forEach(function (el) {
        el.addEventListener('click', function () {
            setMode(el.getAttribute('data-sp-mode'));
        });
    });
    document.querySelectorAll('[data-sp-goto]').forEach(function (el) {
        el.addEventListener('click', function () {
            if (el.id === 'spBtnBackFromConfirm' && busy) return;
            setStep(el.getAttribute('data-sp-goto'));
        });
    });
    document.querySelectorAll('[data-sp-next]').forEach(function (el) {
        el.addEventListener('click', function () {
            setStep(el.getAttribute('data-sp-next'));
        });
    });

    const btnRetry = $('spBtnRetryFailed');
    if (btnRetry) {
        btnRetry.addEventListener('click', function () {
            Promise.resolve(runCreate({ onlyFailed: true })).catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    }
    const btnRestart = $('spBtnRestart');
    if (btnRestart) {
        btnRestart.addEventListener('click', function () {
            lastRunResults = [];
            showStep4Phase('confirm');
            setStep(1);
        });
    }

    ['spLabel', 'spYear', 'spAsDemo'].forEach(function (id) {
        const el = $(id);
        if (!el) return;
        el.addEventListener('input', onNamingChanged);
        el.addEventListener('change', onNamingChanged);
    });

    const bind = [
        ['spBtnFindDemo', findDemoStudents],
        ['spBtnEnsureDemo', ensureDemoStudents],
        [
            'spBtnLoadTeachers',
            function () {
                loadTeachers();
                return Promise.resolve();
            }
        ],
        [
            'spBtnSelAll',
            function () {
                teachers.forEach(function (t) {
                    if (t.email) t.selected = true;
                });
                renderTeachers();
                refreshRunSummary();
                return Promise.resolve();
            }
        ],
        [
            'spBtnSelNone',
            function () {
                teachers.forEach(function (t) {
                    t.selected = false;
                });
                renderTeachers();
                refreshRunSummary();
                return Promise.resolve();
            }
        ],
        ['spBtnRun', runCreate]
    ];
    bind.forEach(function (pair) {
        const btn = $(pair[0]);
        if (!btn || btn.dataset.bound === '1') return;
        btn.dataset.bound = '1';
        btn.addEventListener('click', function () {
            Promise.resolve(pair[1]()).catch(function (e) {
                log('FEHLER: ' + (e.message || e));
                toast(e.message || String(e));
            });
        });
    });

    // Lehrkräfte vorladen wenn Stammdaten da
    try {
        if (typeof window.ms365TenantSettingsLoad === 'function') loadTeachers();
    } catch {
        /* ignore */
    }
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();
