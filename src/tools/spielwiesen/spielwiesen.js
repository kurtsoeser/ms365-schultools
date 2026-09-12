/**
 * Spielwiesen-Wizard: Demo-Klasse, Lehrer-Bulk, Einzel-Team.
 */
import {
    MAX_DEMO_STUDENTS,
    DEMO_CLASS_CODE,
    buildSpielwiesenPlan,
    buildBulkTeacherPlans,
    buildDemoStudentPlan,
    isDemoStudentUser,
    filterDemoStudents,
    validateDemoPool,
    generateDemoPassword
} from './spielwiesen-logic.js';
import {
    licensesFromAssigned,
    summarizeLicenses,
    studentUserPlanSkuIds
} from '../../shared/graph-licenses.js';

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
    'https://graph.microsoft.com/Organization.Read.All'
];

/** @type {'teacher'|'single'} */
let mode = 'teacher';
let step = 1;
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

function setStep(n) {
    step = Math.max(1, Math.min(4, Number(n) || 1));
    document.querySelectorAll('[data-sp-step]').forEach(function (el) {
        el.classList.toggle('is-active', Number(el.getAttribute('data-sp-step')) === step);
    });
    document.querySelectorAll('[data-sp-goto]').forEach(function (el) {
        el.classList.toggle('is-active', Number(el.getAttribute('data-sp-goto')) === step);
    });
    if (step === 3) refreshRecvUi();
    if (step === 4) refreshRunSummary();
}

function setMode(m) {
    mode = m === 'single' ? 'single' : 'teacher';
    document.querySelectorAll('[data-sp-mode]').forEach(function (el) {
        el.classList.toggle('is-selected', el.getAttribute('data-sp-mode') === mode);
    });
    refreshRecvUi();
    refreshSinglePreview();
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
                ? 'Lehrkräfte aus den Stammdaten – ein Team pro Auswahl, gemeinsame Demo-Schüler.'
                : 'Einzelnes Schilf-/Fach-Team benennen (Alias spiel-{jahr}-{slug}).';
    }
}

function refreshSinglePreview() {
    const plan = buildSpielwiesenPlan({
        label: ($('spLabel') && $('spLabel').value) || '',
        year: getYear(),
        asDemo: asDemo()
    });
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
    const year = getYear();
    const bulk = buildBulkTeacherPlans({
        teachers: teachers,
        year: year,
        asDemo: asDemo(),
        selectedCodes: teachers.map(function (t) {
            return t.code;
        })
    });
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
            ? '<code style="font-size:0.8em;">' + escapeHtml(p.mailNickname) + '</code>'
            : '';
        tr.appendChild(tdP);
        body.appendChild(tr);
    });
}

function refreshRunSummary() {
    const el = $('spRunSummary');
    if (!el) return;
    const demos = selectedDemo();
    if (mode === 'single') {
        const plan = buildSpielwiesenPlan({
            label: ($('spLabel') && $('spLabel').value) || '',
            year: getYear(),
            asDemo: asDemo()
        });
        el.innerHTML =
            '<p><strong>Einzel-Team:</strong> ' +
            escapeHtml(plan.displayName) +
            '</p><p><code>' +
            escapeHtml(plan.mailNickname) +
            '</code></p><p>Demo-Schüler: ' +
            demos.length +
            '</p>';
        return;
    }
    const selected = teachers.filter(function (t) {
        return t.selected && t.email;
    });
    const bulk = buildBulkTeacherPlans({
        teachers: selected,
        year: getYear(),
        asDemo: asDemo(),
        selectedCodes: selected.map(function (t) {
            return t.code;
        })
    });
    el.innerHTML =
        '<p><strong>Lehrer-Spielwiesen:</strong> ' +
        bulk.plans.length +
        ' Team(s), Jahr ' +
        escapeHtml(getYear()) +
        '</p><p>Demo-Schüler im Pool: ' +
        demos.length +
        ' (Klasse ' +
        DEMO_CLASS_CODE +
        ')</p>' +
        (bulk.ok
            ? ''
            : '<p class="muted">' + escapeHtml(bulk.issues.join('; ')) + '</p>');
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
    log('Lege Team an: ' + plan.displayName + ' …');
    group = await G().createUnifiedGroup(token, plan.displayName, plan.mailNickname, plan.description);
    try {
        await G().provisionTeamForGroup(token, group.id);
    } catch (e) {
        log('  Teams-Provision: ' + (e.message || e) + ' (Gruppe ist trotzdem da)');
    }
    if (ownerUserId) {
        try {
            if (typeof G().addOwnerWithMemberFallback === 'function') {
                await G().addOwnerWithMemberFallback(token, group.id, ownerUserId);
            } else {
                await G().addGroupOwner(token, group.id, ownerUserId);
            }
            log('  Besitzer gesetzt.');
        } catch (e) {
            log('  Besitzer: ' + (e.message || e));
        }
    }
    return { group: group, created: true };
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

async function runCreate() {
    const token = await getToken();
    const demos = selectedDemo();
    const addMembers = !!($('spAddDemoMembers') && $('spAddDemoMembers').checked);
    if (addMembers) {
        const v = validateDemoPool(demos);
        if (!v.ok) throw new Error(v.issues.join(', '));
    }

    if (mode === 'single') {
        const plan = buildSpielwiesenPlan({
            label: ($('spLabel') && $('spLabel').value) || '',
            year: getYear(),
            asDemo: asDemo()
        });
        if (!plan.ok) throw new Error(plan.issues.join(', '));
        if (!window.confirm('Einzel-Team anlegen?\n\n' + plan.displayName)) return;
        const res = await createSpielTeam(token, plan, null);
        if (res.group && res.group.id) await addDemoMembers(token, res.group.id);
        toast(res.created ? 'Team angelegt.' : 'Team existierte bereits.');
        return;
    }

    const selected = teachers.filter(function (t) {
        return t.selected && t.email;
    });
    const bulk = buildBulkTeacherPlans({
        teachers: selected,
        year: getYear(),
        asDemo: asDemo(),
        selectedCodes: selected.map(function (t) {
            return t.code;
        })
    });
    if (!bulk.ok) throw new Error(bulk.issues.join('; '));
    if (
        !window.confirm(
            bulk.plans.length +
                ' Lehrer-Spielwiesen anlegen?\nDemo-Schüler: ' +
                demos.length +
                '\n\nBestehende Alias werden übersprungen.'
        )
    ) {
        return;
    }

    let ok = 0;
    let skip = 0;
    let fail = 0;
    for (let i = 0; i < bulk.plans.length; i++) {
        const plan = bulk.plans[i];
        try {
            let ownerId = null;
            try {
                const owner = await G().resolveUserByEmail(token, plan.email);
                ownerId = owner && owner.id ? owner.id : null;
            } catch (e) {
                log(plan.code + ': Besitzer nicht gefunden (' + plan.email + ') – ' + (e.message || e));
            }
            if (!ownerId) {
                fail++;
                continue;
            }
            const res = await createSpielTeam(token, plan, ownerId);
            if (!res.created) skip++;
            else ok++;
            if (res.group && res.group.id) await addDemoMembers(token, res.group.id);
        } catch (e) {
            fail++;
            log(plan.code + ' FEHLER: ' + (e.message || e));
        }
        await G().sleep(400);
    }
    log('Fertig: ' + ok + ' neu, ' + skip + ' übersprungen, ' + fail + ' Fehler.');
    toast('Fertig: ' + ok + ' neu, ' + skip + ' übersprungen, ' + fail + ' Fehler');
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
    setMode('teacher');
    setStep(1);
    refreshSinglePreview();

    document.querySelectorAll('[data-sp-mode]').forEach(function (el) {
        el.addEventListener('click', function () {
            setMode(el.getAttribute('data-sp-mode'));
        });
        el.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter' || ev.key === ' ') {
                ev.preventDefault();
                setMode(el.getAttribute('data-sp-mode'));
            }
        });
    });
    document.querySelectorAll('[data-sp-goto]').forEach(function (el) {
        el.addEventListener('click', function () {
            setStep(el.getAttribute('data-sp-goto'));
        });
    });
    document.querySelectorAll('[data-sp-next]').forEach(function (el) {
        el.addEventListener('click', function () {
            setStep(el.getAttribute('data-sp-next'));
        });
    });

    ['spLabel', 'spYear', 'spAsDemo'].forEach(function (id) {
        const el = $(id);
        if (!el) return;
        el.addEventListener('input', function () {
            refreshSinglePreview();
            renderTeachers();
            refreshRunSummary();
        });
        el.addEventListener('change', function () {
            refreshSinglePreview();
            renderTeachers();
            refreshRunSummary();
        });
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
