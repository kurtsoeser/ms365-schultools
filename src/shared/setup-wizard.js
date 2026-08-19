import {
    SW_ADMIN_DEFAULT_ROLES,
    normStr,
    normEmail,
    escapeHtml,
    normCode,
    resolveAdminSlotFromRow,
    getAdminDisplayRowsFromSettings,
    collectDirektionOwnerEmails,
    collectAdminOwnerEmails,
    collectEmails,
    deriveAdministrationEntries,
    randomTempPassword
} from './setup-wizard-admin-model.js';
import {
    SW_MATCH_MEMBER_PREVIEW,
    buildLinkedGroupSummaryHtml
} from './setup-wizard-group-summary.js';

(function () {
    'use strict';

    const G = function () {
        const x = window.ms365GraphUnifiedGroups;
        if (!x) throw new Error('graph-unified-groups.js fehlt.');
        return x;
    };

    let swActiveKind = 'lehrer';
    let swMatched = {
        schuelerGroupId: null,
        lehrerGroupId: null,
        verwaltungGroupId: null,
        sgaGroupId: null,
        studentCouncilGroupId: null
    };
    let swListCache = { students: [], teachers: [], direktion: [], adminEmails: [] };
    /** @type {{ lehrer: { dn: string, nick: string, desc: string, team: boolean }, schueler: { dn: string, nick: string, desc: string, team: boolean } }} */
    let swGroupFormCache = {
        lehrer: { dn: '', nick: '', desc: '', team: false },
        schueler: { dn: '', nick: '', desc: '', team: false }
    };
    let swPrevWizardStep = 1;
    let swGroupFormCacheBootstrapped = false;

    let swVerwaltungFormCache = { dn: '', nick: '', desc: '', team: false };

    function groupUiSuffix(kind) {
        if (kind === 'lehrer') return 'Lehrer';
        if (kind === 'verwaltung') return 'Verwaltung';
        if (kind === 'sga') return 'Sga';
        if (kind === 'studentCouncil') return 'StudentCouncil';
        return 'Schueler';
    }

    /** Schulbezug für Gruppen-Beschreibungen (Domain aus Stammdaten, falls gesetzt). */
    function schoolPhraseForWizardGroupDesc() {
        try {
            const s = loadTenantSettings();
            const d = normStr(s && s.domain || '').replace(/^@+/, '');
            if (d) return ' der Schule (' + d + ')';
        } catch {
            // ignore
        }
        return ' der Schule';
    }

    function defaultLehrerForm() {
        return {
            dn: 'Lehrer:innen',
            nick: 'lehrer',
            desc: 'Alle Lehrer:innen' + schoolPhraseForWizardGroupDesc(),
            team: false
        };
    }

    function defaultSchuelerForm() {
        return {
            dn: 'Schüler:innen',
            nick: 'schueler',
            desc: 'Alle Schüler:innen' + schoolPhraseForWizardGroupDesc(),
            team: false
        };
    }

    function defaultVerwaltungForm() {
        return {
            dn: 'Schulverwaltung',
            nick: 'verwaltung',
            desc: 'Kontakte der Schulverwaltung' + schoolPhraseForWizardGroupDesc(),
            team: false
        };
    }

    function readGroupFormFromDom(kind) {
        const s = groupUiSuffix(kind);
        const dn = document.getElementById('swNewDn' + s);
        const nn = document.getElementById('swNewNick' + s);
        const dd = document.getElementById('swNewDesc' + s);
        const ct = document.getElementById('swCreateTeam' + s);
        return {
            dn: dn ? String(dn.value || '') : '',
            nick: nn ? String(nn.value || '') : '',
            desc: dd ? String(dd.value || '') : '',
            team: ct ? !!ct.checked : false
        };
    }

    function writeGroupFormToDom(kind) {
        const c = swGroupFormCache[kind];
        const s = groupUiSuffix(kind);
        const dn = document.getElementById('swNewDn' + s);
        const nn = document.getElementById('swNewNick' + s);
        const dd = document.getElementById('swNewDesc' + s);
        const ct = document.getElementById('swCreateTeam' + s);
        if (dn) dn.value = c.dn;
        if (nn) nn.value = c.nick;
        if (dd) dd.value = c.desc;
        if (ct) ct.checked = c.team;
    }

    function ensureDefaultsInCache(kind) {
        const d = kind === 'lehrer' ? defaultLehrerForm() : defaultSchuelerForm();
        const other = kind === 'lehrer' ? defaultSchuelerForm() : defaultLehrerForm();
        const c = swGroupFormCache[kind];
        const nickOther = normStr(other.nick).toLowerCase();
        const dnOther = normStr(other.dn);
        const nickCur = normStr(c.nick).toLowerCase();
        const dnCur = normStr(c.dn);
        // slgDraft speichert nur ein slgNew*-Tupel für beide Gruppen; bei activeKind/Altlasten
        // können Lehrer-Vorgaben fälschlich im Schüler-Cache (oder umgekehrt) landen.
        if (kind === 'schueler' && (nickCur === nickOther || dnCur === dnOther)) {
            c.dn = d.dn;
            c.nick = d.nick;
            c.desc = d.desc;
            return;
        }
        if (kind === 'lehrer' && (nickCur === nickOther || dnCur === dnOther)) {
            c.dn = d.dn;
            c.nick = d.nick;
            c.desc = d.desc;
            return;
        }
        if (!normStr(c.dn)) c.dn = d.dn;
        if (!normStr(c.nick)) c.nick = d.nick;
        if (!normStr(c.desc)) c.desc = d.desc;
    }

    function captureAllGroupForms() {
        swGroupFormCache.lehrer = readGroupFormFromDom('lehrer');
        swGroupFormCache.schueler = readGroupFormFromDom('schueler');
    }

    function initSwGroupFormCacheFromSlgDraft() {
        swGroupFormCache.lehrer = defaultLehrerForm();
        swGroupFormCache.schueler = defaultSchuelerForm();
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const d = window.ms365AppDataV2.getSetup().slgDraft || {};
                const ak = d.activeKind === 'lehrer' ? 'lehrer' : 'schueler';
                if (d.slgNewDisplayName != null && String(d.slgNewDisplayName).trim() !== '') {
                    swGroupFormCache[ak].dn = String(d.slgNewDisplayName);
                }
                if (d.slgNewMailNick != null && String(d.slgNewMailNick).trim() !== '') {
                    swGroupFormCache[ak].nick = String(d.slgNewMailNick);
                }
                if (d.slgNewDescription != null && String(d.slgNewDescription).trim() !== '') {
                    swGroupFormCache[ak].desc = String(d.slgNewDescription);
                }
                swGroupFormCache[ak].team = !!d.slgNewCreateTeam;
            }
        } catch {
            // ignore
        }
        ensureDefaultsInCache('lehrer');
        ensureDefaultsInCache('schueler');
        applySlgOwnerDraftFromSetupToDom();
    }

    function initSwVerwaltungFormCacheFromDraft() {
        swVerwaltungFormCache = defaultVerwaltungForm();
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const vd = window.ms365AppDataV2.getSetup().verwaltungDraft || {};
                if (vd.vwNewDisplayName != null && String(vd.vwNewDisplayName).trim() !== '') {
                    swVerwaltungFormCache.dn = String(vd.vwNewDisplayName);
                }
                if (vd.vwNewMailNick != null && String(vd.vwNewMailNick).trim() !== '') {
                    swVerwaltungFormCache.nick = String(vd.vwNewMailNick);
                }
                if (vd.vwNewDescription != null && String(vd.vwNewDescription).trim() !== '') {
                    swVerwaltungFormCache.desc = String(vd.vwNewDescription);
                }
                swVerwaltungFormCache.team = !!vd.vwNewCreateTeam;
            }
        } catch {
            // ignore
        }
    }

    function readVerwaltungFormFromDom() {
        const dn = document.getElementById('swNewDnVerwaltung');
        const nn = document.getElementById('swNewNickVerwaltung');
        const dd = document.getElementById('swNewDescVerwaltung');
        const ct = document.getElementById('swCreateTeamVerwaltung');
        return {
            dn: dn ? String(dn.value || '') : '',
            nick: nn ? String(nn.value || '') : '',
            desc: dd ? String(dd.value || '') : '',
            team: ct ? !!ct.checked : false
        };
    }

    function writeVerwaltungFormToDom() {
        const c = swVerwaltungFormCache;
        const dn = document.getElementById('swNewDnVerwaltung');
        const nn = document.getElementById('swNewNickVerwaltung');
        const dd = document.getElementById('swNewDescVerwaltung');
        const ct = document.getElementById('swCreateTeamVerwaltung');
        if (dn) dn.value = c.dn;
        if (nn) nn.value = c.nick;
        if (dd) dd.value = c.desc;
        if (ct) ct.checked = c.team;
    }

    function captureVerwaltungFormToCache() {
        swVerwaltungFormCache = readVerwaltungFormFromDom();
    }

    function buildVerwaltungDraftFromCache() {
        const c = swVerwaltungFormCache;
        const ow = readVerwaltungOwnerDraftFromDom();
        return {
            vwNewDisplayName: c.dn,
            vwNewMailNick: c.nick,
            vwNewDescription: c.desc,
            vwNewCreateTeam: c.team,
            vwOwnerSource: ow.vwOwnerSource,
            vwOwnerManualEmails: ow.vwOwnerManualEmails
        };
    }

    function persistVerwaltungFull() {
        captureVerwaltungFormToCache();
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { verwaltungGroupId: swMatched.verwaltungGroupId },
                    verwaltungDraft: buildVerwaltungDraftFromCache()
                });
            }
        } catch {
            // ignore
        }
    }

    function persistVerwaltungDraftPatch() {
        captureVerwaltungFormToCache();
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { verwaltungGroupId: swMatched.verwaltungGroupId },
                    verwaltungDraft: buildVerwaltungDraftFromCache()
                });
            }
        } catch {
            // ignore
        }
    }

    function getAdminDisplayRowsForWizard() {
        return getAdminDisplayRowsFromSettings(loadTenantSettings());
    }

    function seedDefaultAdminRolesToSettings() {
        const s = loadTenantSettings() || {};
        const admin = SW_ADMIN_DEFAULT_ROLES.map(function (slot) {
            return { defaultKey: slot, role: slot, name: '', email: '' };
        });
        const roles =
            typeof window.ms365TenantSettingsDefaultAdminRoles === 'function'
                ? window.ms365TenantSettingsDefaultAdminRoles()
                : [];
        s.administration = deriveAdministrationEntries(roles, admin);
        s.admin = admin;
        s.adminRoles = roles;
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(s);
        }
        readLists();
        renderSwAdminTableBody();
        toast('Standardrollen eingefügt und gespeichert.');
    }

    function appendSwAdminM365AndActionCells(tr, emailForPreview) {
        const tdMs = createSwDirectoryMatchTd(emailForPreview);
        const tdAct = document.createElement('td');
        tdAct.className = 'action-cell';
        tdAct.style.whiteSpace = 'nowrap';
        const wrap = document.createElement('div');
        wrap.style.cssText =
            'display:inline-flex;flex-wrap:nowrap;gap:6px;align-items:center;justify-content:flex-end;';
        const btnMs = document.createElement('button');
        btnMs.type = 'button';
        btnMs.className = 'mini-btn';
        btnMs.style.background = '#5e72e4';
        btnMs.title = 'Diese E‑Mail in Microsoft Entra prüfen';
        btnMs.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
        btnMs.addEventListener('click', function () {
            const inpE = tr.querySelector('input[data-sw-admin-email]');
            verifyGraphDirectoryOneEmail(inpE ? inpE.value : '', 'admin');
        });
        const btnCr = document.createElement('button');
        btnCr.type = 'button';
        btnCr.className = 'mini-btn';
        btnCr.style.background = '#11cdef';
        btnCr.title = 'Neuen Benutzer in Microsoft Entra ID anlegen (User.ReadWrite.All)';
        btnCr.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
        btnCr.addEventListener('click', function () {
            const inpE = tr.querySelector('input[data-sw-admin-email]');
            const inpN = tr.querySelector('input[data-sw-admin-name]');
            runCreateEntraUserForVerwaltungRow(inpE ? inpE.value : '', inpN ? inpN.value : '');
        });
        const btnDel = document.createElement('button');
        btnDel.type = 'button';
        btnDel.className = 'mini-btn';
        btnDel.textContent = '✕';
        btnDel.title = 'Zeile entfernen';
        btnDel.addEventListener('click', function () {
            tr.remove();
            refreshSwStatsVerwaltung();
            refreshSwOwnerSummary('Verwaltung', 'verwaltung');
        });
        wrap.appendChild(btnMs);
        wrap.appendChild(btnCr);
        wrap.appendChild(btnDel);
        tdAct.appendChild(wrap);
        tr.appendChild(tdMs);
        tr.appendChild(tdAct);
    }

    function appendSwAdminCustomRowToWizardTable(tbody, row) {
        const tr = document.createElement('tr');
        tr.setAttribute('data-sw-admin-custom', '1');
        const tdR = document.createElement('td');
        const inpR = document.createElement('input');
        inpR.type = 'text';
        inpR.setAttribute('data-sw-admin-role', '1');
        inpR.autocomplete = 'off';
        inpR.placeholder = 'Rollenbezeichnung';
        inpR.value = row && row.role ? String(row.role) : '';
        inpR.style.width = '100%';
        tdR.appendChild(inpR);
        const tdN = document.createElement('td');
        const inpN = document.createElement('input');
        inpN.type = 'text';
        inpN.setAttribute('data-sw-admin-name', '1');
        inpN.autocomplete = 'off';
        inpN.value = row && row.name ? String(row.name) : '';
        inpN.style.width = '100%';
        tdN.appendChild(inpN);
        const tdE = document.createElement('td');
        const inpE = document.createElement('input');
        inpE.type = 'email';
        inpE.setAttribute('data-sw-admin-email', '1');
        inpE.autocomplete = 'off';
        inpE.spellcheck = false;
        inpE.inputMode = 'email';
        inpE.placeholder = 'name@schule.at';
        inpE.value = row && row.email ? String(row.email) : '';
        inpE.style.width = '100%';
        tdE.appendChild(inpE);
        tr.appendChild(tdR);
        tr.appendChild(tdN);
        tr.appendChild(tdE);
        appendSwAdminM365AndActionCells(tr, row && row.email ? String(row.email) : '');
        tbody.appendChild(tr);
    }

    function renderSwAdminTableBody() {
        const tbody = document.getElementById('swAdminTableBody');
        if (!tbody) return;
        const rows = getAdminDisplayRowsForWizard();
        tbody.replaceChildren();

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Keine Einträge – „Standardrollen einfügen“ oder „Weitere Rolle“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
        } else {
            rows.forEach(function (row) {
                const slot = resolveAdminSlotFromRow(row);
                const tr = document.createElement('tr');
                const em0 = normStr(row.email || '');

                if (slot) {
                    tr.setAttribute('data-sw-admin-slot', slot);
                    const label = row && normStr(row.role) ? normStr(row.role) : slot;
                    const tdR = document.createElement('td');
                    tdR.setAttribute('data-sw-admin-role-display', '1');
                    tdR.textContent = label;
                    tdR.title = 'Doppelklick zum Bearbeiten';
                    tdR.addEventListener('dblclick', function () {
                        const current = normStr(tdR.textContent) || slot;
                        wizardStartCellEdit(tdR, current, function (next, meta) {
                            const text = meta && meta.cancelled ? current : normStr(next);
                            tdR.textContent = text || slot;
                            refreshSwStatsVerwaltung();
                            refreshSwOwnerSummary('Verwaltung', 'verwaltung');
                        });
                    });
                    const tdN = document.createElement('td');
                    const inpN = document.createElement('input');
                    inpN.type = 'text';
                    inpN.setAttribute('data-sw-admin-name', '1');
                    inpN.autocomplete = 'off';
                    inpN.value = row && row.name ? String(row.name) : '';
                    inpN.style.width = '100%';
                    tdN.appendChild(inpN);
                    const tdE = document.createElement('td');
                    const inpE = document.createElement('input');
                    inpE.type = 'email';
                    inpE.setAttribute('data-sw-admin-email', '1');
                    inpE.autocomplete = 'off';
                    inpE.spellcheck = false;
                    inpE.inputMode = 'email';
                    inpE.placeholder = 'name@schule.at';
                    inpE.value = em0;
                    inpE.style.width = '100%';
                    tdE.appendChild(inpE);
                    tr.appendChild(tdR);
                    tr.appendChild(tdN);
                    tr.appendChild(tdE);
                    appendSwAdminM365AndActionCells(tr, em0);
                } else {
                    tr.setAttribute('data-sw-admin-custom', '1');
                    const tdR = document.createElement('td');
                    const inpR = document.createElement('input');
                    inpR.type = 'text';
                    inpR.setAttribute('data-sw-admin-role', '1');
                    inpR.autocomplete = 'off';
                    inpR.placeholder = 'Rollenbezeichnung';
                    inpR.value = row && row.role ? String(row.role) : '';
                    inpR.style.width = '100%';
                    tdR.appendChild(inpR);
                    const tdN = document.createElement('td');
                    const inpN = document.createElement('input');
                    inpN.type = 'text';
                    inpN.setAttribute('data-sw-admin-name', '1');
                    inpN.autocomplete = 'off';
                    inpN.value = row && row.name ? String(row.name) : '';
                    inpN.style.width = '100%';
                    tdN.appendChild(inpN);
                    const tdE = document.createElement('td');
                    const inpE = document.createElement('input');
                    inpE.type = 'email';
                    inpE.setAttribute('data-sw-admin-email', '1');
                    inpE.autocomplete = 'off';
                    inpE.spellcheck = false;
                    inpE.inputMode = 'email';
                    inpE.placeholder = 'name@schule.at';
                    inpE.value = em0;
                    inpE.style.width = '100%';
                    tdE.appendChild(inpE);
                    tr.appendChild(tdR);
                    tr.appendChild(tdN);
                    tr.appendChild(tdE);
                    appendSwAdminM365AndActionCells(tr, em0);
                }
                tbody.appendChild(tr);
            });
        }

        if (!tbody.dataset.swAdminInputBound) {
            tbody.dataset.swAdminInputBound = '1';
            tbody.addEventListener('input', function () {
                refreshSwStatsVerwaltung();
                refreshSwOwnerSummary('Verwaltung', 'verwaltung');
            });
        }
        refreshSwStatsVerwaltung();
        refreshSwOwnerSummary('Verwaltung', 'verwaltung');
    }

    function gatherSwAdminRowsFromTable() {
        const tbody = document.getElementById('swAdminTableBody');
        if (!tbody) return [];
        const out = [];
        tbody.querySelectorAll('tr[data-sw-admin-slot]').forEach(function (tr) {
            const slotKey = normStr(tr.getAttribute('data-sw-admin-slot'));
            const tdR = tr.querySelector('[data-sw-admin-role-display="1"]');
            const displayRole = tdR ? normStr(tdR.textContent) : slotKey;
            const inpN = tr.querySelector('input[data-sw-admin-name]');
            const inpE = tr.querySelector('input[data-sw-admin-email]');
            const name = inpN ? normStr(inpN.value) : '';
            const email = inpE ? normStr(inpE.value).toLowerCase() : '';
            const roleOut = displayRole || slotKey;
            if (!roleOut && !name && !email) return;
            out.push({
                defaultKey: slotKey,
                role: roleOut,
                name: name || '',
                email: email || ''
            });
        });
        tbody.querySelectorAll('tr[data-sw-admin-custom="1"]').forEach(function (tr) {
            const inpR = tr.querySelector('input[data-sw-admin-role]');
            const inpN = tr.querySelector('input[data-sw-admin-name]');
            const inpE = tr.querySelector('input[data-sw-admin-email]');
            const role = inpR ? normStr(inpR.value) : '';
            const name = inpN ? normStr(inpN.value) : '';
            const email = inpE ? normStr(inpE.value).toLowerCase() : '';
            if (!role && !name && !email) return;
            out.push({ role: role || '', name: name || '', email: email || '' });
        });
        return out;
    }

    function persistSwAdminListFromTableSilent() {
        const tbody = document.getElementById('swAdminTableBody');
        if (!tbody) return;
        const s = loadTenantSettings() || {};
        const rows = gatherSwAdminRowsFromTable().filter(function (r) {
            return r && (normStr(r.role) || normStr(r.name) || normStr(r.email));
        });
        s.admin = rows;
        s.administration = deriveAdministrationEntries(Array.isArray(s.adminRoles) ? s.adminRoles : [], rows);
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(s);
        }
        readLists();
    }

    function saveSwAdminList() {
        persistSwAdminListFromTableSilent();
        renderSwAdminTableBody();
        refreshSwStatsVerwaltung();
        refreshSwOwnerSummary('Verwaltung', 'verwaltung');
        toast('Verwaltung gespeichert.');
    }

    function toast(msg) {
        if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
        } else if (typeof window.ms365AppDialogAlert === 'function') {
            void window.ms365AppDialogAlert(msg, { title: 'Hinweis' });
        } else {
            window.alert(msg);
        }
    }

    function dlgAlert(msg) {
        if (typeof window.ms365AppDialogAlert === 'function') {
            return window.ms365AppDialogAlert(msg);
        }
        window.alert(msg);
        return Promise.resolve();
    }

    function dlgConfirm(msg, opts) {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(msg, opts);
        }
        return Promise.resolve(window.confirm(msg));
    }

    function dlgPrompt(msg, def, opts) {
        if (typeof window.ms365AppDialogPrompt === 'function') {
            return window.ms365AppDialogPrompt(msg, def, opts);
        }
        return Promise.resolve(window.prompt(msg, def));
    }

    function loadTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return null;
        return window.ms365TenantSettingsLoad();
    }

    function readLists() {
        const settings = loadTenantSettings();
        swListCache.students = collectEmails(settings && settings.students);
        swListCache.teachers = collectEmails(settings && settings.teachers);
        swListCache.direktion = collectDirektionOwnerEmails(settings);
        swListCache.adminEmails = collectAdminOwnerEmails(settings);
    }

    function syncSetupFromAppData() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                if (su && su.matched) {
                    swMatched.schuelerGroupId = su.matched.schuelerGroupId || null;
                    swMatched.lehrerGroupId = su.matched.lehrerGroupId || null;
                    swMatched.verwaltungGroupId = su.matched.verwaltungGroupId || null;
                    swMatched.sgaGroupId = su.matched.sgaGroupId || null;
                    swMatched.studentCouncilGroupId = su.matched.studentCouncilGroupId || null;
                }
            }
        } catch {
            // ignore
        }
    }

    function persistMatched() {
        captureAllGroupForms();
        const cur = swGroupFormCache[swActiveKind];
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                const own = readSlgOwnerDraftFromDom();
                window.ms365AppDataV2.patchSetup({
                    matched: swMatched,
                    slgDraft: Object.assign(
                        {
                            activeKind: swActiveKind,
                            slgNewDisplayName: cur.dn,
                            slgNewMailNick: cur.nick,
                            slgNewDescription: cur.desc,
                            slgNewCreateTeam: cur.team
                        },
                        own
                    )
                });
            }
        } catch {
            // ignore
        }
        try {
            localStorage.setItem(
                'ms365-schueler-lehrer-gruppen-v2',
                JSON.stringify({
                    kind: 'ms365-schueler-lehrer-gruppen-v2',
                    savedAt: new Date().toISOString(),
                    activeKind: swActiveKind,
                    matched: swMatched,
                    slgNewDisplayName: cur.dn,
                    slgNewMailNick: cur.nick,
                    slgNewDescription: cur.desc,
                    slgNewCreateTeam: cur.team
                })
            );
        } catch {
            // ignore
        }
    }

    function getActiveGid() {
        return swActiveKind === 'lehrer' ? swMatched.lehrerGroupId : swMatched.schuelerGroupId;
    }

    function setActiveGid(id) {
        if (swActiveKind === 'lehrer') swMatched.lehrerGroupId = id;
        else swMatched.schuelerGroupId = id;
    }

    async function ensureOwnersDirektionOnly(token, groupId) {
        return G().ensureOwners(token, groupId, swListCache.direktion || []);
    }

    async function ensureOwnersForSlgKind(token, groupId, forKind) {
        return G().ensureOwners(token, groupId, resolveOwnerEmailsForWizard(forKind));
    }

    async function ensureOwnersForVerwaltungWizard(token, groupId) {
        return G().ensureOwners(token, groupId, resolveOwnerEmailsForWizard('verwaltung'));
    }

    function appendSwLogToSuffix(suffix, msg, logKind) {
        const el = document.getElementById('swSyncLog' + suffix);
        if (!el) return;
        const line = document.createElement('div');
        line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
        if (logKind === 'err') line.style.color = '#b00020';
        else if (logKind === 'ok') line.style.color = '#0d8050';
        else if (logKind === 'warn') line.style.color = '#856404';
        el.appendChild(line);
        el.scrollTop = el.scrollHeight;
    }

    function appendSwLog(msg, logKind) {
        appendSwLogToSuffix(groupUiSuffix(swActiveKind), msg, logKind);
    }

    function getGroupIdForKind(kind) {
        if (kind === 'lehrer') return swMatched.lehrerGroupId;
        if (kind === 'verwaltung') return swMatched.verwaltungGroupId;
        if (kind === 'sga') return swMatched.sgaGroupId;
        if (kind === 'studentCouncil') return swMatched.studentCouncilGroupId;
        return swMatched.schuelerGroupId;
    }

    function setGroupIdForKind(kind, id) {
        if (kind === 'lehrer') swMatched.lehrerGroupId = id;
        else if (kind === 'verwaltung') swMatched.verwaltungGroupId = id;
        else if (kind === 'sga') swMatched.sgaGroupId = id;
        else if (kind === 'studentCouncil') swMatched.studentCouncilGroupId = id;
        else swMatched.schuelerGroupId = id;
    }

    /** @type {Record<string, { gen: number, gid: string, snapshot: object|null, fullyLoaded: boolean, inFlight: boolean }>} */
    const swMatchLiveByKind = {
        lehrer: { gen: 0, gid: '', snapshot: null, fullyLoaded: false, inFlight: false },
        schueler: { gen: 0, gid: '', snapshot: null, fullyLoaded: false, inFlight: false },
        verwaltung: { gen: 0, gid: '', snapshot: null, fullyLoaded: false, inFlight: false },
        sga: { gen: 0, gid: '', snapshot: null, fullyLoaded: false, inFlight: false },
        studentCouncil: { gen: 0, gid: '', snapshot: null, fullyLoaded: false, inFlight: false }
    };

    function swMatchLiveSlot(kind) {
        return swMatchLiveByKind[kind] || swMatchLiveByKind.schueler;
    }

    function swAuthIsLoggedIn() {
        try {
            return typeof window.ms365AuthIsLoggedIn === 'function' && !!window.ms365AuthIsLoggedIn();
        } catch {
            return false;
        }
    }

    function refreshSwAuthStatus() {
        const el = document.getElementById('swAuthStatus');
        if (!el) return;
        if (swAuthIsLoggedIn()) {
            const label =
                typeof window.ms365AuthGetAccountLabel === 'function' ? window.ms365AuthGetAccountLabel() : '';
            el.textContent = label
                ? 'Angemeldet als ' + label + '.'
                : 'Angemeldet (Token für Microsoft Graph erhalten).';
        } else {
            el.textContent = 'Noch nicht angemeldet.';
        }
    }

    function emptySwMatchSnapshot(gid, group) {
        return {
            groupId: gid,
            group: group || null,
            owners: [],
            members: [],
            memberCount: -1,
            membersTruncated: false,
            artLabel: '',
            hasTeam: false,
            status: group ? 'partial' : 'loading',
            error: ''
        };
    }

    function mergeSwMatchGroup(snapshot, group) {
        if (!group) return snapshot;
        snapshot.group = Object.assign({}, snapshot.group || {}, group);
        if (!snapshot.groupId && group.id) snapshot.groupId = String(group.id);
        return snapshot;
    }

    function paintSwMatchSummary(forKind) {
        const el = document.getElementById('swMatchSummary' + groupUiSuffix(forKind));
        if (!el) return;
        const gid = getGroupIdForKind(forKind);
        if (!gid) {
            el.innerHTML = '<span class="sw-match-empty">Noch keine Gruppe gewählt.</span>';
            return;
        }
        const slot = swMatchLiveSlot(forKind);
        const snap =
            slot.gid === gid && slot.snapshot
                ? slot.snapshot
                : emptySwMatchSnapshot(gid, slot.snapshot && slot.snapshot.group);
        el.innerHTML = buildLinkedGroupSummaryHtml(snap);
    }

    function renderSwMatchSummaryForKind(forKind, g) {
        const gid = getGroupIdForKind(forKind);
        const slot = swMatchLiveSlot(forKind);
        if (!gid) {
            slot.gid = '';
            slot.snapshot = null;
            slot.fullyLoaded = false;
            paintSwMatchSummary(forKind);
            return;
        }
        if (slot.gid !== gid) {
            slot.gid = gid;
            slot.snapshot = emptySwMatchSnapshot(gid, g || null);
            slot.fullyLoaded = false;
            slot.inFlight = false;
            slot.gen += 1;
        } else {
            if (!slot.snapshot) slot.snapshot = emptySwMatchSnapshot(gid, g || null);
            else mergeSwMatchGroup(slot.snapshot, g);
        }
        paintSwMatchSummary(forKind);
        if (!slot.fullyLoaded) loadSwMatchLiveDetails(forKind);
    }

    async function fetchSwMatchMemberPreview(token, gid) {
        const select = G().PERSON_SELECT || 'id,displayName,mail,userPrincipalName';
        const path =
            '/groups/' +
            encodeURIComponent(gid) +
            '/members?$select=' +
            encodeURIComponent(select) +
            '&$top=' +
            String(SW_MATCH_MEMBER_PREVIEW);
        const data = await G().graphJson('GET', path, token, undefined);
        const items = (data && data.value) || [];
        return { items: items, truncated: !!(data && data['@odata.nextLink']) };
    }

    async function loadSwMatchLiveDetails(forKind, opts) {
        const gid = getGroupIdForKind(forKind);
        const el = document.getElementById('swMatchSummary' + groupUiSuffix(forKind));
        if (!gid || !el) return;
        const slot = swMatchLiveSlot(forKind);
        const force = !!(opts && opts.force);
        if (slot.inFlight && slot.gid === gid && !force) return;
        if (slot.fullyLoaded && slot.gid === gid && !force) return;

        if (!force && !swAuthIsLoggedIn()) {
            if (!slot.snapshot || slot.gid !== gid) slot.snapshot = emptySwMatchSnapshot(gid, null);
            slot.snapshot.status = 'needsLogin';
            slot.gid = gid;
            paintSwMatchSummary(forKind);
            return;
        }

        const gen = ++slot.gen;
        slot.gid = gid;
        slot.inFlight = true;
        if (!slot.snapshot || slot.snapshot.groupId !== gid) slot.snapshot = emptySwMatchSnapshot(gid, slot.snapshot && slot.snapshot.group);
        slot.snapshot.status = 'loading';
        slot.snapshot.error = '';
        paintSwMatchSummary(forKind);

        try {
            const token = await G().getGraphToken();
            if (gen !== slot.gen || getGroupIdForKind(forKind) !== gid) return;
            const [group, owners, memberCount, memberPage] = await Promise.all([
                G().fetchGroup(token, gid),
                G().fetchGroupOwners(token, gid),
                G().fetchGroupMemberCount(token, gid).catch(function () {
                    return -1;
                }),
                fetchSwMatchMemberPreview(token, gid).catch(function () {
                    return { items: [], truncated: false };
                })
            ]);
            if (gen !== slot.gen || getGroupIdForKind(forKind) !== gid) return;
            const members = memberPage && memberPage.items ? memberPage.items : [];
            const count = typeof memberCount === 'number' ? memberCount : members.length;
            slot.snapshot = {
                groupId: gid,
                group: group || (slot.snapshot && slot.snapshot.group) || null,
                owners: owners || [],
                members: members,
                memberCount: count,
                membersTruncated: !!(memberPage && memberPage.truncated) || (count > members.length && members.length > 0),
                artLabel: group ? G().groupArtLabel(group) : '',
                hasTeam: !!(group && G().groupHasTeam(group)),
                status: 'ready',
                error: ''
            };
            slot.fullyLoaded = true;
            paintSwMatchSummary(forKind);
        } catch (e) {
            if (gen !== slot.gen || getGroupIdForKind(forKind) !== gid) return;
            if (!slot.snapshot) slot.snapshot = emptySwMatchSnapshot(gid, null);
            slot.snapshot.status = 'error';
            slot.snapshot.error = 'Details konnten nicht geladen werden: ' + (e && e.message ? e.message : String(e));
            slot.fullyLoaded = false;
            paintSwMatchSummary(forKind);
        } finally {
            if (gen === slot.gen) slot.inFlight = false;
        }
    }

    function wireSwMatchSummaryRefresh(forKind) {
        const el = document.getElementById('swMatchSummary' + groupUiSuffix(forKind));
        if (!el || el.dataset.swMatchRefreshWired === '1') return;
        el.dataset.swMatchRefreshWired = '1';
        el.addEventListener('click', function (ev) {
            const t = ev.target;
            const node = t && t.nodeType === 1 ? t : t && t.parentElement;
            const btn = node && node.closest ? node.closest('[data-sw-match-refresh]') : null;
            if (!btn) return;
            const slot = swMatchLiveSlot(forKind);
            slot.fullyLoaded = false;
            loadSwMatchLiveDetails(forKind, { force: true });
        });
    }

    function renderSwSearchResults(list, forKind) {
        const host = document.getElementById('swSearchResults' + groupUiSuffix(forKind));
        if (!host) return;
        host.replaceChildren();
        if (!list || !list.length) {
            host.innerHTML = '<div style="color:var(--muted)">Keine Treffer.</div>';
            return;
        }
        list.forEach(function (g) {
            const row = document.createElement('div');
            row.style.cssText =
                'display:flex;flex-wrap:wrap;justify-content:space-between;align-items:center;gap:10px;padding:10px 12px;margin-bottom:8px;border:1px solid var(--border);border-radius:12px;background:var(--card,#fff);';
            const left = document.createElement('div');
            const mailLine = g.mail ? ' · ' + escapeHtml(g.mail) : '';
            const descLine = g.description
                ? '<div style="font-size:0.86em;color:var(--muted);margin-top:4px;">' +
                  escapeHtml(g.description) +
                  '</div>'
                : '';
            left.innerHTML =
                '<div style="font-weight:800">' +
                escapeHtml(g.displayName || '') +
                '</div><div style="font-size:0.9em;color:var(--muted)"><code>' +
                escapeHtml(g.mailNickname || '') +
                '</code>' +
                mailLine +
                '</div>' +
                descLine;
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-success btn-sm';
            btn.textContent = 'Verknüpfen';
            btn.addEventListener('click', function () {
                setGroupIdForKind(forKind, String(g.id));
                if (forKind === 'verwaltung') persistVerwaltungFull();
                else persistMatched();
                const slot = swMatchLiveSlot(forKind);
                slot.gid = String(g.id);
                slot.snapshot = emptySwMatchSnapshot(String(g.id), g);
                slot.fullyLoaded = false;
                slot.inFlight = false;
                slot.gen += 1;
                renderSwMatchSummaryForKind(forKind, g);
                toast('Gruppe verknüpft.');
            });
            row.appendChild(left);
            row.appendChild(btn);
            host.appendChild(row);
        });
    }

    function showStep(n) {
        const step = Math.max(1, Math.min(11, parseInt(n, 10) || 1));
        if (swPrevWizardStep === 3 && step !== swPrevWizardStep) {
            captureVerwaltungFormToCache();
            persistVerwaltungDraftPatch();
        }
        if ((swPrevWizardStep === 4 || swPrevWizardStep === 5) && step !== swPrevWizardStep) {
            captureAllGroupForms();
            patchSlgOwnerDraftFromDom();
        }
        swPrevWizardStep = step;

        for (let i = 1; i <= 11; i++) {
            const panel = document.getElementById('swStep' + i);
            if (panel) panel.style.display = i === step ? 'block' : 'none';
        }
        document.querySelectorAll('[data-sw-step]').forEach(function (btn) {
            const sn = parseInt(btn.getAttribute('data-sw-step'), 10);
            const on = sn === step;
            btn.classList.toggle('active', on);
            btn.setAttribute('aria-selected', on ? 'true' : 'false');
        });
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.touchWizardVisit === 'function') {
                window.ms365AppDataV2.touchWizardVisit(step);
            }
        } catch {
            // ignore
        }
        if (step === 2) {
            const inp = document.getElementById('swDomain');
            const s = loadTenantSettings();
            if (inp && s && s.domain) inp.value = s.domain;
            refreshSwGraphDefaultDomainHint();
        }
        if (step === 3) {
            syncSetupFromAppData();
            initSwVerwaltungFormCacheFromDraft();
            writeVerwaltungFormToDom();
            readLists();
            renderSwAdminTableBody();
            applyVerwaltungOwnerDraftFromSetupToDom();
            renderSwMatchSummaryForKind('verwaltung');
            refreshSwWizardAuxiliaryForStep(3);
        }
        if (step === 4) {
            syncSetupFromAppData();
            swActiveKind = 'lehrer';
            ensureDefaultsInCache('lehrer');
            writeGroupFormToDom('lehrer');
            readLists();
            fillTeachersTextarea();
            applySlgOwnerDraftFromSetupToDom();
            renderSwMatchSummaryForKind('lehrer');
            refreshSwWizardAuxiliaryForStep(4);
        }
        if (step === 5) {
            syncSetupFromAppData();
            swActiveKind = 'schueler';
            ensureDefaultsInCache('schueler');
            writeGroupFormToDom('schueler');
            readLists();
            fillStudentsTextarea();
            applySlgOwnerDraftFromSetupToDom();
            renderSwMatchSummaryForKind('schueler');
            refreshSwWizardAuxiliaryForStep(5);
        }
        if (step === 6) {
            syncSetupFromAppData();
            fillSgaTextarea();
            renderSwMatchSummaryForKind('sga');
            refreshSwWizardAuxiliaryForStep(6);
        }
        if (step === 7) {
            syncSetupFromAppData();
            fillStudentCouncilTextarea();
            renderSwMatchSummaryForKind('studentCouncil');
            refreshSwWizardAuxiliaryForStep(7);
        }
        if (step === 8) {
            readGroupPrefixesFromSetupToDom();
            fillSubjectsBulkFromSettings();
            fillCatalogSlice('subject');
        }
        if (step === 9) {
            readGroupPrefixesFromSetupToDom();
            fillArgesBulkFromSettings();
            fillCatalogSlice('arge');
        }
        if (step === 10) {
            fillClassesBulkTextarea();
            renderClassesTable();
        }
        if (step === 11) {
            refreshSwStep9Summary();
        }
    }

    function refreshSwStep9Summary() {
        const box = document.getElementById('swStep9SummaryBody');
        if (!box) return;
        const s = loadTenantSettings() || {};
        const domain = normStr(s.domain || '');
        const teachers = Array.isArray(s.teachers) ? s.teachers.length : 0;
        const students = Array.isArray(s.students) ? s.students.length : 0;
        const admin = Array.isArray(s.admin) ? s.admin.length : 0;
        const subjects = Array.isArray(s.subjects) ? s.subjects.length : 0;
        const arges = Array.isArray(s.arges) ? s.arges.length : 0;
        const classes = Array.isArray(s.classes) ? s.classes.length : 0;
        const sga = Array.isArray(s.sga) ? s.sga.length : 0;
        const studentCouncil = Array.isArray(s.studentCouncil) ? s.studentCouncil.length : 0;
        let schoolYear = '–';
        let klassenTeams = 0;
        let klassenTeamsLinked = 0;
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                if (c && c.years && c.years.current) schoolYear = String(c.years.current);
                const teams = Array.isArray(c.core && c.core.classTeams) ? c.core.classTeams : [];
                klassenTeams = teams.length;
                teams.forEach(function (t) {
                    if (t && normStr(t.graphGroupId)) klassenTeamsLinked++;
                });
            }
        } catch {
            // ignore
        }
        let catSubjLinked = 0;
        let catArgeLinked = 0;
        let mVerw = false;
        let mLehr = false;
        let mSch = false;
        let mSga = false;
        let mSv = false;
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                const links = Array.isArray(su.catalogLinks) ? su.catalogLinks : [];
                links.forEach(function (L) {
                    if (!L || !normStr(L.graphGroupId)) return;
                    if (L.kind === 'arge') catArgeLinked++;
                    else catSubjLinked++;
                });
                const m = su.matched || {};
                mVerw = !!normStr(m.verwaltungGroupId);
                mLehr = !!normStr(m.lehrerGroupId);
                mSch = !!normStr(m.schuelerGroupId);
                mSga = !!normStr(m.sgaGroupId);
                mSv = !!normStr(m.studentCouncilGroupId);
            }
        } catch {
            // ignore
        }
        const yn = function (v) {
            return v ? 'Ja' : 'Nein';
        };
        box.innerHTML =
            '<p style="margin:0 0 10px;font-weight:700;color:var(--heading);">Kurzüberblick</p>' +
            '<ul style="margin:0;padding-left:1.2em;line-height:1.55;">' +
            '<li><strong>Schuljahr</strong> (Listen): ' +
            escapeHtml(schoolYear) +
            '</li>' +
            '<li><strong>Schul‑Domain</strong>: ' +
            escapeHtml(domain || '–') +
            '</li>' +
            '<li><strong>Verwaltung</strong>: ' +
            admin +
            ' Kontakt(e) · Sammelgruppe verknüpft: ' +
            yn(mVerw) +
            '</li>' +
            '<li><strong>Lehrer:innen</strong>: ' +
            teachers +
            ' · Sammelgruppe verknüpft: ' +
            yn(mLehr) +
            '</li>' +
            '<li><strong>Schüler:innen</strong>: ' +
            students +
            ' · Sammelgruppe verknüpft: ' +
            yn(mSch) +
            '</li>' +
            '<li><strong>SGA</strong>: ' +
            sga +
            ' Eintrag(e) · Gruppe verknüpft: ' +
            yn(mSga) +
            '</li>' +
            '<li><strong>Schülervertretung</strong>: ' +
            studentCouncil +
            ' Eintrag(e) · Gruppe verknüpft: ' +
            yn(mSv) +
            '</li>' +
            '<li><strong>Fächer</strong>: ' +
            subjects +
            ' in der Liste · mit Microsoft‑365‑Gruppe verknüpft (laut Einrichtung): ' +
            catSubjLinked +
            '</li>' +
            '<li><strong>ARGE</strong>: ' +
            arges +
            ' in der Liste · mit Microsoft‑365‑Gruppe verknüpft (laut Einrichtung): ' +
            catArgeLinked +
            '</li>' +
            '<li><strong>Klassen</strong>: ' +
            classes +
            ' · Klassen‑Teams im Speicher: ' +
            klassenTeams +
            ' (mit Graph‑ID: ' +
            klassenTeamsLinked +
            ')</li>' +
            '</ul>' +
            '<p style="margin:12px 0 0;color:var(--muted);font-size:0.92em;line-height:1.45;">Hinweis: Die Zahlen stammen aus den lokal gespeicherten <strong>Stammdaten</strong> und dem Einrichtungs‑Setup (Gruppen‑Verknüpfungen). Nach Änderungen in früheren Schritten hierher wechseln, um die Übersicht zu aktualisieren.</p>';
    }

    function fillTeachersTextarea() {
        const s = loadTenantSettings();
        const tt = document.getElementById('swTeachersLines');
        if (tt && s && Array.isArray(s.teachers)) {
            tt.value = s.teachers
                .map(function (t) {
                    return [t.code || '', t.name || '', t.email || ''].filter(Boolean).join(';');
                })
                .join('\n');
        }
        renderSwTeachersTableFromTextarea();
    }

    function fillStudentsTextarea() {
        const s = loadTenantSettings();
        const st = document.getElementById('swStudentsLines');
        if (st && s && Array.isArray(s.students)) {
            st.value = s.students
                .map(function (t) {
                    return [t.klasse || '', t.name || '', t.email || ''].filter(Boolean).join(';');
                })
                .join('\n');
        }
        renderSwStudentsTableFromTextarea();
    }

    function applySchuldatenMasterImportPayload(payload) {
        if (!payload || typeof payload !== 'object') return;
        const s = loadTenantSettings() || {};
        let touched = 0;
        try {
            if (normStr(payload.verwaltungLines) && window.ms365TenantSettingsParseAdminLines) {
                s.admin = window.ms365TenantSettingsParseAdminLines(payload.verwaltungLines);
                s.administration = deriveAdministrationEntries(Array.isArray(s.adminRoles) ? s.adminRoles : [], s.admin);
                touched++;
            }
            if (normStr(payload.lehrerLines) && window.ms365TenantSettingsParseTeachersLines) {
                s.teachers = window.ms365TenantSettingsParseTeachersLines(payload.lehrerLines);
                touched++;
            }
            if (normStr(payload.schuelerLines) && window.ms365TenantSettingsParseStudentsLines) {
                s.students = window.ms365TenantSettingsParseStudentsLines(payload.schuelerLines);
                touched++;
            }
            if (normStr(payload.sgaLines) && window.ms365TenantSettingsParseSgaLines) {
                s.sga = window.ms365TenantSettingsParseSgaLines(payload.sgaLines);
                touched++;
            }
            if (normStr(payload.studentCouncilLines) && window.ms365TenantSettingsParseStudentCouncilLines) {
                s.studentCouncil = window.ms365TenantSettingsParseStudentCouncilLines(payload.studentCouncilLines);
                touched++;
            }
            if (normStr(payload.faecherLines) && window.ms365TenantSettingsParseSubjectsLines) {
                s.subjects = window.ms365TenantSettingsParseSubjectsLines(payload.faecherLines);
                touched++;
            }
            if (normStr(payload.argeLines) && window.ms365TenantSettingsParseArgesLines) {
                s.arges = window.ms365TenantSettingsParseArgesLines(payload.argeLines);
                touched++;
            }
            if (normStr(payload.klassenLines) && window.ms365TenantSettingsParseClassesLines) {
                s.classes = window.ms365TenantSettingsParseClassesLines(payload.klassenLines);
                touched++;
            }
            if (touched && typeof window.ms365TenantSettingsSave === 'function') {
                window.ms365TenantSettingsSave(s);
            }
            if (touched) {
                readLists();
                renderSwAdminTableBody();
                fillTeachersTextarea();
                fillStudentsTextarea();
                fillSgaTextarea();
                fillStudentCouncilTextarea();
                fillSubjectsBulkFromSettings();
                fillCatalogSlice('subject');
                fillArgesBulkFromSettings();
                fillCatalogSlice('arge');
                fillClassesBulkTextarea();
                renderClassesTable();
            }
            toast(
                touched
                    ? 'Gesamt-Excel: ' + touched + ' Liste(n) in die Stammdaten übernommen (gespeichert).'
                    : 'Gesamt-Excel: keine Daten erkannt (Blattnamen wie „Schueler“, „Lehrer“, „SGA“, „StudentCouncil“, … und erste Datenzeile prüfen).'
            );
        } catch (e) {
            toast('Gesamt-Import: ' + (e.message || e));
        }
    }

    function wireSchuldatenMasterDownloadClick(btnId) {
        const b = document.getElementById(btnId);
        if (!b) return;
        b.addEventListener('click', function () {
            const api = window.ms365SchuldatenMasterImport;
            if (!api || typeof api.downloadTemplate !== 'function') {
                toast('Vorlagen-Modul fehlt (tenant-settings-ui.js).');
                return;
            }
            if (typeof api.isXlsxReady === 'function' && !api.isXlsxReady()) {
                toast('Excel-Bibliothek noch nicht geladen – Seite kurz warten und erneut versuchen.');
                return;
            }
            if (!api.downloadTemplate()) {
                toast('Vorlage konnte nicht erzeugt werden.');
            }
        });
    }

    function saveTeachersList() {
        const s = loadTenantSettings() || {};
        const tt = document.getElementById('swTeachersLines');
        if (window.ms365TenantSettingsParseTeachersLines && tt) {
            s.teachers = window.ms365TenantSettingsParseTeachersLines(tt.value);
        }
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(s);
        }
        readLists();
        renderSwTeachersTableFromTextarea();
        refreshSwStatsTeachers();
        refreshSwOwnerSummary('Lehrer', 'lehrer');
        toast('Lehrkräfte gespeichert.');
    }

    function saveStudentsList() {
        const s = loadTenantSettings() || {};
        const st = document.getElementById('swStudentsLines');
        if (window.ms365TenantSettingsParseStudentsLines && st) {
            s.students = window.ms365TenantSettingsParseStudentsLines(st.value);
        }
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(s);
        }
        readLists();
        renderSwStudentsTableFromTextarea();
        refreshSwStatsStudents();
        refreshSwOwnerSummary('Schueler', 'schueler');
        toast('Schüler:innen gespeichert.');
    }

    function sgaToLines(rows) {
        return (rows || [])
            .map(function (x) {
                let scope = normStr(x && x.scope);
                if (scope === 'teacher') scope = 'Lehrer';
                else if (scope === 'student') scope = 'Schueler';
                else if (scope === 'external') scope = 'Extern';
                return [scope, normStr(x && x.name), normEmail(x && x.email)].join(';').trim();
            })
            .filter(Boolean)
            .join('\n');
    }

    function getSwSgaFromTextarea() {
        const ta = document.getElementById('swSgaLines');
        if (!ta || typeof window.ms365TenantSettingsParseSgaLines !== 'function') return [];
        return window.ms365TenantSettingsParseSgaLines(ta.value);
    }

    function setSwSgaTextareaFromRows(rows) {
        const ta = document.getElementById('swSgaLines');
        if (ta) ta.value = sgaToLines(rows);
    }

    function fillSgaTextarea() {
        const ta = document.getElementById('swSgaLines');
        const sel = document.getElementById('swSgaMode');
        const s = loadTenantSettings();
        if (ta && s && Array.isArray(s.sga)) ta.value = sgaToLines(s.sga);
        if (sel && s) sel.value = String(s.sgaMode || 'group') === 'distribution' ? 'distribution' : 'group';
        renderSwSgaTableFromTextarea();
        restoreWizardGroupStatus('sga');
    }

    function saveSgaList() {
        const s = loadTenantSettings() || {};
        const ta = document.getElementById('swSgaLines');
        const sel = document.getElementById('swSgaMode');
        if (window.ms365TenantSettingsParseSgaLines && ta) {
            s.sga = window.ms365TenantSettingsParseSgaLines(ta.value);
        }
        s.sgaMode = sel && String(sel.value || '') === 'distribution' ? 'distribution' : 'group';
        if (typeof window.ms365TenantSettingsSave === 'function') window.ms365TenantSettingsSave(s);
        readLists();
        renderSwSgaTableFromTextarea();
        refreshSwStatsSga();
        toast('SGA gespeichert.');
    }

    function studentCouncilToLines(rows) {
        return (rows || [])
            .map(function (x) {
                return [normStr(x && x.klasse), normStr(x && x.name), normEmail(x && x.email)].join(';').trim();
            })
            .filter(Boolean)
            .join('\n');
    }

    function getSwStudentCouncilFromTextarea() {
        const ta = document.getElementById('swStudentCouncilLines');
        if (!ta || typeof window.ms365TenantSettingsParseStudentCouncilLines !== 'function') return [];
        return window.ms365TenantSettingsParseStudentCouncilLines(ta.value);
    }

    function setSwStudentCouncilTextareaFromRows(rows) {
        const ta = document.getElementById('swStudentCouncilLines');
        if (ta) ta.value = studentCouncilToLines(rows);
    }

    function fillStudentCouncilTextarea() {
        const ta = document.getElementById('swStudentCouncilLines');
        const s = loadTenantSettings();
        if (ta && s && Array.isArray(s.studentCouncil)) ta.value = studentCouncilToLines(s.studentCouncil);
        renderSwStudentCouncilTableFromTextarea();
        restoreWizardGroupStatus('studentCouncil');
    }

    function saveStudentCouncilList() {
        const s = loadTenantSettings() || {};
        const ta = document.getElementById('swStudentCouncilLines');
        if (window.ms365TenantSettingsParseStudentCouncilLines && ta) {
            s.studentCouncil = window.ms365TenantSettingsParseStudentCouncilLines(ta.value);
        }
        if (typeof window.ms365TenantSettingsSave === 'function') window.ms365TenantSettingsSave(s);
        readLists();
        renderSwStudentCouncilTableFromTextarea();
        refreshSwStatsStudentCouncil();
        toast('Schülervertretung gespeichert.');
    }

    function renderWizardDirectoryActionCell(tr, row, rerenderKind, onDelete) {
        const tdAct = document.createElement('td');
        tdAct.className = 'action-cell';
        tdAct.style.whiteSpace = 'nowrap';
        const wrap = document.createElement('div');
        wrap.style.cssText = 'display:inline-flex;flex-wrap:nowrap;gap:6px;align-items:center;justify-content:flex-end;';
        const btnMs = document.createElement('button');
        btnMs.type = 'button';
        btnMs.className = 'mini-btn';
        btnMs.style.background = '#5e72e4';
        btnMs.title = 'Diese E‑Mail in Microsoft Entra prüfen';
        btnMs.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
        btnMs.addEventListener('click', function () {
            verifyGraphDirectoryOneEmail(row && row.email, rerenderKind);
        });
        const btnCr = document.createElement('button');
        btnCr.type = 'button';
        btnCr.className = 'mini-btn';
        btnCr.style.background = '#11cdef';
        btnCr.title = 'Neuen Benutzer in Microsoft Entra ID anlegen';
        btnCr.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
        btnCr.addEventListener('click', function () {
            runCreateEntraUserInteractive(row && row.email, row && row.name, rerenderKind);
        });
        const btnDel = document.createElement('button');
        btnDel.type = 'button';
        btnDel.className = 'mini-btn';
        btnDel.textContent = '✕';
        btnDel.title = 'Zeile entfernen';
        btnDel.addEventListener('click', onDelete);
        wrap.appendChild(btnMs);
        wrap.appendChild(btnCr);
        wrap.appendChild(btnDel);
        tdAct.appendChild(wrap);
        return tdAct;
    }

    function renderSwSgaTableFromTextarea() {
        const tbody = document.getElementById('swSgaTableBody');
        if (!tbody) return;
        const rows = getSwSgaFromTextarea();
        tbody.replaceChildren();
        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Noch keine SGA-Mitglieder – oben einfügen oder „+ Zeile“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            refreshSwStatsSga();
            return;
        }
        rows.forEach(function (row, idx) {
            const tr = document.createElement('tr');
            const tdScope = document.createElement('td');
            const sel = document.createElement('select');
            sel.style.width = '100%';
            sel.style.font = 'inherit';
            [
                { value: '', label: '— Gruppe —' },
                { value: 'teacher', label: 'Lehrer' },
                { value: 'student', label: 'Schüler' },
                { value: 'external', label: 'Extern' }
            ].forEach(function (entry) {
                const opt = document.createElement('option');
                opt.value = entry.value;
                opt.textContent = entry.label;
                if ((row.scope || '') === entry.value) opt.selected = true;
                sel.appendChild(opt);
            });
            sel.addEventListener('change', function () {
                const all = getSwSgaFromTextarea();
                if (!all[idx]) return renderSwSgaTableFromTextarea();
                all[idx].scope = normStr(sel.value);
                setSwSgaTextareaFromRows(all);
                renderSwSgaTableFromTextarea();
            });
            tdScope.appendChild(sel);

            const tdName = document.createElement('td');
            tdName.textContent = row.name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdName, row.name, function (next, meta) {
                    const all = getSwSgaFromTextarea();
                    if (!all[idx]) return renderSwSgaTableFromTextarea();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                    setSwSgaTextareaFromRows(all);
                    renderSwSgaTableFromTextarea();
                });
            });

            const tdEmail = document.createElement('td');
            tdEmail.textContent = row.email || '';
            tdEmail.title = 'Doppelklick zum Bearbeiten';
            tdEmail.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdEmail, row.email, function (next, meta) {
                    const all = getSwSgaFromTextarea();
                    if (!all[idx]) return renderSwSgaTableFromTextarea();
                    const prev = all[idx].email;
                    all[idx].email = meta && meta.cancelled ? prev : normEmail(next);
                    setSwSgaTextareaFromRows(all);
                    renderSwSgaTableFromTextarea();
                });
            });

            const tdMs = createSwDirectoryMatchTd(row.email);
            const tdAct = renderWizardDirectoryActionCell(tr, row, 'sga', function () {
                const all = getSwSgaFromTextarea();
                all.splice(idx, 1);
                setSwSgaTextareaFromRows(all);
                renderSwSgaTableFromTextarea();
            });
            tr.appendChild(tdScope);
            tr.appendChild(tdName);
            tr.appendChild(tdEmail);
            tr.appendChild(tdMs);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);
        });
        refreshSwStatsSga();
    }

    function renderSwStudentCouncilTableFromTextarea() {
        const tbody = document.getElementById('swStudentCouncilTableBody');
        if (!tbody) return;
        const rows = getSwStudentCouncilFromTextarea();
        tbody.replaceChildren();
        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Noch keine Einträge – oben einfügen oder „+ Zeile“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            refreshSwStatsStudentCouncil();
            return;
        }
        rows.forEach(function (row, idx) {
            const tr = document.createElement('tr');
            const tdClass = document.createElement('td');
            tdClass.textContent = row.klasse || '';
            tdClass.title = 'Doppelklick zum Bearbeiten';
            tdClass.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdClass, row.klasse, function (next, meta) {
                    const all = getSwStudentCouncilFromTextarea();
                    if (!all[idx]) return renderSwStudentCouncilTableFromTextarea();
                    const prev = all[idx].klasse;
                    all[idx].klasse = meta && meta.cancelled ? prev : normStr(next);
                    setSwStudentCouncilTextareaFromRows(all);
                    renderSwStudentCouncilTableFromTextarea();
                });
            });
            const tdName = document.createElement('td');
            tdName.textContent = row.name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdName, row.name, function (next, meta) {
                    const all = getSwStudentCouncilFromTextarea();
                    if (!all[idx]) return renderSwStudentCouncilTableFromTextarea();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                    setSwStudentCouncilTextareaFromRows(all);
                    renderSwStudentCouncilTableFromTextarea();
                });
            });
            const tdEmail = document.createElement('td');
            tdEmail.textContent = row.email || '';
            tdEmail.title = 'Doppelklick zum Bearbeiten';
            tdEmail.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdEmail, row.email, function (next, meta) {
                    const all = getSwStudentCouncilFromTextarea();
                    if (!all[idx]) return renderSwStudentCouncilTableFromTextarea();
                    const prev = all[idx].email;
                    all[idx].email = meta && meta.cancelled ? prev : normEmail(next);
                    setSwStudentCouncilTextareaFromRows(all);
                    renderSwStudentCouncilTableFromTextarea();
                });
            });
            const tdMs = createSwDirectoryMatchTd(row.email);
            const tdAct = renderWizardDirectoryActionCell(tr, row, 'studentCouncil', function () {
                const all = getSwStudentCouncilFromTextarea();
                all.splice(idx, 1);
                setSwStudentCouncilTextareaFromRows(all);
                renderSwStudentCouncilTableFromTextarea();
            });
            tr.appendChild(tdClass);
            tr.appendChild(tdName);
            tr.appendChild(tdEmail);
            tr.appendChild(tdMs);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);
        });
        refreshSwStatsStudentCouncil();
    }

    function getDirectoryMatchForEmail(emailRaw) {
        const em = normEmail(emailRaw);
        if (!em || em.indexOf('@') === -1) return null;
        try {
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getSetup !== 'function') return null;
            const map = window.ms365AppDataV2.getSetup().directoryMatchByEmail || {};
            return map[em] || null;
        } catch {
            return null;
        }
    }

    function createSwDirectoryMatchTd(emailRaw) {
        const td = document.createElement('td');
        td.className = 'sw-dir-match';
        td.style.fontSize = '0.88em';
        td.style.lineHeight = '1.35';
        const em = normEmail(emailRaw);
        if (!em || em.indexOf('@') === -1) {
            td.style.color = 'var(--muted)';
            td.textContent = '–';
            td.title = 'E‑Mail nötig für Abgleich mit Microsoft Entra';
            return td;
        }
        const m = getDirectoryMatchForEmail(em);
        if (m && m.graphUserId) {
            const gid = String(m.graphUserId);
            const short = gid.length > 14 ? gid.slice(0, 12) + '…' : gid;
            td.innerHTML =
                '<span style="color:#0d8050;font-weight:700;">✓</span> <code style="font-size:0.82em;">' +
                escapeHtml(short) +
                '</code>';
            td.title =
                (m.displayName ? m.displayName : '') +
                (m.userPrincipalName ? '\n' + m.userPrincipalName : '') +
                '\nObject-ID: ' +
                gid;
        } else if (m && m.notFound) {
            td.innerHTML =
                '<span style="color:#856404;font-weight:700;">✗</span> <span style="color:var(--muted)">nicht gefunden</span>';
            td.title = 'Kein Benutzer mit mail oder UPN gleich dieser E‑Mail';
        } else {
            td.style.color = 'var(--muted)';
            td.textContent = '–';
            td.title = 'Noch nicht geprüft – „E‑Mails mit Microsoft prüfen“ oder Symbol in der Aktionsspalte';
        }
        return td;
    }

    async function verifyGraphDirectoryOneEmail(emailRaw, rerenderKind) {
        const em = normEmail(emailRaw);
        if (!em || em.indexOf('@') === -1) {
            toast('E‑Mail fehlt oder ungültig.');
            return;
        }
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') {
            toast('Datenmodul fehlt.');
            return;
        }
        try {
            const token = await G().getGraphToken();
            const u = await G().resolveUserByEmail(token, em);
            const iso = new Date().toISOString();
            const patch = {};
            if (u && u.id) {
                patch[em] = {
                    graphUserId: u.id,
                    displayName: String(u.displayName || '').trim(),
                    userPrincipalName: String(u.userPrincipalName || '').trim(),
                    notFound: false,
                    checkedAt: iso
                };
                window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: patch });
                toast('Microsoft 365: ' + (u.displayName || em));
            } else {
                window.ms365AppDataV2.patchSetup({
                    directoryMatchByEmail: { [em]: { notFound: true, checkedAt: iso } }
                });
                toast('Kein Entra‑Benutzer für: ' + em);
            }
            if (rerenderKind === 'teachers') renderSwTeachersTableFromTextarea();
            else if (rerenderKind === 'students') renderSwStudentsTableFromTextarea();
            else if (rerenderKind === 'sga') renderSwSgaTableFromTextarea();
            else if (rerenderKind === 'studentCouncil') renderSwStudentCouncilTableFromTextarea();
            else if (rerenderKind === 'admin') {
                persistSwAdminListFromTableSilent();
                renderSwAdminTableBody();
            }
        } catch (e) {
            toast('Abgleich: ' + (e.message || e));
        }
    }

    function mailNicknameFromUpn(upn) {
        const local = String(upn || '').split('@')[0] || 'user';
        return G().sanitizeMailNickname(local.replace(/[^a-zA-Z0-9]/g, '') || 'user');
    }

    function patchDirectoryMatchKeys(emailKeys, payload) {
        const patch = {};
        const seen = {};
        (emailKeys || []).forEach(function (k) {
            const em = normEmail(k);
            if (!em || em.indexOf('@') === -1 || seen[em]) return;
            seen[em] = true;
            patch[em] = payload;
        });
        if (!Object.keys(patch).length) return;
        if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
            window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: patch });
        }
    }

    function directoryMatchUserPayload(user, iso) {
        return {
            graphUserId: String(user && user.id ? user.id : '').trim(),
            displayName: String((user && (user.displayName || user.displayNameHint)) || '').trim(),
            userPrincipalName: String((user && (user.userPrincipalName || user.upn)) || '').trim(),
            notFound: false,
            checkedAt: iso || new Date().toISOString()
        };
    }

    async function createEntraUserViaGraph(token, upn, displayName) {
        const pwd = randomTempPassword();
        const mailNick = mailNicknameFromUpn(upn);
        const body = {
            accountEnabled: true,
            displayName: String(displayName).trim(),
            mailNickname: mailNick,
            userPrincipalName: String(upn).trim(),
            passwordProfile: {
                forceChangePasswordNextSignIn: true,
                password: pwd
            }
        };
        const created = await G().graphJson('POST', '/users', token, body, undefined);
        const uid = String(created.id || '').trim();
        if (!uid) throw new Error('Keine Benutzer-ID von Graph erhalten.');
        return {
            id: uid,
            password: pwd,
            upn: String(upn).trim(),
            displayName: String(displayName).trim(),
            mailNickname: mailNick
        };
    }

    function rerenderAfterDirectoryChange(rerenderKind) {
        if (rerenderKind === 'teachers') renderSwTeachersTableFromTextarea();
        else if (rerenderKind === 'students') renderSwStudentsTableFromTextarea();
        else if (rerenderKind === 'sga') renderSwSgaTableFromTextarea();
        else if (rerenderKind === 'studentCouncil') renderSwStudentCouncilTableFromTextarea();
        else if (rerenderKind === 'admin') {
            persistSwAdminListFromTableSilent();
            renderSwAdminTableBody();
        }
    }

    async function runCreateEntraUserInteractive(emailRaw, nameHint, rerenderKind) {
        const em = normEmail(emailRaw);
        if (!em || em.indexOf('@') === -1) {
            toast('Bitte zuerst eine gültige E‑Mail eintragen.');
            return;
        }
        const kind = rerenderKind || 'admin';
        const dom = em.split('@')[1] || '';
        const upn = await dlgPrompt('Benutzer-Principalname (UPN), z. B. vorname.nachname@' + dom + ':', em, {
            title: 'Entra-Benutzer',
            inputLabel: 'UPN'
        });
        if (upn == null || !normStr(upn)) return;
        const displayName = await dlgPrompt(
            'Anzeigename in Microsoft 365:',
            normStr(nameHint) || normStr(String(upn).split('@')[0]),
            { title: 'Entra-Benutzer', inputLabel: 'Anzeigename' }
        );
        if (displayName == null || !normStr(displayName)) return;
        const mailNick = mailNicknameFromUpn(upn);
        try {
            const tokenProbe = await G().getGraphToken();
            const existing = await G().resolveUserByEmail(tokenProbe, em);
            if (existing && existing.id) {
                toast('Unter dieser E‑Mail existiert bereits ein Entra-Benutzer – Abgleich verwenden.');
                verifyGraphDirectoryOneEmail(em, kind);
                return;
            }
        } catch {
            // weiter mit Anlage
        }
        if (
            !(await dlgConfirm(
                'Benutzer in Entra ID anlegen?\n\nAnzeigename: ' +
                    displayName +
                    '\nUPN: ' +
                    upn +
                    '\nMail‑Nickname: ' +
                    mailNick +
                    '\n\nEs wird ein temporäres Kennwort gesetzt (Wechsel beim ersten Anmelden).',
                { title: 'Entra-Benutzer anlegen' }
            ))
        ) {
            return;
        }
        try {
            const token = await G().getGraphToken();
            const created = await createEntraUserViaGraph(token, upn, displayName);
            const iso = new Date().toISOString();
            patchDirectoryMatchKeys(
                [em, created.upn],
                directoryMatchUserPayload(
                    {
                        id: created.id,
                        displayName: created.displayName,
                        userPrincipalName: created.upn
                    },
                    iso
                )
            );
            await dlgAlert(
                'Benutzer angelegt.\n\nEinmaliges Kennwort:\n' +
                    created.password +
                    '\n\n(Bitte sicher übergeben / in Entra ändern.)',
                { title: 'Kennwort notieren', okText: 'Verstanden' }
            );
            rerenderAfterDirectoryChange(kind);
            toast('Entra-Benutzer angelegt.');
        } catch (e) {
            toast('Benutzer anlegen: ' + (e.message || e));
        }
    }

    async function runCreateEntraUserForVerwaltungRow(emailRaw, nameHint) {
        return runCreateEntraUserInteractive(emailRaw, nameHint, 'admin');
    }

    async function runVerifyGraphDirectoryRows(rows, getEmail, label, btnId, onDone) {
        const btn = document.getElementById(btnId);
        const updates = {};
        let found = 0;
        let missed = 0;
        let skipped = 0;
        const seen = new Set();
        const iso = new Date().toISOString();
        try {
            if (btn) {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
            }
            const token = await G().getGraphToken();
            for (let i = 0; i < rows.length; i++) {
                const em = normEmail(getEmail(rows[i]) || '');
                if (!em || em.indexOf('@') === -1) {
                    skipped++;
                    continue;
                }
                if (seen.has(em)) continue;
                seen.add(em);
                try {
                    const u = await G().resolveUserByEmail(token, em);
                    if (u && u.id) {
                        updates[em] = {
                            graphUserId: u.id,
                            displayName: String(u.displayName || '').trim(),
                            userPrincipalName: String(u.userPrincipalName || '').trim(),
                            notFound: false,
                            checkedAt: iso
                        };
                        found++;
                    } else {
                        updates[em] = { notFound: true, checkedAt: iso };
                        missed++;
                    }
                } catch {
                    updates[em] = { notFound: true, checkedAt: iso };
                    missed++;
                }
            }
            if (Object.keys(updates).length && window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: updates });
            }
            toast(label + ': ' + found + ' gefunden, ' + missed + ' nicht gefunden, ' + skipped + ' ohne gültige E‑Mail');
            if (typeof onDone === 'function') onDone();
        } catch (e) {
            toast('Microsoft‑Abgleich: ' + (e.message || e));
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

    async function runVerifyTeachersGraph() {
        const rows = getSwTeachersFromTextarea();
        await runVerifyGraphDirectoryRows(
            rows,
            function (r) {
                return r.email;
            },
            'Lehrkräfte',
            'swBtnVerifyTeachersGraph',
            function () {
                renderSwTeachersTableFromTextarea();
            }
        );
    }

    function personNeedsEntraAccount(row) {
        const em = normEmail(row && row.email);
        if (!em || em.indexOf('@') === -1) return false;
        const m = getDirectoryMatchForEmail(em);
        if (m && m.graphUserId) return false;
        return true;
    }

    async function runCreateMissingEntraUsers(opts) {
        const getRows = opts && opts.getRows;
        const rerenderKind = (opts && opts.rerenderKind) || 'teachers';
        const btnId = opts && opts.btnId;
        const title = (opts && opts.title) || 'Benutzer anlegen';
        const toastPrefix = (opts && opts.toastPrefix) || 'Benutzer';
        const noneMsg =
            (opts && opts.noneMsg) ||
            'Keine fehlenden Personen: alle mit E‑Mail sind bereits Microsoft 365 zugeordnet – oder es fehlt eine E‑Mail.';
        const whoOne = (opts && opts.whoOne) || '1 Person hat';
        const whoMany = opts && opts.whoMany;
        const rows = typeof getRows === 'function' ? getRows() : [];
        const seen = new Set();
        const candidates = [];
        rows.forEach(function (row) {
            if (!personNeedsEntraAccount(row)) return;
            const em = normEmail(row.email);
            if (seen.has(em)) return;
            seen.add(em);
            candidates.push(row);
        });
        if (!candidates.length) {
            toast(noneMsg);
            return;
        }
        const n = candidates.length;
        const who = n === 1 ? whoOne : typeof whoMany === 'function' ? whoMany(n) : n + ' Personen haben';
        if (
            !(await dlgConfirm(
                who +
                    ' noch keinen zugeordneten Microsoft‑365‑Benutzer.\n\n' +
                    'Es werden Konten mit der Listen‑E‑Mail als Anmeldename (UPN) und dem Namen als Anzeigename angelegt. Bereits vorhandene Konten werden nur zugeordnet.\n' +
                    'Temporäre Kennwörter werden danach einmal angezeigt.\n\nFortfahren?',
                { title: title, okText: 'Anlegen' }
            ))
        ) {
            return;
        }
        const btn = btnId ? document.getElementById(btnId) : null;
        const pwdNotes = [];
        const failNotes = [];
        let created = 0;
        let existed = 0;
        let failed = 0;
        try {
            if (btn) {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
            }
            const token = await G().getGraphToken();
            const iso = new Date().toISOString();
            const updates = {};
            for (let i = 0; i < candidates.length; i++) {
                const row = candidates[i];
                const em = normEmail(row.email);
                const displayName = normStr(row.name) || em.split('@')[0];
                try {
                    const existing = await G().resolveUserByEmail(token, em);
                    if (existing && existing.id) {
                        updates[em] = directoryMatchUserPayload(existing, iso);
                        existed++;
                    } else {
                        const createdUser = await createEntraUserViaGraph(token, em, displayName);
                        updates[em] = directoryMatchUserPayload(
                            {
                                id: createdUser.id,
                                displayName: createdUser.displayName,
                                userPrincipalName: createdUser.upn
                            },
                            iso
                        );
                        pwdNotes.push(displayName + ' <' + em + '>: ' + createdUser.password);
                        created++;
                    }
                } catch (e) {
                    failed++;
                    failNotes.push(displayName + ' <' + em + '>: ' + (e.message || e));
                }
                if (i < candidates.length - 1 && typeof G().sleep === 'function') {
                    await G().sleep(450);
                }
            }
            if (Object.keys(updates).length && window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({ directoryMatchByEmail: updates });
            }
            rerenderAfterDirectoryChange(rerenderKind);
            const parts = [];
            if (created) parts.push(created + ' angelegt');
            if (existed) parts.push(existed + ' bereits vorhanden');
            if (failed) parts.push(failed + ' fehlgeschlagen');
            toast(toastPrefix + ': ' + (parts.length ? parts.join(', ') : 'Keine Änderung.'));
            if (pwdNotes.length) {
                await dlgAlert(
                    'Temporäre Kennwörter (bitte sicher notieren und übergeben):\n\n' + pwdNotes.join('\n'),
                    { title: 'Kennwörter notieren', okText: 'Verstanden' }
                );
            }
            if (failNotes.length) {
                await dlgAlert('Nicht angelegt:\n\n' + failNotes.join('\n'), {
                    title: 'Fehler bei der Anlage',
                    okText: 'OK'
                });
            }
        } catch (e) {
            toast(toastPrefix + ' anlegen: ' + (e.message || e));
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

    async function runCreateMissingTeachersEntra() {
        return runCreateMissingEntraUsers({
            getRows: getSwTeachersFromTextarea,
            rerenderKind: 'teachers',
            btnId: 'swBtnCreateMissingTeachers',
            title: 'Lehrkräfte anlegen',
            toastPrefix: 'Lehrkräfte',
            noneMsg:
                'Keine fehlenden Lehrkräfte: alle mit E‑Mail sind bereits Microsoft 365 zugeordnet – oder es fehlt eine E‑Mail.',
            whoOne: '1 Lehrkraft hat',
            whoMany: function (n) {
                return n + ' Lehrkräfte haben';
            }
        });
    }

    async function runCreateMissingStudentsEntra() {
        return runCreateMissingEntraUsers({
            getRows: getSwStudentsFromTextarea,
            rerenderKind: 'students',
            btnId: 'swBtnCreateMissingStudents',
            title: 'Schüler:innen anlegen',
            toastPrefix: 'Schüler:innen',
            noneMsg:
                'Keine fehlenden Schüler:innen: alle mit E‑Mail sind bereits Microsoft 365 zugeordnet – oder es fehlt eine E‑Mail.',
            whoOne: '1 Schüler:in hat',
            whoMany: function (n) {
                return n + ' Schüler:innen haben';
            }
        });
    }

    async function runCreateMissingSgaEntra() {
        return runCreateMissingEntraUsers({
            getRows: getSwSgaFromTextarea,
            rerenderKind: 'sga',
            btnId: 'swBtnCreateMissingSga',
            title: 'SGA-Mitglieder anlegen',
            toastPrefix: 'SGA',
            noneMsg:
                'Keine fehlenden SGA-Mitglieder: alle mit E‑Mail sind bereits Microsoft 365 zugeordnet – oder es fehlt eine E‑Mail.',
            whoOne: '1 SGA-Mitglied hat',
            whoMany: function (n) {
                return n + ' SGA-Mitglieder haben';
            }
        });
    }

    async function runCreateMissingStudentCouncilEntra() {
        return runCreateMissingEntraUsers({
            getRows: getSwStudentCouncilFromTextarea,
            rerenderKind: 'studentCouncil',
            btnId: 'swBtnCreateMissingStudentCouncil',
            title: 'Schülervertretung anlegen',
            toastPrefix: 'Schülervertretung',
            noneMsg:
                'Keine fehlenden Einträge in der Schülervertretung: alle mit E‑Mail sind bereits Microsoft 365 zugeordnet – oder es fehlt eine E‑Mail.',
            whoOne: '1 Eintrag der Schülervertretung hat',
            whoMany: function (n) {
                return n + ' Einträge der Schülervertretung haben';
            }
        });
    }

    async function runVerifyStudentsGraph() {
        const rows = getSwStudentsFromTextarea();
        await runVerifyGraphDirectoryRows(
            rows,
            function (r) {
                return r.email;
            },
            'Schüler:innen',
            'swBtnVerifyStudentsGraph',
            function () {
                renderSwStudentsTableFromTextarea();
            }
        );
    }

    async function runVerifySgaGraph() {
        const rows = getSwSgaFromTextarea();
        await runVerifyGraphDirectoryRows(
            rows,
            function (r) {
                return r.email;
            },
            'SGA',
            'swBtnVerifySgaGraph',
            function () {
                renderSwSgaTableFromTextarea();
            }
        );
    }

    async function runVerifyStudentCouncilGraph() {
        const rows = getSwStudentCouncilFromTextarea();
        await runVerifyGraphDirectoryRows(
            rows,
            function (r) {
                return r.email;
            },
            'Schülervertretung',
            'swBtnVerifyStudentCouncilGraph',
            function () {
                renderSwStudentCouncilTableFromTextarea();
            }
        );
    }

    async function runVerifyVerwaltungGraph() {
        const rows = gatherSwAdminRowsFromTable();
        await runVerifyGraphDirectoryRows(
            rows,
            function (r) {
                return r.email;
            },
            'Verwaltung',
            'swBtnVerifyVerwaltungGraph',
            function () {
                persistSwAdminListFromTableSilent();
                renderSwAdminTableBody();
            }
        );
    }

    function saveDomainStep() {
        const inp = document.getElementById('swDomain');
        const domain = inp ? normStr(inp.value).replace(/^@+/, '') : '';
        if (inp) inp.value = domain;
        if (!domain) {
            toast('Bitte eine Domain ohne @ eingeben, z. B. schule.at.');
            return;
        }
        const s = loadTenantSettings() || {};
        s.domain = domain;
        if (typeof window.ms365SetSchoolDomainNoAt === 'function') {
            window.ms365SetSchoolDomainNoAt(domain);
        }
        let saved = s;
        if (typeof window.ms365TenantSettingsSave === 'function') {
            saved = window.ms365TenantSettingsSave(s) || s;
        }
        const stored = normStr(saved && saved.domain).replace(/^@+/, '') || domain;
        if (inp) inp.value = stored;
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-tenant-settings-changed', {
                    detail: { settings: saved, reason: 'wizard-domain' }
                })
            );
        } catch {
            // ignore
        }
        if (typeof window.ms365RefreshContextBar === 'function') {
            window.ms365RefreshContextBar();
        }
        refreshSwGraphDefaultDomainHint();
        readLists();
        refreshSwSmtpPreviewHint('Lehrer');
        refreshSwSmtpPreviewHint('Schueler');
        refreshSwSmtpPreviewHint('Verwaltung');
        toast('Domain gespeichert: ' + stored);
    }

    function wizardStartCellEdit(td, initialValue, onCommit) {
        const prevText = String(initialValue ?? '');
        const input = document.createElement('input');
        input.className = 'cell-editor';
        input.type = 'text';
        input.value = prevText;
        td.replaceChildren(input);
        input.focus();
        input.select();

        const commit = function () {
            const next = normStr(input.value);
            onCommit(next);
        };
        const cancel = function () {
            onCommit(prevText, { cancelled: true });
        };
        input.addEventListener('keydown', function (e) {
            if (e.key === 'Enter') {
                e.preventDefault();
                commit();
            } else if (e.key === 'Escape') {
                e.preventDefault();
                cancel();
            }
        });
        input.addEventListener('blur', function () {
            commit();
        });
    }

    function swTeachersToLines(rows) {
        return (rows || [])
            .map(function (x) {
                return (
                    normCode(x.code || '') +
                    ';' +
                    normStr(x.name || '') +
                    ';' +
                    normStr(x.email || '').toLowerCase()
                );
            })
            .map(function (s) {
                return s.trim();
            })
            .filter(Boolean)
            .join('\n');
    }

    function getSwTeachersFromTextarea() {
        const ta = document.getElementById('swTeachersLines');
        if (!ta || typeof window.ms365TenantSettingsParseTeachersLines !== 'function') return [];
        return window.ms365TenantSettingsParseTeachersLines(ta.value);
    }

    function setSwTeachersTextareaFromRows(rows) {
        const ta = document.getElementById('swTeachersLines');
        if (!ta) return;
        ta.value = swTeachersToLines(rows);
    }

    function renderSwTeachersTableFromTextarea() {
        const tbody = document.getElementById('swTeachersTableBody');
        if (!tbody) return;
        const rows = getSwTeachersFromTextarea();
        tbody.replaceChildren();

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Noch keine Einträge – oben einfügen, aus Microsoft 365 einlesen, importieren oder „+ Zeile“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            refreshSwStatsTeachers();
            refreshSwOwnerSummary('Lehrer', 'lehrer');
            return;
        }

        rows.forEach(function (row, idx) {
            const tr = document.createElement('tr');

            const tdCode = document.createElement('td');
            tdCode.innerHTML = '<code>' + escapeHtml(row.code || '') + '</code>';
            tdCode.title = 'Doppelklick zum Bearbeiten';
            tdCode.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdCode, row.code, function (next, meta) {
                    const all = getSwTeachersFromTextarea();
                    if (!all[idx]) return renderSwTeachersTableFromTextarea();
                    const prev = all[idx].code;
                    all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                    setSwTeachersTextareaFromRows(all);
                    renderSwTeachersTableFromTextarea();
                });
            });

            const tdName = document.createElement('td');
            tdName.textContent = row.name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdName, row.name, function (next, meta) {
                    const all = getSwTeachersFromTextarea();
                    if (!all[idx]) return renderSwTeachersTableFromTextarea();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                    setSwTeachersTextareaFromRows(all);
                    renderSwTeachersTableFromTextarea();
                });
            });

            const tdEmail = document.createElement('td');
            tdEmail.textContent = row.email || '';
            tdEmail.title = 'Doppelklick zum Bearbeiten';
            tdEmail.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdEmail, row.email, function (next, meta) {
                    const all = getSwTeachersFromTextarea();
                    if (!all[idx]) return renderSwTeachersTableFromTextarea();
                    const prev = all[idx].email;
                    all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                    setSwTeachersTextareaFromRows(all);
                    renderSwTeachersTableFromTextarea();
                });
            });

            const tdMs = createSwDirectoryMatchTd(row.email);

            const tdAction = document.createElement('td');
            tdAction.className = 'action-cell';
            tdAction.style.whiteSpace = 'nowrap';
            const wrap = document.createElement('div');
            wrap.style.cssText =
                'display:inline-flex;flex-wrap:nowrap;gap:6px;align-items:center;justify-content:flex-end;';
            const btnMs = document.createElement('button');
            btnMs.type = 'button';
            btnMs.className = 'mini-btn';
            btnMs.style.background = '#5e72e4';
            btnMs.title = 'Diese E‑Mail in Microsoft Entra prüfen';
            btnMs.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
            btnMs.addEventListener('click', function () {
                verifyGraphDirectoryOneEmail(row.email, 'teachers');
            });
            const btnCr = document.createElement('button');
            btnCr.type = 'button';
            btnCr.className = 'mini-btn';
            btnCr.style.background = '#11cdef';
            btnCr.title = 'Neuen Benutzer in Microsoft Entra ID anlegen (User.ReadWrite.All)';
            btnCr.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
            btnCr.addEventListener('click', function () {
                runCreateEntraUserInteractive(row.email, row.name, 'teachers');
            });
            const btnDel = document.createElement('button');
            btnDel.type = 'button';
            btnDel.className = 'mini-btn';
            btnDel.textContent = '✕';
            btnDel.title = 'Zeile löschen';
            btnDel.addEventListener('click', function () {
                const all = getSwTeachersFromTextarea();
                all.splice(idx, 1);
                setSwTeachersTextareaFromRows(all);
                renderSwTeachersTableFromTextarea();
            });
            wrap.appendChild(btnMs);
            wrap.appendChild(btnCr);
            wrap.appendChild(btnDel);
            tdAction.appendChild(wrap);

            tr.appendChild(tdCode);
            tr.appendChild(tdName);
            tr.appendChild(tdEmail);
            tr.appendChild(tdMs);
            tr.appendChild(tdAction);
            tbody.appendChild(tr);
        });
        refreshSwStatsTeachers();
        refreshSwOwnerSummary('Lehrer', 'lehrer');
    }

    function swStudentsToLines(rows) {
        return (rows || [])
            .map(function (x) {
                return (
                    normStr(x.klasse || '') +
                    ';' +
                    normStr(x.name || '') +
                    ';' +
                    normStr(x.email || '').toLowerCase()
                );
            })
            .map(function (s) {
                return s.trim();
            })
            .filter(Boolean)
            .join('\n');
    }

    function getSwStudentsFromTextarea() {
        const ta = document.getElementById('swStudentsLines');
        if (!ta || typeof window.ms365TenantSettingsParseStudentsLines !== 'function') return [];
        return window.ms365TenantSettingsParseStudentsLines(ta.value);
    }

    function setSwStudentsTextareaFromRows(rows) {
        const ta = document.getElementById('swStudentsLines');
        if (!ta) return;
        ta.value = swStudentsToLines(rows);
    }

    function renderSwStudentsTableFromTextarea() {
        const tbody = document.getElementById('swStudentsTableBody');
        if (!tbody) return;
        const rows = getSwStudentsFromTextarea();
        tbody.replaceChildren();

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Noch keine Einträge – oben einfügen, aus Microsoft 365 einlesen oder „+ Zeile“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            refreshSwStatsStudents();
            refreshSwOwnerSummary('Schueler', 'schueler');
            return;
        }

        rows.forEach(function (row, idx) {
            const tr = document.createElement('tr');

            const tdClass = document.createElement('td');
            tdClass.innerHTML = '<code>' + escapeHtml(row.klasse || '') + '</code>';
            tdClass.title = 'Doppelklick zum Bearbeiten';
            tdClass.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdClass, row.klasse, function (next, meta) {
                    const all = getSwStudentsFromTextarea();
                    if (!all[idx]) return renderSwStudentsTableFromTextarea();
                    const prev = all[idx].klasse;
                    all[idx].klasse = meta && meta.cancelled ? prev : normStr(next);
                    setSwStudentsTextareaFromRows(all);
                    renderSwStudentsTableFromTextarea();
                });
            });

            const tdName = document.createElement('td');
            tdName.textContent = row.name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdName, row.name, function (next, meta) {
                    const all = getSwStudentsFromTextarea();
                    if (!all[idx]) return renderSwStudentsTableFromTextarea();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                    setSwStudentsTextareaFromRows(all);
                    renderSwStudentsTableFromTextarea();
                });
            });

            const tdEmail = document.createElement('td');
            tdEmail.textContent = row.email || '';
            tdEmail.title = 'Doppelklick zum Bearbeiten';
            tdEmail.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdEmail, row.email, function (next, meta) {
                    const all = getSwStudentsFromTextarea();
                    if (!all[idx]) return renderSwStudentsTableFromTextarea();
                    const prev = all[idx].email;
                    all[idx].email = meta && meta.cancelled ? prev : normStr(next).toLowerCase();
                    setSwStudentsTextareaFromRows(all);
                    renderSwStudentsTableFromTextarea();
                });
            });

            const tdMs = createSwDirectoryMatchTd(row.email);

            const tdAction = document.createElement('td');
            tdAction.className = 'action-cell';
            tdAction.style.whiteSpace = 'nowrap';
            const wrap = document.createElement('div');
            wrap.style.cssText =
                'display:inline-flex;flex-wrap:nowrap;gap:6px;align-items:center;justify-content:flex-end;';
            const btnMs = document.createElement('button');
            btnMs.type = 'button';
            btnMs.className = 'mini-btn';
            btnMs.style.background = '#5e72e4';
            btnMs.title = 'Diese E‑Mail in Microsoft Entra prüfen';
            btnMs.innerHTML = '<i class="bi bi-microsoft" aria-hidden="true"></i>';
            btnMs.addEventListener('click', function () {
                verifyGraphDirectoryOneEmail(row.email, 'students');
            });
            const btnCr = document.createElement('button');
            btnCr.type = 'button';
            btnCr.className = 'mini-btn';
            btnCr.style.background = '#11cdef';
            btnCr.title = 'Neuen Benutzer in Microsoft Entra ID anlegen (User.ReadWrite.All)';
            btnCr.innerHTML = '<i class="bi bi-person-plus" aria-hidden="true"></i>';
            btnCr.addEventListener('click', function () {
                runCreateEntraUserInteractive(row.email, row.name, 'students');
            });
            const btnDel = document.createElement('button');
            btnDel.type = 'button';
            btnDel.className = 'mini-btn';
            btnDel.textContent = '✕';
            btnDel.title = 'Zeile löschen';
            btnDel.addEventListener('click', function () {
                const all = getSwStudentsFromTextarea();
                all.splice(idx, 1);
                setSwStudentsTextareaFromRows(all);
                renderSwStudentsTableFromTextarea();
            });
            wrap.appendChild(btnMs);
            wrap.appendChild(btnCr);
            wrap.appendChild(btnDel);
            tdAction.appendChild(wrap);

            tr.appendChild(tdClass);
            tr.appendChild(tdName);
            tr.appendChild(tdEmail);
            tr.appendChild(tdMs);
            tr.appendChild(tdAction);
            tbody.appendChild(tr);
        });
        refreshSwStatsStudents();
        refreshSwOwnerSummary('Schueler', 'schueler');
    }

    function getTenantDomainForPreview() {
        const s = loadTenantSettings() || {};
        return normStr(s.domain || '').replace(/^@+/, '');
    }

    function domainFromEmail(mail) {
        const m = String(mail || '').trim().toLowerCase();
        const i = m.lastIndexOf('@');
        return i >= 0 ? m.slice(i + 1) : '';
    }

    async function refreshSwGraphDefaultDomainHint() {
        const el = document.getElementById('swDomainGraphHint');
        if (!el) return;
        const school = getTenantDomainForPreview();
        el.textContent =
            'Neue Microsoft‑365‑Gruppen erhalten die E‑Mail‑Domain, die in Entra als Standard festgelegt ist – nicht automatisch die Schul‑Domain oben.';
        try {
            const token = await G().getGraphToken();
            const org = await G().graphJson('GET', '/organization?$select=verifiedDomains', token, undefined);
            const row = org && Array.isArray(org.value) ? org.value[0] : org;
            const list = (row && row.verifiedDomains) || [];
            let defName = '';
            list.forEach(function (d) {
                if (d && d.isDefault) defName = String(d.name || '').trim();
            });
            if (!defName) return;
            if (school && defName.toLowerCase() !== school.toLowerCase()) {
                el.innerHTML =
                    'Dieser Tenant legt neue Gruppen‑Mails unter <code>' +
                    escapeHtml(defName) +
                    '</code> an (Entra‑Standarddomain). Die Schul‑Domain <code>' +
                    escapeHtml(school) +
                    '</code> wird dafür von Graph nicht gesetzt. In Microsoft 365 unter <strong>Einstellungen → Domains</strong> die Schul‑Domain als Standard festlegen, wenn Gruppen darunter liegen sollen.';
            } else {
                el.innerHTML =
                    'Tenant‑Standarddomain: <code>' +
                    escapeHtml(defName) +
                    '</code> – neue Gruppen‑Mails sollten darunter liegen.';
            }
        } catch {
            // Hinweis bleibt der Fallback-Text
        }
    }

    function toastGroupCreated(label, group, nick) {
        const actual = String((group && group.mail) || '').trim();
        const school = getTenantDomainForPreview().toLowerCase();
        const actDom = domainFromEmail(actual);
        let msg = (label || 'Gruppe') + ' angelegt';
        if (actual) msg += ': ' + actual;
        else if (nick) msg += ' (Alias ' + nick + ')';
        if (school && actDom && actDom !== school) {
            msg += ' — Graph nutzt die Tenant‑Standarddomain, nicht ' + school + ' aus Schritt 2.';
        }
        toast(msg);
    }

    function sanitizeNickForSmtpPreview(raw) {
        try {
            if (window.ms365GraphUnifiedGroups && typeof window.ms365GraphUnifiedGroups.sanitizeUnifiedGroupMailNickname === 'function') {
                return String(window.ms365GraphUnifiedGroups.sanitizeUnifiedGroupMailNickname(raw) || '').trim();
            }
        } catch {
            // ignore
        }
        return normStr(raw)
            .toLowerCase()
            .replace(/[^a-z0-9._-]/g, '')
            .slice(0, 60);
    }

    function refreshSwSmtpPreviewHint(uiSuffix) {
        const inp = document.getElementById('swNewNick' + uiSuffix);
        const out = document.getElementById('swSmtpPreview' + uiSuffix);
        if (!out) return;
        const nick = inp ? sanitizeNickForSmtpPreview(inp.value) : '';
        const dom = getTenantDomainForPreview();
        if (!nick) {
            out.innerHTML =
                '<span>Ziel‑SMTP nach Speichern der Domain (Schritt 2): </span><code>…@…</code>';
            return;
        }
        if (!dom) {
            out.innerHTML =
                'Ziel‑SMTP (Schul‑Domain): <code>' +
                escapeHtml(nick) +
                '@…</code> <span style="color:var(--muted)">(Schul‑Domain in Schritt 2 speichern)</span>';
            return;
        }
        out.innerHTML =
            'Ziel‑SMTP (Schul‑Domain): <code>' +
            escapeHtml(nick) +
            '@' +
            escapeHtml(dom) +
            '</code> <span style="color:var(--muted)">(Graph legt oft die Tenant‑Standarddomain an)</span>';
    }

    function countSwTeacherListStats() {
        const rows = getSwTeachersFromTextarea();
        let withEm = 0;
        rows.forEach(function (r) {
            const em = normEmail(r && r.email);
            if (em && em.indexOf('@') !== -1) withEm++;
        });
        return { total: rows.length, withEmail: withEm };
    }

    function countSwStudentListStats() {
        const rows = getSwStudentsFromTextarea();
        let withEm = 0;
        rows.forEach(function (r) {
            const em = normEmail(r && r.email);
            if (em && em.indexOf('@') !== -1) withEm++;
        });
        return { total: rows.length, withEmail: withEm };
    }

    function countSwSgaListStats() {
        const rows = getSwSgaFromTextarea();
        let withEm = 0;
        rows.forEach(function (r) {
            const em = normEmail(r && r.email);
            if (em && em.indexOf('@') !== -1) withEm++;
        });
        return { total: rows.length, withEmail: withEm };
    }

    function countSwStudentCouncilListStats() {
        const rows = getSwStudentCouncilFromTextarea();
        let withEm = 0;
        rows.forEach(function (r) {
            const em = normEmail(r && r.email);
            if (em && em.indexOf('@') !== -1) withEm++;
        });
        return { total: rows.length, withEmail: withEm };
    }

    function countSwVerwaltungTableStats() {
        const rows = gatherSwAdminRowsFromTable();
        let withEm = 0;
        rows.forEach(function (r) {
            const em = normEmail(r && r.email);
            if (em && em.indexOf('@') !== -1) withEm++;
        });
        return { total: rows.length, withEmail: withEm };
    }

    function refreshSwStatsSga() {
        const el = document.getElementById('swStatsSga');
        if (!el) return;
        const st = countSwSgaListStats();
        el.textContent = 'SGA: ' + st.total + ' Einträge, davon ' + st.withEmail + ' mit E‑Mail.';
    }

    function refreshSwStatsStudentCouncil() {
        const el = document.getElementById('swStatsStudentCouncil');
        if (!el) return;
        const st = countSwStudentCouncilListStats();
        const y = getCurrentWizardSchoolYearLabel();
        el.textContent =
            'Schülervertretung: ' +
            st.total +
            ' Einträge, davon ' +
            st.withEmail +
            ' mit E‑Mail' +
            (y ? ' — Schuljahr ' + y + '.' : '.');
    }

    function refreshSwStatsTeachers() {
        const el = document.getElementById('swStatsTeachers');
        if (!el) return;
        const st = countSwTeacherListStats();
        el.textContent =
            'Liste: ' +
            st.total +
            ' Lehrkräfte, ' +
            st.withEmail +
            ' mit gültiger E‑Mail (für Synchronisation / Besitzer „alle Lehrkräfte“).';
    }

    function refreshSwStatsStudents() {
        const el = document.getElementById('swStatsStudents');
        if (!el) return;
        const st = countSwStudentListStats();
        el.textContent =
            'Liste: ' +
            st.total +
            ' Schüler:innen, ' +
            st.withEmail +
            ' mit gültiger E‑Mail (für Synchronisation).';
    }

    function refreshSwStatsVerwaltung() {
        const el = document.getElementById('swStatsVerwaltung');
        if (!el) return;
        const st = countSwVerwaltungTableStats();
        el.textContent =
            'Tabelle: ' +
            st.total +
            ' Rollenzeilen, ' +
            st.withEmail +
            ' mit gültiger E‑Mail (für Synchronisation / Besitzer „alle Verwaltungs‑E‑Mails“).';
    }

    function parseManualOwnerEmails(text) {
        const out = [];
        const seen = new Set();
        String(text || '')
            .split(/[\s,;]+/g)
            .forEach(function (p) {
                const em = normEmail(p);
                if (!em || em.indexOf('@') === -1) return;
                if (seen.has(em)) return;
                seen.add(em);
                out.push(em);
            });
        return out;
    }

    function readSlgOwnerDraftFromDom() {
        const lr = document.querySelector('input[name="swOwnerLehrer"]:checked');
        const sr = document.querySelector('input[name="swOwnerSchueler"]:checked');
        const ml = document.getElementById('swOwnerManualLehrer');
        const ms = document.getElementById('swOwnerManualSchueler');
        return {
            slgOwnerSourceLehrer: lr ? String(lr.value || 'direktion') : 'direktion',
            slgOwnerManualEmailsLehrer: ml ? String(ml.value || '') : '',
            slgOwnerSourceSchueler: sr ? String(sr.value || 'direktion') : 'direktion',
            slgOwnerManualEmailsSchueler: ms ? String(ms.value || '') : ''
        };
    }

    function readVerwaltungOwnerDraftFromDom() {
        const r = document.querySelector('input[name="swOwnerVerwaltung"]:checked');
        const ta = document.getElementById('swOwnerManualVerwaltung');
        return {
            vwOwnerSource: r ? String(r.value || 'admin') : 'admin',
            vwOwnerManualEmails: ta ? String(ta.value || '') : ''
        };
    }

    function toggleSwOwnerManualFields() {
        const map = [
            { name: 'swOwnerLehrer', ta: 'swOwnerManualLehrer' },
            { name: 'swOwnerSchueler', ta: 'swOwnerManualSchueler' },
            { name: 'swOwnerVerwaltung', ta: 'swOwnerManualVerwaltung' }
        ];
        map.forEach(function (x) {
            const r = document.querySelector('input[name="' + x.name + '"]:checked');
            const ta = document.getElementById(x.ta);
            if (!ta) return;
            if (r && r.value === 'manual') ta.classList.add('is-visible');
            else ta.classList.remove('is-visible');
        });
    }

    function applySlgOwnerDraftFromSetupToDom() {
        try {
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getSetup !== 'function') return;
            const d = window.ms365AppDataV2.getSetup().slgDraft || {};
            const sl = d.slgOwnerSourceLehrer === 'teachers' || d.slgOwnerSourceLehrer === 'manual' ? d.slgOwnerSourceLehrer : 'direktion';
            const ss = d.slgOwnerSourceSchueler === 'admin' || d.slgOwnerSourceSchueler === 'manual' ? d.slgOwnerSourceSchueler : 'direktion';
            document.querySelectorAll('input[name="swOwnerLehrer"]').forEach(function (inp) {
                inp.checked = inp.value === sl;
            });
            document.querySelectorAll('input[name="swOwnerSchueler"]').forEach(function (inp) {
                inp.checked = inp.value === ss;
            });
            const ml = document.getElementById('swOwnerManualLehrer');
            const ms = document.getElementById('swOwnerManualSchueler');
            if (ml && d.slgOwnerManualEmailsLehrer != null) ml.value = String(d.slgOwnerManualEmailsLehrer);
            if (ms && d.slgOwnerManualEmailsSchueler != null) ms.value = String(d.slgOwnerManualEmailsSchueler);
        } catch {
            // ignore
        }
        toggleSwOwnerManualFields();
    }

    function applyVerwaltungOwnerDraftFromSetupToDom() {
        try {
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getSetup !== 'function') return;
            const vd = window.ms365AppDataV2.getSetup().verwaltungDraft || {};
            const src = vd.vwOwnerSource === 'direktion' || vd.vwOwnerSource === 'manual' ? vd.vwOwnerSource : 'admin';
            document.querySelectorAll('input[name="swOwnerVerwaltung"]').forEach(function (inp) {
                inp.checked = inp.value === src;
            });
            const ta = document.getElementById('swOwnerManualVerwaltung');
            if (ta && vd.vwOwnerManualEmails != null) ta.value = String(vd.vwOwnerManualEmails);
        } catch {
            // ignore
        }
        toggleSwOwnerManualFields();
    }

    function patchSlgOwnerDraftFromDom() {
        try {
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
            const o = readSlgOwnerDraftFromDom();
            window.ms365AppDataV2.patchSetup({ slgDraft: o });
        } catch {
            // ignore
        }
    }

    function patchVerwaltungOwnerDraftFromDom() {
        try {
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
            window.ms365AppDataV2.patchSetup({ verwaltungDraft: readVerwaltungOwnerDraftFromDom() });
        } catch {
            // ignore
        }
    }

    function resolveOwnerEmailsForWizard(kind) {
        readLists();
        const owL = readSlgOwnerDraftFromDom();
        const owV = readVerwaltungOwnerDraftFromDom();
        if (kind === 'lehrer') {
            const src = owL.slgOwnerSourceLehrer;
            if (src === 'teachers') return (swListCache.teachers || []).slice();
            if (src === 'manual') return parseManualOwnerEmails(owL.slgOwnerManualEmailsLehrer);
            return (swListCache.direktion || []).slice();
        }
        if (kind === 'schueler') {
            const src = owL.slgOwnerSourceSchueler;
            if (src === 'admin') {
                const seen = new Set();
                const out = [];
                (swListCache.adminEmails || []).forEach(function (e) {
                    if (e && !seen.has(e)) {
                        seen.add(e);
                        out.push(e);
                    }
                });
                collectEmails(gatherSwAdminRowsFromTable()).forEach(function (e) {
                    if (e && !seen.has(e)) {
                        seen.add(e);
                        out.push(e);
                    }
                });
                return out;
            }
            if (src === 'manual') return parseManualOwnerEmails(owL.slgOwnerManualEmailsSchueler);
            return (swListCache.direktion || []).slice();
        }
        if (kind === 'verwaltung') {
            const src = owV.vwOwnerSource;
            if (src === 'direktion') return (swListCache.direktion || []).slice();
            if (src === 'manual') return parseManualOwnerEmails(owV.vwOwnerManualEmails);
            const seen = new Set();
            const out = [];
            (swListCache.adminEmails || []).forEach(function (e) {
                if (e && !seen.has(e)) {
                    seen.add(e);
                    out.push(e);
                }
            });
            collectEmails(gatherSwAdminRowsFromTable()).forEach(function (e) {
                if (e && !seen.has(e)) {
                    seen.add(e);
                    out.push(e);
                }
            });
            return out;
        }
        return (swListCache.direktion || []).slice();
    }

    function refreshSwOwnerSummary(uiSuffix, kind) {
        const el = document.getElementById('swOwnerSummary' + uiSuffix);
        if (!el) return;
        const list = resolveOwnerEmailsForWizard(kind);
        if (!list.length) {
            el.textContent =
                'Aktuell keine Besitzer‑E‑Mails ermittelt (Quelle prüfen oder manuell eintragen). Microsoft fügt ggf. den angemeldeten Administrator hinzu.';
            return;
        }
        const show = list.slice(0, 4).join(', ');
        const more = list.length > 4 ? ' … +' + String(list.length - 4) : '';
        el.textContent = 'Besitzer: ' + String(list.length) + ' Adresse(n) — ' + show + more;
    }

    function refreshSwWizardAuxiliaryForStep(step) {
        if (step === 3) {
            readLists();
            refreshSwSmtpPreviewHint('Verwaltung');
            refreshSwStatsVerwaltung();
            refreshSwOwnerSummary('Verwaltung', 'verwaltung');
        } else if (step === 4) {
            readLists();
            refreshSwSmtpPreviewHint('Lehrer');
            refreshSwStatsTeachers();
            refreshSwOwnerSummary('Lehrer', 'lehrer');
        } else if (step === 5) {
            readLists();
            refreshSwSmtpPreviewHint('Schueler');
            refreshSwStatsStudents();
            refreshSwOwnerSummary('Schueler', 'schueler');
            refreshWizardYearHints();
        } else if (step === 6) {
            readLists();
            refreshSwStatsSga();
        } else if (step === 7) {
            readLists();
            refreshSwStatsStudentCouncil();
            refreshWizardYearHints();
        } else if (step === 10) {
            refreshWizardYearHints();
        }
    }

    function getCurrentWizardSchoolYearLabel() {
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getContainer === 'function') {
                const c = window.ms365AppDataV2.getContainer();
                return c && c.years && c.years.current ? String(c.years.current) : '';
            }
        } catch {
            // ignore
        }
        return '';
    }

    function refreshWizardYearHints() {
        const y = getCurrentWizardSchoolYearLabel();
        const txt = y
            ? 'Aktives Schuljahr: ' + y + '. Die Daten in diesem Schritt werden genau für dieses Schuljahr gespeichert.'
            : 'Die Daten in diesem Schritt sind schuljahresbezogen.';
        const a = document.getElementById('swYearHintStudents');
        const b = document.getElementById('swYearHintStudentCouncil');
        const c = document.getElementById('swYearHintClasses');
        if (a) a.textContent = txt;
        if (b) b.textContent = txt;
        if (c) c.textContent = txt;
    }

    function wizardGroupStatusCell(kind) {
        return document.getElementById(kind === 'sga' ? 'swSgaGroupMatchCell' : 'swStudentCouncilGroupMatchCell');
    }

    function setWizardGroupStatus(kind, payload, expectedNick) {
        const el = wizardGroupStatusCell(kind);
        if (!el) return;
        el.style.fontSize = '0.88em';
        el.style.lineHeight = '1.35';
        el.style.background = '';
        el.style.color = 'var(--muted)';
        el.title = '';
        if (!payload) {
            el.textContent = 'Noch nicht geprüft';
            return;
        }
        if (payload.loading) {
            el.textContent = 'Prüfe …';
            return;
        }
        if (payload.found && payload.group) {
            const g = payload.group || {};
            const gid = normStr(g.id);
            const short = gid.length > 14 ? gid.slice(0, 12) + '…' : gid;
            const shownNick = expectedNick ? expectedNick : normStr(g.mailNickname);
            el.innerHTML =
                '<span style="color:#0d8050;font-weight:700;">✓</span> <code style="font-size:0.92em;">' +
                escapeHtml(shownNick || '') +
                '</code>';
            el.title =
                (g.displayName ? String(g.displayName) : '') +
                (g.mail ? '\nMail: ' + String(g.mail) : '') +
                (gid ? '\nObject-ID: ' + short : '');
            el.style.background = 'color-mix(in srgb, #0d8050 8%, transparent)';
            el.style.color = 'var(--text)';
            return;
        }
        if (payload.notFound) {
            el.innerHTML =
                '<span style="color:#856404;font-weight:700;">✗</span> <span style="color:var(--muted)">nicht gefunden</span>';
            el.title = 'Keine passende Gruppe in Microsoft 365 gefunden';
            return;
        }
        if (payload.error) {
            el.textContent = 'Fehler';
            el.title = String(payload.error || '');
            return;
        }
        el.textContent = '–';
    }

    function schoolBaseNickForWizard() {
        const s = loadTenantSettings() || {};
        const nm = normStr(s && s.schoolName);
        return nm.replace(/[^a-zA-Z0-9]/g, '').toLowerCase().slice(0, 24);
    }

    function expectedSgaGroupMailNicknameForWizard() {
        const base = schoolBaseNickForWizard();
        if (!base) return '';
        return G().sanitizeUnifiedGroupMailNickname('sga' + base);
    }

    function expectedSgaGroupDisplayNameForWizard() {
        const s = loadTenantSettings() || {};
        const nm = normStr(s && s.schoolName);
        return nm ? 'SGA ' + nm : '';
    }

    function expectedStudentCouncilGroupMailNicknameForWizard() {
        const base = schoolBaseNickForWizard();
        if (!base) return '';
        const yLbl = getCurrentWizardSchoolYearLabel();
        const digits = String(yLbl || '').replace(/\D/g, '').slice(0, 6);
        return G().sanitizeUnifiedGroupMailNickname('sv' + digits + base);
    }

    function expectedStudentCouncilGroupDisplayNameForWizard() {
        const s = loadTenantSettings() || {};
        const nm = normStr(s && s.schoolName);
        const yLbl = getCurrentWizardSchoolYearLabel();
        if (!nm) return '';
        return 'Schülervertretung ' + nm + ' ' + yLbl;
    }

    function restoreWizardGroupStatus(kind) {
        const gid = getGroupIdForKind(kind);
        if (!gid) {
            setWizardGroupStatus(kind, null);
            return;
        }
        if (kind === 'sga') {
            setWizardGroupStatus(
                kind,
                { found: true, group: { id: gid, displayName: expectedSgaGroupDisplayNameForWizard(), mailNickname: expectedSgaGroupMailNicknameForWizard() } },
                expectedSgaGroupMailNicknameForWizard()
            );
        } else {
            setWizardGroupStatus(
                kind,
                { found: true, group: { id: gid, displayName: expectedStudentCouncilGroupDisplayNameForWizard(), mailNickname: expectedStudentCouncilGroupMailNicknameForWizard() } },
                expectedStudentCouncilGroupMailNicknameForWizard()
            );
        }
    }

    async function verifyWizardSchoolGroup(kind) {
        const expectedNick =
            kind === 'sga' ? expectedSgaGroupMailNicknameForWizard() : expectedStudentCouncilGroupMailNicknameForWizard();
        const expectedDn =
            kind === 'sga' ? expectedSgaGroupDisplayNameForWizard() : expectedStudentCouncilGroupDisplayNameForWizard();
        if (!expectedNick || !expectedDn) {
            setGroupIdForKind(kind, null);
            persistMatched();
            setWizardGroupStatus(kind, { notFound: true }, expectedNick || '');
            toast((kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': Bitte zuerst Schulname' + (kind === 'studentCouncil' ? ' und Schuljahr' : '') + ' setzen.');
            return { found: false, notFound: true };
        }
        setWizardGroupStatus(kind, { loading: true }, expectedNick);
        const token = await G().getGraphToken();
        const queries = [expectedNick, expectedDn].filter(Boolean);
        let hits = [];
        for (let i = 0; i < queries.length; i++) {
            const q = queries[i];
            hits = await G().searchUnifiedGroups(token, q);
            if (hits && hits.length) break;
        }
        const storedId = getGroupIdForKind(kind);
        const nickLc = normStr(expectedNick).toLowerCase();
        const matchById =
            storedId && Array.isArray(hits)
                ? hits.find(function (g) {
                    return normStr(g && g.id) === storedId;
                })
                : null;
        const matchByNick =
            Array.isArray(hits) &&
            hits.find(function (g) {
                const mn = normStr(g && g.mailNickname).toLowerCase();
                if (mn && mn === nickLc) return true;
                const mail = normStr(g && g.mail).toLowerCase();
                return mail && mail.startsWith(nickLc + '@');
            });
        const match = matchById || matchByNick || null;
        if (match) {
            setGroupIdForKind(kind, String(match.id || ''));
            persistMatched();
            setWizardGroupStatus(kind, { found: true, group: match }, expectedNick);
            renderSwMatchSummaryForKind(kind, match);
            return { found: true, group: match };
        }
        setGroupIdForKind(kind, null);
        persistMatched();
        setWizardGroupStatus(kind, { notFound: true }, expectedNick);
        renderSwMatchSummaryForKind(kind);
        return { found: false, notFound: true };
    }

    async function createWizardSchoolGroup(kind) {
        const expectedNick =
            kind === 'sga' ? expectedSgaGroupMailNicknameForWizard() : expectedStudentCouncilGroupMailNicknameForWizard();
        const expectedDn =
            kind === 'sga' ? expectedSgaGroupDisplayNameForWizard() : expectedStudentCouncilGroupDisplayNameForWizard();
        if (!expectedNick || !expectedDn) {
            toast((kind === 'sga' ? 'SGA' : 'Schülervertretung') + ': Bitte zuerst Schulname' + (kind === 'studentCouncil' ? ' und Schuljahr' : '') + ' setzen.');
            return null;
        }
        const existing = await verifyWizardSchoolGroup(kind).catch(function () {
            return { found: false };
        });
        if (existing && existing.found && existing.group && existing.group.id) return existing.group;
        setWizardGroupStatus(kind, { loading: true }, expectedNick);
        const token = await G().getGraphToken();
        const desc = kind === 'sga' ? 'MS365-Schulverwaltung – SGA' : 'MS365-Schulverwaltung – Schülervertretung';
        const created = await G().createUnifiedGroup(token, expectedDn, expectedNick, desc);
        if (kind === 'sga') {
            const modeEl = document.getElementById('swSgaMode');
            const mode = modeEl && String(modeEl.value || '') === 'distribution' ? 'distribution' : 'group';
            if (mode === 'group') {
                try {
                    await G().provisionTeamForGroup(token, created && created.id ? created.id : '');
                } catch {
                    // optional
                }
            }
            if (created && created.id) await ensureOwnersForVerwaltungWizard(token, created.id);
        } else if (created && created.id) {
            await ensureOwnersForSlgKind(token, created.id, 'schueler');
        }
        setGroupIdForKind(kind, String((created && created.id) || ''));
        persistMatched();
        setWizardGroupStatus(kind, { found: true, group: created }, expectedNick);
        renderSwMatchSummaryForKind(kind, created);
        toastGroupCreated(kind === 'sga' ? 'SGA-Gruppe' : 'Schülervertretung-Gruppe', created, expectedNick);
        return created;
    }

    /**
     * @param {'subject'|'arge'} sliceKind
     */
    function sanitizeWizardMailPrefix(raw, maxLen) {
        const api = window.ms365AppDataV2 && window.ms365AppDataV2.mailNicknamePrefixSanitize;
        if (typeof api === 'function') {
            return api(raw, maxLen);
        }
        return String(raw ?? '')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9._-]/g, '')
            .slice(0, maxLen || 24);
    }

    function getWizardMailPrefixFromDomOrSetup(sliceKind) {
        const fb = sliceKind === 'arge' ? 'ag' : 'fach';
        const id = sliceKind === 'arge' ? 'swArgeGroupPrefix' : 'swSubjectGroupPrefix';
        const inp = document.getElementById(id);
        if (inp) {
            const t = sanitizeWizardMailPrefix(inp.value, 24);
            if (t) return t;
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                const raw = sliceKind === 'arge' ? su.argeGroupMailPrefix : su.subjectGroupMailPrefix;
                const t = sanitizeWizardMailPrefix(raw, 24);
                if (t) return t;
            }
        } catch {
            // ignore
        }
        return fb;
    }

    /**
     * @param {'subject'|'arge'} sliceKind
     */
    function deriveCatalogMailNickname(sliceKind, code) {
        const pre = getWizardMailPrefixFromDomOrSetup(sliceKind);
        const tail = G().sanitizeMailNickname(String(code || 'x')).slice(0, 40);
        /** Wie bei createUnifiedGroup: einheitliche Graph-Regel, damit Vorschau = tatsächlicher mailNickname (inkl. Präfix mit . - _). */
        return G().sanitizeUnifiedGroupMailNickname(String(pre + tail).toLowerCase()).slice(0, 60);
    }

    /**
     * @param {'subject'|'arge'} sliceKind
     */
    function buildSmtpPreviewPartsForRow(sliceKind, code, link) {
        const dom = getTenantDomainForPreview();
        let nick = '';
        if (link && normStr(link.graphGroupId) && normStr(link.mailNickname)) {
            nick = normStr(link.mailNickname);
        } else {
            nick = deriveCatalogMailNickname(sliceKind, code);
        }
        const smtp = dom && nick ? nick + '@' + dom : nick || '–';
        return { nick: nick, domain: dom, smtp: smtp };
    }

    /**
     * @param {'subject'|'arge'} sliceKind
     */
    function smtpPreviewCellHtml(sliceKind, code, link) {
        const p = buildSmtpPreviewPartsForRow(sliceKind, code, link);
        const isLinked = !!(link && normStr(link.graphGroupId) && normStr(link.mailNickname));
        const title = isLinked
            ? 'Verknüpfte Gruppe (Mail‑Nickname @ Schul‑Domain; tatsächliche Graph‑Mail kann die Tenant‑Standarddomain nutzen)'
            : 'Ziel‑SMTP beim Anlegen (Präfix + Kürzel @ Schul‑Domain). Graph setzt oft die Tenant‑Standarddomain.';
        return (
            '<td class="sw-smtp-preview" title="' +
            escapeHtml(title) +
            '"><code style="font-size:0.88em;">' +
            escapeHtml(p.smtp) +
            '</code></td>'
        );
    }

    function readGroupPrefixesFromSetupToDom() {
        try {
            if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getSetup !== 'function') return;
            const su = window.ms365AppDataV2.getSetup();
            const ins = document.getElementById('swSubjectGroupPrefix');
            const ina = document.getElementById('swArgeGroupPrefix');
            if (ins) ins.value = su.subjectGroupMailPrefix || 'fach';
            if (ina) ina.value = su.argeGroupMailPrefix || 'ag';
        } catch {
            // ignore
        }
    }

    function persistSubjectGroupPrefixFromDom() {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
        const ins = document.getElementById('swSubjectGroupPrefix');
        if (!ins) return;
        const v = sanitizeWizardMailPrefix(ins.value, 24) || 'fach';
        window.ms365AppDataV2.patchSetup({ subjectGroupMailPrefix: v });
    }

    function persistArgeGroupPrefixFromDom() {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.patchSetup !== 'function') return;
        const ina = document.getElementById('swArgeGroupPrefix');
        if (!ina) return;
        const v = sanitizeWizardMailPrefix(ina.value, 24) || 'ag';
        window.ms365AppDataV2.patchSetup({ argeGroupMailPrefix: v });
    }

    function swSubjectsToLines(rows) {
        return (rows || [])
            .map(function (x) {
                return normCode(x.code || '') + ';' + normStr(x.name || '');
            })
            .map(function (s) {
                return s.trim();
            })
            .filter(function (s) {
                return s.length > 0;
            })
            .join('\n');
    }

    function getSwSubjectsFromTextarea() {
        const ta = document.getElementById('swSubjectsBulk');
        if (!ta || typeof window.ms365TenantSettingsParseSubjectsLines !== 'function') return [];
        return window.ms365TenantSettingsParseSubjectsLines(ta.value);
    }

    function setSwSubjectsTextareaFromRows(rows) {
        const ta = document.getElementById('swSubjectsBulk');
        if (!ta) return;
        ta.value = swSubjectsToLines(rows);
    }

    function swArgesToLines(rows) {
        return (rows || [])
            .map(function (x) {
                let line = normCode(x.code || '') + ';' + normStr(x.name || '');
                const subj = Array.isArray(x.subjects) ? x.subjects.filter(Boolean) : [];
                if (subj.length) line += ';' + subj.join(',');
                return line.trim();
            })
            .map(function (s) {
                return s.trim();
            })
            .filter(function (s) {
                return s.length > 0;
            })
            .join('\n');
    }

    function getSwArgesFromTextarea() {
        const ta = document.getElementById('swArgesBulk');
        if (!ta || typeof window.ms365TenantSettingsParseArgesLines !== 'function') return [];
        return window.ms365TenantSettingsParseArgesLines(ta.value);
    }

    function setSwArgesTextareaFromRows(rows) {
        const ta = document.getElementById('swArgesBulk');
        if (!ta) return;
        ta.value = swArgesToLines(rows);
    }

    function renderSwSubjectsUnifiedTable() {
        const tbody = document.getElementById('swSubjectsUnifiedBody');
        if (!tbody) return;
        const rows = getSwSubjectsFromTextarea();
        tbody.replaceChildren();

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Noch keine Fächer – oben einfügen oder „+ Zeile“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        rows.forEach(function (row, idx) {
            const tr = document.createElement('tr');
            const code = row.code;
            const name = row.name;
            const linkSub = getCatalogLink('subject', code || '');

            const tdCode = document.createElement('td');
            tdCode.innerHTML = '<code>' + escapeHtml(code || '') + '</code>';
            tdCode.title = 'Doppelklick zum Bearbeiten';
            tdCode.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdCode, code, function (next, meta) {
                    const all = getSwSubjectsFromTextarea();
                    if (!all[idx]) return renderSwSubjectsUnifiedTable();
                    const prev = all[idx].code;
                    all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                    setSwSubjectsTextareaFromRows(all);
                    renderSwSubjectsUnifiedTable();
                });
            });

            const tdName = document.createElement('td');
            tdName.textContent = name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdName, name, function (next, meta) {
                    const all = getSwSubjectsFromTextarea();
                    if (!all[idx]) return renderSwSubjectsUnifiedTable();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                    setSwSubjectsTextareaFromRows(all);
                    renderSwSubjectsUnifiedTable();
                });
            });

            const tdPrev = document.createElement('td');
            tdPrev.className = 'sw-smtp-preview';
            const parts = buildSmtpPreviewPartsForRow('subject', code, linkSub);
            tdPrev.title =
                linkSub && normStr(linkSub.graphGroupId) ? 'Verknüpfte Gruppe (Mail‑Nickname)' : 'Vorschau beim Anlegen';
            tdPrev.innerHTML = '<code style="font-size:0.88em;">' + escapeHtml(parts.smtp) + '</code>';

            const tdM365 = document.createElement('td');
            tdM365.className = 'sw-catalog-grp';
            if (linkSub && normStr(linkSub.graphGroupId)) {
                tdM365.innerHTML = '<code style="font-size:0.82em;">' + escapeHtml(linkSub.graphGroupId) + '</code>';
            } else {
                tdM365.innerHTML = '<span style="color:var(--muted)">–</span>';
            }

            const tdAct = document.createElement('td');
            tdAct.className = 'action-cell';
            tdAct.style.whiteSpace = 'nowrap';
            const wrap = document.createElement('div');
            wrap.style.cssText =
                'display:inline-flex;flex-wrap:nowrap;gap:6px;align-items:center;justify-content:flex-end;';
            const btnS = document.createElement('button');
            btnS.type = 'button';
            btnS.className = 'mini-btn';
            btnS.style.background = '#5e72e4';
            btnS.title = 'Bestehende Microsoft‑365‑Gruppe suchen und verknüpfen';
            btnS.setAttribute('aria-label', 'Suchen');
            btnS.innerHTML = '<i class="bi bi-search" aria-hidden="true"></i>';
            btnS.addEventListener('click', function () {
                runCatalogSearch('subject', code, name);
            });
            const btnC = document.createElement('button');
            btnC.type = 'button';
            btnC.className = 'mini-btn';
            btnC.style.background = '#2dce89';
            btnC.title = 'Neue Microsoft‑365‑Gruppe anlegen';
            btnC.setAttribute('aria-label', 'Anlegen');
            btnC.innerHTML = '<i class="bi bi-plus-circle" aria-hidden="true"></i>';
            btnC.addEventListener('click', function () {
                runCatalogCreate('subject', code, name);
            });
            const btnDel = document.createElement('button');
            btnDel.type = 'button';
            btnDel.className = 'mini-btn';
            btnDel.title = 'Zeile löschen';
            btnDel.setAttribute('aria-label', 'Zeile löschen');
            btnDel.innerHTML = '<i class="bi bi-x-lg" aria-hidden="true"></i>';
            btnDel.addEventListener('click', function () {
                const all = getSwSubjectsFromTextarea();
                all.splice(idx, 1);
                setSwSubjectsTextareaFromRows(all);
                renderSwSubjectsUnifiedTable();
            });
            wrap.appendChild(btnS);
            wrap.appendChild(btnC);
            wrap.appendChild(btnDel);
            tdAct.appendChild(wrap);

            tr.appendChild(tdCode);
            tr.appendChild(tdName);
            tr.appendChild(tdPrev);
            tr.appendChild(tdM365);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);
        });
    }

    function renderSwArgesUnifiedTable() {
        const tbody = document.getElementById('swArgesUnifiedBody');
        if (!tbody) return;
        const rows = getSwArgesFromTextarea();
        tbody.replaceChildren();

        if (!rows.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 5;
            td.style.color = 'var(--muted)';
            td.textContent = 'Noch keine Arbeitsgruppen – oben einfügen oder „+ Zeile“.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        rows.forEach(function (row, idx) {
            const tr = document.createElement('tr');
            const code = row.code;
            const name = row.name;
            const linkAr = getCatalogLink('arge', code || '');

            const tdCode = document.createElement('td');
            tdCode.innerHTML = '<code>' + escapeHtml(code || '') + '</code>';
            tdCode.title = 'Doppelklick zum Bearbeiten';
            tdCode.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdCode, code, function (next, meta) {
                    const all = getSwArgesFromTextarea();
                    if (!all[idx]) return renderSwArgesUnifiedTable();
                    const prev = all[idx].code;
                    all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                    setSwArgesTextareaFromRows(all);
                    renderSwArgesUnifiedTable();
                });
            });

            const tdName = document.createElement('td');
            tdName.textContent = name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdName, name, function (next, meta) {
                    const all = getSwArgesFromTextarea();
                    if (!all[idx]) return renderSwArgesUnifiedTable();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                    setSwArgesTextareaFromRows(all);
                    renderSwArgesUnifiedTable();
                });
            });

            const tdPrev = document.createElement('td');
            tdPrev.className = 'sw-smtp-preview';
            const partsAr = buildSmtpPreviewPartsForRow('arge', code, linkAr);
            tdPrev.title =
                linkAr && normStr(linkAr.graphGroupId) ? 'Verknüpfte Gruppe (Mail‑Nickname)' : 'Vorschau beim Anlegen';
            tdPrev.innerHTML = '<code style="font-size:0.88em;">' + escapeHtml(partsAr.smtp) + '</code>';

            const tdM365 = document.createElement('td');
            tdM365.className = 'sw-catalog-grp';
            if (linkAr && normStr(linkAr.graphGroupId)) {
                tdM365.innerHTML = '<code style="font-size:0.82em;">' + escapeHtml(linkAr.graphGroupId) + '</code>';
            } else {
                tdM365.innerHTML = '<span style="color:var(--muted)">–</span>';
            }

            const tdAct = document.createElement('td');
            tdAct.className = 'action-cell';
            tdAct.style.whiteSpace = 'nowrap';
            const wrapAr = document.createElement('div');
            wrapAr.style.cssText =
                'display:inline-flex;flex-wrap:nowrap;gap:6px;align-items:center;justify-content:flex-end;';
            const btnS = document.createElement('button');
            btnS.type = 'button';
            btnS.className = 'mini-btn';
            btnS.style.background = '#5e72e4';
            btnS.title = 'Bestehende Microsoft‑365‑Gruppe suchen und verknüpfen';
            btnS.setAttribute('aria-label', 'Suchen');
            btnS.innerHTML = '<i class="bi bi-search" aria-hidden="true"></i>';
            btnS.addEventListener('click', function () {
                runCatalogSearch('arge', code, name);
            });
            const btnC = document.createElement('button');
            btnC.type = 'button';
            btnC.className = 'mini-btn';
            btnC.style.background = '#2dce89';
            btnC.title = 'Neue Microsoft‑365‑Gruppe anlegen';
            btnC.setAttribute('aria-label', 'Anlegen');
            btnC.innerHTML = '<i class="bi bi-plus-circle" aria-hidden="true"></i>';
            btnC.addEventListener('click', function () {
                runCatalogCreate('arge', code, name);
            });
            const btnDel = document.createElement('button');
            btnDel.type = 'button';
            btnDel.className = 'mini-btn';
            btnDel.title = 'Zeile löschen';
            btnDel.setAttribute('aria-label', 'Zeile löschen');
            btnDel.innerHTML = '<i class="bi bi-x-lg" aria-hidden="true"></i>';
            btnDel.addEventListener('click', function () {
                const all = getSwArgesFromTextarea();
                all.splice(idx, 1);
                setSwArgesTextareaFromRows(all);
                renderSwArgesUnifiedTable();
            });
            wrapAr.appendChild(btnS);
            wrapAr.appendChild(btnC);
            wrapAr.appendChild(btnDel);
            tdAct.appendChild(wrapAr);

            tr.appendChild(tdCode);
            tr.appendChild(tdName);
            tr.appendChild(tdPrev);
            tr.appendChild(tdM365);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);
        });
    }

    /**
     * @param {'subject'|'arge'} sliceKind
     */
    function scoreCatalogGroupMatch(sliceKind, g, code, name) {
        const c = normCode(code);
        const n = normStr(name).toLowerCase();
        const nick = String(g.mailNickname || '')
            .toLowerCase()
            .replace(/[^a-z0-9]/g, '');
        const dn = String(g.displayName || '').toLowerCase();
        const mail = String(g.mail || '').toLowerCase();
        let s = 0;
        if (!c && !n) return 0;
        try {
            const expectedNick = deriveCatalogMailNickname(sliceKind, c)
                .toLowerCase()
                .replace(/[^a-z0-9]/g, '');
            if (expectedNick && nick === expectedNick) s += 92;
            else if (c.length >= 2) {
                const cCompact = c.toLowerCase().replace(/[^a-z0-9]/g, '');
                if (cCompact && nick.indexOf(cCompact) !== -1) s += 36;
            }
        } catch {
            // ignore
        }
        if (n.length >= 2) {
            if (dn === n) s += 80;
            else if (dn.indexOf(n) !== -1) s += 50;
            const firstWord = (n.split(/\s+/)[0] || '').trim();
            if (firstWord.length >= 4 && dn.indexOf(firstWord) !== -1) s += 22;
            const nCompact = n.replace(/\s+/g, '');
            if (nCompact.length >= 4 && nick.indexOf(nCompact.slice(0, 14)) !== -1) s += 40;
            if (firstWord.length >= 4 && mail.indexOf(firstWord) !== -1) s += 24;
        }
        if (c.length >= 2 && dn.indexOf(c.toLowerCase()) !== -1) s += 16;
        return Math.min(100, s);
    }

    /**
     * @param {'subject'|'arge'} sliceKind
     */
    function pickBestAutomatchGroup(groups, code, name, sliceKind) {
        if (!groups || !groups.length) return null;
        const scored = groups.map(function (g) {
            return { g: g, s: scoreCatalogGroupMatch(sliceKind, g, code, name) };
        });
        scored.sort(function (a, b) {
            return b.s - a.s;
        });
        const top = scored[0];
        if (!top || top.s < 44) return null;
        const second = scored[1];
        if (second && second.s >= top.s - 8 && second.s >= 40) return null;
        return top.g;
    }

    async function collectUnifiedGroupsForCatalogSearch(token, code, name) {
        const queries = [];
        const n = normStr(name);
        const c = normCode(code);
        if (n) queries.push(n);
        if (c.length >= 2 && queries.indexOf(c) === -1) queries.push(c);
        if (n && c && queries.indexOf(n + ' ' + c) === -1) queries.push(n + ' ' + c);
        const seen = new Map();
        for (let i = 0; i < queries.length; i++) {
            const list = await G().searchUnifiedGroups(token, queries[i]);
            (list || []).forEach(function (g) {
                if (g && g.id && !seen.has(g.id)) seen.set(g.id, g);
            });
            if (seen.size >= 24) break;
        }
        const out = [];
        seen.forEach(function (g) {
            out.push(g);
        });
        return out;
    }

    async function runSubjectsAutomatch() {
        const btn = document.getElementById('swBtnSubjectsAutoMatch');
        const rows = getSwSubjectsFromTextarea();
        if (!rows.length) {
            toast('Keine Fächer in der Liste.');
            return;
        }
        let linked = 0;
        let skipped = 0;
        let noPick = 0;
        try {
            if (btn) {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
            }
            const token = await G().getGraphToken();
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const code = row.code;
                const name = row.name;
                const existing = getCatalogLink('subject', code);
                if (existing && normStr(existing.graphGroupId)) {
                    skipped++;
                    continue;
                }
                const groups = await collectUnifiedGroupsForCatalogSearch(token, code, name);
                const pick = pickBestAutomatchGroup(groups, code, name, 'subject');
                if (!pick) {
                    noPick++;
                    continue;
                }
                upsertCatalogLink({
                    kind: 'subject',
                    code: code,
                    graphGroupId: pick.id,
                    displayName: pick.displayName || '',
                    mailNickname: pick.mailNickname || '',
                    mode: 'matched'
                });
                linked++;
            }
            renderSwSubjectsUnifiedTable();
            const parts = [];
            if (linked) parts.push(linked + ' neu verknüpft');
            if (skipped) parts.push(skipped + ' unverändert (schon verknüpft)');
            if (noPick) parts.push(noPick + ' ohne sicheren Treffer');
            toast('Auto‑Match: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'));
        } catch (e) {
            toast('Auto‑Match: ' + (e.message || e));
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

    async function runArgesAutomatch() {
        const btn = document.getElementById('swBtnArgesAutoMatch');
        const rows = getSwArgesFromTextarea();
        if (!rows.length) {
            toast('Keine Arbeitsgruppen in der Liste.');
            return;
        }
        let linked = 0;
        let skipped = 0;
        let noPick = 0;
        try {
            if (btn) {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
            }
            const token = await G().getGraphToken();
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const code = row.code;
                const name = row.name;
                const existing = getCatalogLink('arge', code);
                if (existing && normStr(existing.graphGroupId)) {
                    skipped++;
                    continue;
                }
                const groups = await collectUnifiedGroupsForCatalogSearch(token, code, name);
                const pick = pickBestAutomatchGroup(groups, code, name, 'arge');
                if (!pick) {
                    noPick++;
                    continue;
                }
                upsertCatalogLink({
                    kind: 'arge',
                    code: code,
                    graphGroupId: pick.id,
                    displayName: pick.displayName || '',
                    mailNickname: pick.mailNickname || '',
                    mode: 'matched'
                });
                linked++;
            }
            renderSwArgesUnifiedTable();
            const parts = [];
            if (linked) parts.push(linked + ' neu verknüpft');
            if (skipped) parts.push(skipped + ' unverändert (schon verknüpft)');
            if (noPick) parts.push(noPick + ' ohne sicheren Treffer');
            toast('Auto‑Match ARGE: ' + (parts.length ? parts.join(' · ') : 'Keine Änderung.'));
        } catch (e) {
            toast('Auto‑Match ARGE: ' + (e.message || e));
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

    function getCatalogLink(kind, code) {
        const c = normCode(code);
        const setup = window.ms365AppDataV2 && window.ms365AppDataV2.getSetup();
        const links = (setup && setup.catalogLinks) || [];
        return links.find(function (x) {
            return x.kind === kind && normCode(x.code) === c;
        });
    }

    function upsertCatalogLink(entry) {
        if (!window.ms365AppDataV2 || typeof window.ms365AppDataV2.getSetup !== 'function') return;
        const cur = window.ms365AppDataV2.getSetup();
        const links = Array.isArray(cur.catalogLinks) ? cur.catalogLinks.slice() : [];
        const code = normCode(entry.code);
        const kind = entry.kind === 'arge' ? 'arge' : 'subject';
        let idx = links.findIndex(function (x) {
            return x.kind === kind && normCode(x.code) === code;
        });
        const row = {
            kind: kind,
            code: code,
            graphGroupId: entry.graphGroupId ? String(entry.graphGroupId) : '',
            displayName: String(entry.displayName || ''),
            mailNickname: String(entry.mailNickname || ''),
            mode: entry.mode === 'matched' || entry.mode === 'created' ? entry.mode : '',
            syncStatus: String(entry.syncStatus || '')
        };
        if (idx >= 0) links[idx] = row;
        else links.push(row);
        window.ms365AppDataV2.patchSetup({ catalogLinks: links });
    }

    /** @param {'subject'|'arge'} sliceKind */
    function fillCatalogSlice(sliceKind) {
        if (sliceKind === 'subject') {
            renderSwSubjectsUnifiedTable();
            return;
        }
        renderSwArgesUnifiedTable();
    }

    function saveSubjectsBulk() {
        const s = loadTenantSettings() || {};
        const subEl = document.getElementById('swSubjectsBulk');
        if (window.ms365TenantSettingsParseSubjectsLines && subEl) {
            s.subjects = window.ms365TenantSettingsParseSubjectsLines(subEl.value);
        }
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(s);
        }
        toast('Fächer gespeichert.');
        fillCatalogSlice('subject');
    }

    function saveArgesBulk() {
        const s = loadTenantSettings() || {};
        const arEl = document.getElementById('swArgesBulk');
        if (window.ms365TenantSettingsParseArgesLines && arEl) {
            s.arges = window.ms365TenantSettingsParseArgesLines(arEl.value);
        }
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(s);
        }
        toast('Arbeitsgruppen gespeichert.');
        fillCatalogSlice('arge');
    }

    function fillSubjectsBulkFromSettings() {
        const s = loadTenantSettings() || {};
        const subEl = document.getElementById('swSubjectsBulk');
        if (subEl && Array.isArray(s.subjects)) {
            subEl.value = s.subjects.map(function (x) {
                return (x.code || '') + ';' + (x.name || '');
            }).join('\n');
        }
    }

    function fillArgesBulkFromSettings() {
        const s = loadTenantSettings() || {};
        const arEl = document.getElementById('swArgesBulk');
        if (arEl && Array.isArray(s.arges)) {
            arEl.value = swArgesToLines(s.arges);
        }
    }

    function fillCatalogTextareas() {
        fillSubjectsBulkFromSettings();
        fillCatalogSlice('subject');
        fillArgesBulkFromSettings();
        fillCatalogSlice('arge');
    }

    async function runCatalogSearch(kind, code, name) {
        const defQ = normStr(name) || normStr(code);
        const q = await dlgPrompt('Gruppe suchen (Name, Mail, Alias):', defQ, {
            title: 'Gruppe suchen',
            inputLabel: 'Suchbegriff (Name, Mail, Alias)'
        });
        if (q == null || !normStr(q)) return;
        try {
            const token = await G().getGraphToken();
            const list = await G().searchUnifiedGroups(token, q);
            if (!list.length) {
                toast('Keine Gruppe gefunden.');
                return;
            }
            const pick = list[0];
            if (
                !(await dlgConfirm(
                    'Diese Gruppe verknüpfen?\n\n' + (pick.displayName || '') + '\n' + (pick.mailNickname || ''),
                    { title: 'Verknüpfung', okText: 'Verknüpfen' }
                ))
            ) {
                return;
            }
            upsertCatalogLink({
                kind: kind,
                code: code,
                graphGroupId: pick.id,
                displayName: pick.displayName || '',
                mailNickname: pick.mailNickname || '',
                mode: 'matched'
            });
            toast('Verknüpft.');
            fillCatalogSlice(kind === 'arge' ? 'arge' : 'subject');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function swClassesToLines(rows) {
        return (rows || [])
            .map(function (cl) {
                const code = normCode(cl.code || '');
                const yearRaw = normStr(cl.year || '');
                const year = /^\d{4}$/.test(yearRaw) ? yearRaw : '';
                const name = normStr(cl.name || '');
                const headName = normStr(cl.headName || '');
                const headEmail = normStr(cl.headEmail || '').toLowerCase();
                let line = code + ';' + year + ';' + name;
                if (headName || headEmail) {
                    line += ';' + headName + ';' + headEmail;
                }
                return line.trim();
            })
            .filter(function (s) {
                return s.replace(/;/g, '').trim().length > 0;
            })
            .join('\n');
    }

    function getSwClassesFromTextarea() {
        const ta = document.getElementById('swClassesBulk');
        if (!ta || typeof window.ms365TenantSettingsParseClassesLines !== 'function') return [];
        return window.ms365TenantSettingsParseClassesLines(ta.value);
    }

    function setSwClassesTextareaFromRows(rows) {
        const ta = document.getElementById('swClassesBulk');
        if (!ta) return;
        ta.value = swClassesToLines(rows);
    }

    function fillClassesBulkTextarea() {
        const ta = document.getElementById('swClassesBulk');
        const s = loadTenantSettings();
        if (!ta || !s || !Array.isArray(s.classes)) return;
        ta.value = swClassesToLines(s.classes);
    }

    function buildClassSmtpPreview(nick) {
        const dom = getTenantDomainForPreview();
        const n = String(nick || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase()
            .slice(0, 60);
        const smtp = dom && n ? n + '@' + dom : n || '–';
        return { nick: n, smtp: smtp };
    }

    function getClassTeamFromContainer(stableNick) {
        const nick = String(stableNick || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase();
        if (!nick || !window.ms365AppDataV2 || typeof window.ms365AppDataV2.getContainer !== 'function') return null;
        const c = window.ms365AppDataV2.getContainer();
        const teams = Array.isArray(c.core && c.core.classTeams) ? c.core.classTeams : [];
        for (let i = 0; i < teams.length; i++) {
            if (teams[i].stableMailNickname === nick) return teams[i];
        }
        return null;
    }

    function saveClassesBulk() {
        const s = loadTenantSettings() || {};
        const ta = document.getElementById('swClassesBulk');
        if (window.ms365TenantSettingsParseClassesLines && ta) {
            s.classes = window.ms365TenantSettingsParseClassesLines(ta.value);
        }
        if (typeof window.ms365TenantSettingsSave === 'function') {
            window.ms365TenantSettingsSave(s);
        }
        toast('Klassen gespeichert.');
        renderClassesTable();
    }

    function renderClassesTable() {
        const tbody = document.getElementById('swClassesBody');
        if (!tbody) return;
        tbody.replaceChildren();
        const classes = getSwClassesFromTextarea();

        if (!classes.length) {
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 6;
            td.style.color = 'var(--muted)';
            td.innerHTML =
                'Noch keine Klassen – oben einfügen, „+ Zeile“ oder in den <a href="tenant.html">Stammdaten</a> pflegen.';
            tr.appendChild(td);
            tbody.appendChild(tr);
            return;
        }

        classes.forEach(function (cl, idx) {
            let nick = String(cl.stableMailNickname || '')
                .trim()
                .replace(/[^a-zA-Z0-9]/g, '')
                .toLowerCase()
                .slice(0, 60);
            if (!nick && cl.year && cl.code && typeof window.ms365DeriveClassStableMailNickname === 'function') {
                nick = String(window.ms365DeriveClassStableMailNickname(cl.year, cl.code) || '')
                    .trim()
                    .replace(/[^a-zA-Z0-9]/g, '')
                    .toLowerCase()
                    .slice(0, 60);
            }
            const ct = nick ? getClassTeamFromContainer(nick) : null;
            const smtpParts = buildClassSmtpPreview(nick);
            const rowSnap = {
                code: cl.code,
                year: cl.year,
                name: cl.name,
                headName: cl.headName,
                headEmail: cl.headEmail,
                stableMailNickname: cl.stableMailNickname
            };

            const tr = document.createElement('tr');

            const tdCode = document.createElement('td');
            tdCode.innerHTML = '<code>' + escapeHtml(cl.code || '') + '</code>';
            tdCode.title = 'Doppelklick zum Bearbeiten';
            tdCode.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdCode, cl.code, function (next, meta) {
                    const all = getSwClassesFromTextarea();
                    if (!all[idx]) return renderClassesTable();
                    const prev = all[idx].code;
                    all[idx].code = meta && meta.cancelled ? prev : normCode(next);
                    setSwClassesTextareaFromRows(all);
                    renderClassesTable();
                });
            });

            const tdYear = document.createElement('td');
            tdYear.textContent = cl.year || '';
            tdYear.title = 'Doppelklick zum Bearbeiten (YYYY)';
            tdYear.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdYear, cl.year || '', function (next, meta) {
                    const all = getSwClassesFromTextarea();
                    if (!all[idx]) return renderClassesTable();
                    const prev = all[idx].year || '';
                    const n = normStr(next);
                    all[idx].year = meta && meta.cancelled ? prev : /^\d{4}$/.test(n) ? n : '';
                    setSwClassesTextareaFromRows(all);
                    renderClassesTable();
                });
            });

            const tdName = document.createElement('td');
            tdName.textContent = cl.name || '';
            tdName.title = 'Doppelklick zum Bearbeiten';
            tdName.addEventListener('dblclick', function () {
                wizardStartCellEdit(tdName, cl.name || '', function (next, meta) {
                    const all = getSwClassesFromTextarea();
                    if (!all[idx]) return renderClassesTable();
                    const prev = all[idx].name;
                    all[idx].name = meta && meta.cancelled ? prev : normStr(next);
                    setSwClassesTextareaFromRows(all);
                    renderClassesTable();
                });
            });

            const tdPrev = document.createElement('td');
            tdPrev.className = 'sw-smtp-preview';
            tdPrev.title =
                ct && normStr(ct.graphGroupId)
                    ? 'Stabiler Mail‑Nickname der verknüpften Klassen‑Gruppe @ Domain'
                    : 'Vorschau aus Abschlussjahr und Kürzel (Schema jg…)';
            tdPrev.innerHTML = '<code style="font-size:0.88em;">' + escapeHtml(smtpParts.smtp) + '</code>';

            const tdM365 = document.createElement('td');
            tdM365.className = 'sw-catalog-grp';
            if (ct && normStr(ct.graphGroupId)) {
                tdM365.innerHTML = '<code style="font-size:0.82em;">' + escapeHtml(ct.graphGroupId) + '</code>';
            } else {
                tdM365.innerHTML = '<span style="color:var(--muted)">–</span>';
            }

            const tdAct = document.createElement('td');
            tdAct.className = 'action-cell';
            tdAct.style.whiteSpace = 'nowrap';
            const wrap = document.createElement('div');
            wrap.style.cssText =
                'display:inline-flex;flex-wrap:nowrap;gap:6px;align-items:center;justify-content:flex-end;width:100%;';
            const btnS = document.createElement('button');
            btnS.type = 'button';
            btnS.className = 'mini-btn';
            btnS.style.background = '#5e72e4';
            btnS.title = 'Bestehende Microsoft‑365‑Gruppe suchen und verknüpfen';
            btnS.setAttribute('aria-label', 'Suchen');
            btnS.innerHTML = '<i class="bi bi-search" aria-hidden="true"></i>';
            btnS.addEventListener('click', function () {
                runClassSearch(rowSnap, nick);
            });
            const btnC = document.createElement('button');
            btnC.type = 'button';
            btnC.className = 'mini-btn';
            btnC.style.background = '#2dce89';
            btnC.title = 'Neue Microsoft‑365‑Gruppe anlegen';
            btnC.setAttribute('aria-label', 'Anlegen');
            btnC.innerHTML = '<i class="bi bi-plus-circle" aria-hidden="true"></i>';
            btnC.addEventListener('click', function () {
                runClassCreate(rowSnap, nick);
            });
            const btnDel = document.createElement('button');
            btnDel.type = 'button';
            btnDel.className = 'mini-btn';
            btnDel.title = 'Zeile löschen';
            btnDel.setAttribute('aria-label', 'Zeile löschen');
            btnDel.innerHTML = '<i class="bi bi-x-lg" aria-hidden="true"></i>';
            btnDel.addEventListener('click', function () {
                const all = getSwClassesFromTextarea();
                all.splice(idx, 1);
                setSwClassesTextareaFromRows(all);
                renderClassesTable();
            });
            wrap.appendChild(btnS);
            wrap.appendChild(btnC);
            wrap.appendChild(btnDel);
            tdAct.appendChild(wrap);

            tr.appendChild(tdCode);
            tr.appendChild(tdYear);
            tr.appendChild(tdName);
            tr.appendChild(tdPrev);
            tr.appendChild(tdM365);
            tr.appendChild(tdAct);
            tbody.appendChild(tr);
        });
    }

    async function runClassSearch(cl, stableNick) {
        const defQ = normStr(cl.name) || normStr(cl.code) || stableNick || '';
        const q = await dlgPrompt('Microsoft‑365‑Gruppe suchen (Name, Mail, Alias):', defQ, {
            title: 'Klassen-Gruppe suchen',
            inputLabel: 'Suchbegriff (Name, Mail, Alias)'
        });
        if (q == null || !normStr(q)) return;
        const nick =
            stableNick ||
            (typeof window.ms365DeriveClassStableMailNickname === 'function'
                ? window.ms365DeriveClassStableMailNickname(cl.year || '', cl.code || '')
                : '');
        const nickNorm = String(nick || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase()
            .slice(0, 60);
        if (!nickNorm) {
            toast('Stabiler Alias fehlt: Kürzel und Abschlussjahr (YYYY) angeben.');
            return;
        }
        try {
            const token = await G().getGraphToken();
            const list = await G().searchUnifiedGroups(token, q);
            if (!list.length) {
                toast('Keine Gruppe gefunden.');
                return;
            }
            const pick = list[0];
            if (
                !(await dlgConfirm(
                    'Diese Gruppe mit dieser Klasse verknüpfen?\n\n' +
                        (pick.displayName || '') +
                        '\nAlias: ' +
                        (pick.mailNickname || '') +
                        '\n\nLokaler stabiler Schlüssel: ' +
                        nickNorm,
                    { title: 'Klassen-Gruppe verknüpfen', okText: 'Verknüpfen' }
                ))
            ) {
                return;
            }
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.upsertClassTeam === 'function') {
                window.ms365AppDataV2.upsertClassTeam({
                    stableMailNickname: nickNorm,
                    graphGroupId: pick.id,
                    displayName: pick.displayName || cl.name || '',
                    classCode: cl.code || '',
                    abschlussJahr: cl.year || '',
                    mode: 'matched'
                });
            }
            toast('Klassen‑Gruppe verknüpft.');
            renderClassesTable();
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function runClassCreate(cl, stableNick) {
        let nick =
            stableNick ||
            (typeof window.ms365DeriveClassStableMailNickname === 'function'
                ? window.ms365DeriveClassStableMailNickname(cl.year || '', cl.code || '')
                : '');
        nick = String(nick || '')
            .trim()
            .replace(/[^a-zA-Z0-9]/g, '')
            .toLowerCase()
            .slice(0, 60);
        if (!nick) {
            toast('Alias: Kürzel und Abschlussjahr (YYYY) sind nötig.');
            return;
        }
        const defDn = normStr(cl.name) || normStr(cl.code) || nick;
        const dn = await dlgPrompt('Anzeigename der Klassen‑Gruppe (Team):', defDn, {
            title: 'Klassen-Gruppe anlegen',
            inputLabel: 'Anzeigename'
        });
        if (dn == null || !normStr(dn)) return;
        try {
            let token = await G().getGraphToken();
            const g = await G().createUnifiedGroup(
                token,
                dn,
                nick,
                'Klassengruppe (stabiler Alias ' + nick + ') – MS365-Schulverwaltung'
            );
            await ensureOwnersDirektionOnly(token, g.id);
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.upsertClassTeam === 'function') {
                window.ms365AppDataV2.upsertClassTeam({
                    stableMailNickname: nick,
                    graphGroupId: g.id,
                    displayName: g.displayName || dn,
                    classCode: cl.code || '',
                    abschlussJahr: cl.year || '',
                    mode: 'created'
                });
            }
            toastGroupCreated('Klassen‑Gruppe', g, nick);
            renderClassesTable();
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function runCatalogCreate(kind, code, name) {
        const sliceKind = kind === 'arge' ? 'arge' : 'subject';
        if (sliceKind === 'subject') persistSubjectGroupPrefixFromDom();
        else persistArgeGroupPrefixFromDom();
        const defNick = deriveCatalogMailNickname(sliceKind, code);
        const defDn = (kind === 'arge' ? 'Arbeitsgruppe ' : 'Fach ') + (name || code);
        const dn = await dlgPrompt('Anzeigename der Microsoft‑365‑Gruppe:', defDn, {
            title: 'Gruppe anlegen',
            inputLabel: 'Anzeigename'
        });
        if (dn == null || !normStr(dn)) return;
        const nick = await dlgPrompt('Mail‑Nickname (ohne Domain):', defNick, {
            title: 'Gruppe anlegen',
            inputLabel: 'Mail-Nickname'
        });
        if (nick == null || !normStr(nick)) return;
        try {
            let token = await G().getGraphToken();
            const g = await G().createUnifiedGroup(
                token,
                dn,
                nick,
                (kind === 'arge' ? 'Arbeitsgruppe ' : 'Fach ') + (name || code) + schoolPhraseForWizardGroupDesc()
            );
            await ensureOwnersDirektionOnly(token, g.id);
            upsertCatalogLink({
                kind: kind,
                code: code,
                graphGroupId: g.id,
                displayName: g.displayName || dn,
                mailNickname: g.mailNickname || nick,
                mode: 'created'
            });
            toastGroupCreated('Gruppe', g, nick);
            fillCatalogSlice(kind === 'arge' ? 'arge' : 'subject');
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    /**
     * @param {'subject'|'arge'} kind
     * @param {string} code
     * @param {string} [name]
     */
    async function catalogProvisionUnifiedGroupNoPrompt(kind, code, name) {
        const sliceKind = kind === 'arge' ? 'arge' : 'subject';
        if (sliceKind === 'subject') persistSubjectGroupPrefixFromDom();
        else persistArgeGroupPrefixFromDom();
        const defNick = deriveCatalogMailNickname(sliceKind, code);
        const defDn = (kind === 'arge' ? 'Arbeitsgruppe ' : 'Fach ') + String(name || code || '').trim();
        const token = await G().getGraphToken();
        const g = await G().createUnifiedGroup(
            token,
            defDn,
            defNick,
            (kind === 'arge' ? 'Arbeitsgruppe ' : 'Fach ') + String(name || code || '') + schoolPhraseForWizardGroupDesc()
        );
        await ensureOwnersDirektionOnly(token, g.id);
        upsertCatalogLink({
            kind: kind,
            code: code,
            graphGroupId: g.id,
            displayName: g.displayName || defDn,
            mailNickname: g.mailNickname || defNick,
            mode: 'created'
        });
    }

    /**
     * Legt für alle Tabellenzeilen ohne graphGroupId jeweils eine M365‑Gruppe an (wie „Anlegen“, ohne Prompts).
     * @param {'subject'|'arge'} kind
     */
    async function runCatalogBulkCreateMissing(kind) {
        const graphKind = kind === 'arge' ? 'arge' : 'subject';
        const sliceKind = graphKind === 'arge' ? 'arge' : 'subject';
        const labelShort = graphKind === 'arge' ? 'ARGE' : 'Fach';
        if (sliceKind === 'subject') persistSubjectGroupPrefixFromDom();
        else persistArgeGroupPrefixFromDom();
        const rows = sliceKind === 'subject' ? getSwSubjectsFromTextarea() : getSwArgesFromTextarea();
        const pending = [];
        rows.forEach(function (row) {
            const code = normCode(row && row.code);
            if (!code) return;
            const link = getCatalogLink(graphKind, code);
            if (link && normStr(link.graphGroupId)) return;
            pending.push({ code: code, name: normStr(row && row.name) });
        });
        if (!pending.length) {
            toast('Alle ' + labelShort + '‑Einträge sind bereits mit einer Microsoft‑365‑Gruppe verknüpft.');
            return;
        }
        if (
            !(await dlgConfirm(
                String(pending.length) +
                    ' Microsoft‑365‑Gruppe(n) für ' +
                    (graphKind === 'arge' ? 'ARGE' : 'Fächer') +
                    ' anlegen?\n\n' +
                    'Es werden nur Zeilen ohne bestehende M365‑Verknüpfung (Object‑ID) bearbeitet.\n' +
                    'Anzeigename wie bei „Anlegen“ (Fach/Arbeitsgruppe …), Mail‑Alias aus Präfix + Kürzel – ohne weitere Rückfragen.\n' +
                    'Bei einem Fehler wird mit der nächsten Zeile fortgefahren.',
                { title: 'Sammel-Anlage', okText: 'Anlegen' }
            ))
        ) {
            return;
        }
        const btnId = graphKind === 'arge' ? 'swBtnArgesBulkCreateM365' : 'swBtnSubjectsBulkCreateM365';
        const btn = document.getElementById(btnId);
        let ok = 0;
        let fail = 0;
        try {
            if (btn) {
                btn.disabled = true;
                btn.setAttribute('aria-busy', 'true');
            }
            for (let i = 0; i < pending.length; i++) {
                const row = pending[i];
                try {
                    await catalogProvisionUnifiedGroupNoPrompt(graphKind, row.code, row.name);
                    ok++;
                } catch (e) {
                    fail++;
                    console.warn(e);
                }
                if (i < pending.length - 1 && typeof G().sleep === 'function') {
                    await G().sleep(450);
                }
            }
            fillCatalogSlice(sliceKind);
            const parts = [];
            if (ok) parts.push(ok + ' angelegt');
            if (fail) parts.push(fail + ' fehlgeschlagen');
            toast('Gruppen: ' + (parts.length ? parts.join(', ') : 'Keine Änderung.'));
        } catch (e) {
            toast('Sammel-Anlage: ' + (e.message || e));
        } finally {
            if (btn) {
                btn.disabled = false;
                btn.removeAttribute('aria-busy');
            }
        }
    }

    function bindSwGroupTabs(btnMatchId, btnNewId, panelMatchId, panelNewId) {
        const bm = document.getElementById(btnMatchId);
        const bn = document.getElementById(btnNewId);
        const pm = document.getElementById(panelMatchId);
        const pn = document.getElementById(panelNewId);
        if (!bm || !bn || !pm || !pn) return;
        function showMatch(isMatch) {
            bm.setAttribute('aria-selected', isMatch ? 'true' : 'false');
            bn.setAttribute('aria-selected', isMatch ? 'false' : 'true');
            pm.classList.toggle('active', isMatch);
            pn.classList.toggle('active', !isMatch);
        }
        bm.addEventListener('click', function () {
            showMatch(true);
        });
        bn.addEventListener('click', function () {
            showMatch(false);
        });
    }

    function wire() {
        document.querySelectorAll('[data-sw-step]').forEach(function (btn) {
            btn.addEventListener('click', function () {
                const n = parseInt(btn.getAttribute('data-sw-step'), 10);
                showStep(n);
            });
        });
        document.getElementById('swBtnLogin') &&
            document.getElementById('swBtnLogin').addEventListener('click', async function () {
                try {
                    await G().getGraphToken();
                    refreshSwAuthStatus();
                    toast('Anmeldung OK.');
                } catch (e) {
                    toast('Anmeldung: ' + (e.message || e));
                }
            });
        document.getElementById('swBtnSaveDomain') &&
            document.getElementById('swBtnSaveDomain').addEventListener('click', saveDomainStep);
        document.getElementById('swBtnSaveTeachers') &&
            document.getElementById('swBtnSaveTeachers').addEventListener('click', saveTeachersList);
        document.getElementById('swBtnSaveStudents') &&
            document.getElementById('swBtnSaveStudents').addEventListener('click', saveStudentsList);
        document.getElementById('swBtnSaveSga') &&
            document.getElementById('swBtnSaveSga').addEventListener('click', saveSgaList);
        document.getElementById('swBtnSaveStudentCouncil') &&
            document.getElementById('swBtnSaveStudentCouncil').addEventListener('click', saveStudentCouncilList);

        function bindGroupStep(forKind) {
            const suf = groupUiSuffix(forKind);
            wireSwMatchSummaryRefresh(forKind);
            document.getElementById('swBtnSearchGrp' + suf) &&
                document.getElementById('swBtnSearchGrp' + suf).addEventListener('click', async function () {
                    const inp = document.getElementById('swGroupSearch' + suf);
                    const q = inp ? inp.value.trim() : '';
                    if (!q) {
                        toast('Suchbegriff fehlt.');
                        return;
                    }
                    try {
                        const token = await G().getGraphToken();
                        const list = await G().searchUnifiedGroups(token, q);
                        renderSwSearchResults(list, forKind);
                    } catch (e) {
                        toast('Fehler: ' + (e.message || e));
                    }
                });
            document.getElementById('swBtnCreateGrp' + suf) &&
                document.getElementById('swBtnCreateGrp' + suf).addEventListener('click', async function () {
                    let c;
                    if (forKind === 'verwaltung') {
                        captureVerwaltungFormToCache();
                        c = swVerwaltungFormCache;
                    } else {
                        captureAllGroupForms();
                        c = swGroupFormCache[forKind];
                    }
                    if (!normStr(c.dn) || !normStr(c.nick)) {
                        toast('Anzeigename und Alias ausfüllen.');
                        return;
                    }
                    try {
                        let token = await G().getGraphToken();
                        const g = await G().createUnifiedGroup(token, c.dn, c.nick, c.desc);
                        if (forKind === 'verwaltung') await ensureOwnersForVerwaltungWizard(token, g.id);
                        else await ensureOwnersForSlgKind(token, g.id, forKind);
                        if (c.team) {
                            await G().provisionTeamForGroup(token, g.id);
                        }
                        setGroupIdForKind(forKind, String(g.id));
                        if (forKind === 'verwaltung') persistVerwaltungFull();
                        else persistMatched();
                        swMatchLiveSlot(forKind).fullyLoaded = false;
                        renderSwMatchSummaryForKind(forKind, g);
                        toastGroupCreated('Gruppe', g, c.nick);
                    } catch (e) {
                        toast('Fehler: ' + (e.message || e));
                    }
                });
            document.getElementById('swBtnSyncMembers' + suf) &&
                document.getElementById('swBtnSyncMembers' + suf).addEventListener('click', async function () {
                    const gid = getGroupIdForKind(forKind);
                    if (!gid) {
                        toast('Zuerst eine Gruppe verknüpfen oder anlegen.');
                        return;
                    }
                    const emails =
                        forKind === 'schueler'
                            ? swListCache.students
                            : forKind === 'verwaltung'
                              ? swListCache.adminEmails
                              : swListCache.teachers;
                    if (!emails.length) {
                        toast(
                            forKind === 'lehrer'
                                ? 'Keine E‑Mails in der Lehrerliste (oben speichern).'
                                : forKind === 'verwaltung'
                                  ? 'Keine E‑Mails in der Verwaltungsliste (oben speichern).'
                                  : 'Keine E‑Mails in der Schülerliste (oben speichern).'
                        );
                        return;
                    }
                    try {
                        const token = await G().getGraphToken();
                        const label =
                            forKind === 'schueler' ? 'Schüler' : forKind === 'verwaltung' ? 'Verwaltung' : 'Lehrer';
                        const logEl = document.getElementById('swSyncLog' + suf);
                        if (logEl) logEl.replaceChildren();
                        const prevK = swActiveKind;
                        swActiveKind = forKind;
                        await G().syncEmailsToGroup(token, gid, emails, label, appendSwLog);
                        swActiveKind = prevK;
                        if (forKind === 'verwaltung') await ensureOwnersForVerwaltungWizard(token, gid);
                        else await ensureOwnersForSlgKind(token, gid, forKind);
                        swMatchLiveSlot(forKind).fullyLoaded = false;
                        loadSwMatchLiveDetails(forKind, { force: true });
                        toast('Synchronisation beendet.');
                    } catch (e) {
                        toast('Fehler: ' + (e.message || e));
                    }
                });
        }
        bindSwGroupTabs(
            'swGrpTabBtnMatchLehrer',
            'swGrpTabBtnNewLehrer',
            'swGrpPanelMatchLehrer',
            'swGrpPanelNewLehrer'
        );
        bindSwGroupTabs(
            'swGrpTabBtnMatchSchueler',
            'swGrpTabBtnNewSchueler',
            'swGrpPanelMatchSchueler',
            'swGrpPanelNewSchueler'
        );
        bindSwGroupTabs(
            'swGrpTabBtnMatchVerwaltung',
            'swGrpTabBtnNewVerwaltung',
            'swGrpPanelMatchVerwaltung',
            'swGrpPanelNewVerwaltung'
        );
        bindGroupStep('lehrer');
        bindGroupStep('schueler');
        bindGroupStep('verwaltung');
        wireSwMatchSummaryRefresh('sga');
        wireSwMatchSummaryRefresh('studentCouncil');

        ['Lehrer', 'Schueler', 'Verwaltung'].forEach(function (suf) {
            const nick = document.getElementById('swNewNick' + suf);
            if (nick && !nick.dataset.swNickBound) {
                nick.dataset.swNickBound = '1';
                nick.addEventListener('input', function () {
                    refreshSwSmtpPreviewHint(suf);
                });
            }
        });
        ['swOwnerLehrer', 'swOwnerSchueler', 'swOwnerVerwaltung'].forEach(function (nm) {
            document.querySelectorAll('input[name="' + nm + '"]').forEach(function (inp) {
                inp.addEventListener('change', function () {
                    toggleSwOwnerManualFields();
                    if (nm === 'swOwnerVerwaltung') {
                        patchVerwaltungOwnerDraftFromDom();
                        refreshSwOwnerSummary('Verwaltung', 'verwaltung');
                    } else {
                        patchSlgOwnerDraftFromDom();
                        if (nm === 'swOwnerLehrer') refreshSwOwnerSummary('Lehrer', 'lehrer');
                        else refreshSwOwnerSummary('Schueler', 'schueler');
                    }
                });
            });
        });
        const swOwnMl = document.getElementById('swOwnerManualLehrer');
        const swOwnMs = document.getElementById('swOwnerManualSchueler');
        const swOwnMv = document.getElementById('swOwnerManualVerwaltung');
        if (swOwnMl && !swOwnMl.dataset.swOwnBound) {
            swOwnMl.dataset.swOwnBound = '1';
            swOwnMl.addEventListener('input', function () {
                patchSlgOwnerDraftFromDom();
                refreshSwOwnerSummary('Lehrer', 'lehrer');
            });
        }
        if (swOwnMs && !swOwnMs.dataset.swOwnBound) {
            swOwnMs.dataset.swOwnBound = '1';
            swOwnMs.addEventListener('input', function () {
                patchSlgOwnerDraftFromDom();
                refreshSwOwnerSummary('Schueler', 'schueler');
            });
        }
        if (swOwnMv && !swOwnMv.dataset.swOwnBound) {
            swOwnMv.dataset.swOwnBound = '1';
            swOwnMv.addEventListener('input', function () {
                patchVerwaltungOwnerDraftFromDom();
                refreshSwOwnerSummary('Verwaltung', 'verwaltung');
            });
        }

        document.getElementById('swBtnSaveAdmin') &&
            document.getElementById('swBtnSaveAdmin').addEventListener('click', saveSwAdminList);

        document.getElementById('swBtnAddAdminRoleRow') &&
            document.getElementById('swBtnAddAdminRoleRow').addEventListener('click', function () {
                const tbody = document.getElementById('swAdminTableBody');
                if (!tbody) return;
                appendSwAdminCustomRowToWizardTable(tbody, {});
                refreshSwStatsVerwaltung();
                refreshSwOwnerSummary('Verwaltung', 'verwaltung');
            });

        document.getElementById('swBtnVerifyVerwaltungGraph') &&
            document.getElementById('swBtnVerifyVerwaltungGraph').addEventListener('click', function () {
                runVerifyVerwaltungGraph();
            });
        document.getElementById('swBtnAdminRestoreDefaults') &&
            document.getElementById('swBtnAdminRestoreDefaults').addEventListener('click', function () {
                void (async function () {
                    if (
                        await dlgConfirm(
                            'Die üblichen acht Standardrollen als Tabelle einfügen und lokal speichern? Bestehende Verwaltungszeilen werden dabei ersetzt.',
                            { title: 'Standardrollen', okText: 'Ersetzen' }
                        )
                    ) {
                        seedDefaultAdminRolesToSettings();
                    }
                })();
            });

        document.getElementById('swBtnTeachersTplXlsx') &&
            document.getElementById('swBtnTeachersTplXlsx').addEventListener('click', function () {
                const api = window.ms365TeacherListImport;
                if (!api || typeof api.downloadTemplate !== 'function') {
                    toast('Vorlagen-Modul fehlt (tenant-settings-ui.js).');
                    return;
                }
                if (typeof api.isXlsxReady === 'function' && !api.isXlsxReady()) {
                    toast('Excel-Bibliothek noch nicht geladen – Seite kurz warten und erneut versuchen.');
                    return;
                }
                if (!api.downloadTemplate()) {
                    toast('Vorlage konnte nicht erzeugt werden.');
                }
            });
        document.getElementById('swBtnTeachersTplCsv') &&
            document.getElementById('swBtnTeachersTplCsv').addEventListener('click', function () {
                const api = window.ms365TeacherListImport;
                if (!api || typeof api.downloadCsvTemplate !== 'function') {
                    toast('Vorlagen-Modul fehlt (tenant-settings-ui.js).');
                    return;
                }
                if (!api.downloadCsvTemplate()) {
                    toast('CSV-Vorlage konnte nicht erzeugt werden.');
                }
            });
        const swTeachersImportFile = document.getElementById('swTeachersImportFile');
        if (swTeachersImportFile) {
            swTeachersImportFile.addEventListener('change', function (e) {
                const api = window.ms365TeacherListImport;
                const f = e.target.files && e.target.files[0];
                if (!f || !api || typeof api.importFile !== 'function') return;
                api.importFile(
                    f,
                    function (lines) {
                        const ta = document.getElementById('swTeachersLines');
                        if (ta) {
                            const cur = normStr(ta.value);
                            ta.value = cur ? cur + '\n' + lines : lines;
                        }
                        renderSwTeachersTableFromTextarea();
                        toast('Import in die Lehrerliste übernommen (noch nicht gespeichert).');
                    },
                    function (err) {
                        toast(err || 'Import fehlgeschlagen.');
                    }
                );
                swTeachersImportFile.value = '';
            });
        }
        const swTaTeachers = document.getElementById('swTeachersLines');
        if (swTaTeachers) {
            swTaTeachers.addEventListener('input', function () {
                renderSwTeachersTableFromTextarea();
            });
        }
        document.getElementById('swBtnTeachersAddRow') &&
            document.getElementById('swBtnTeachersAddRow').addEventListener('click', function () {
                const all = getSwTeachersFromTextarea();
                all.push({ code: 'XXX', name: '', email: '' });
                setSwTeachersTextareaFromRows(all);
                renderSwTeachersTableFromTextarea();
            });
        document.getElementById('swBtnStudentsTplXlsx') &&
            document.getElementById('swBtnStudentsTplXlsx').addEventListener('click', function () {
                const api = window.ms365StudentListImport;
                if (!api || typeof api.downloadTemplate !== 'function') {
                    toast('Vorlagen-Modul fehlt (tenant-settings-ui.js).');
                    return;
                }
                if (typeof api.isXlsxReady === 'function' && !api.isXlsxReady()) {
                    toast('Excel-Bibliothek noch nicht geladen – Seite kurz warten und erneut versuchen.');
                    return;
                }
                if (!api.downloadTemplate()) {
                    toast('Vorlage konnte nicht erzeugt werden.');
                }
            });
        document.getElementById('swBtnStudentsTplCsv') &&
            document.getElementById('swBtnStudentsTplCsv').addEventListener('click', function () {
                const api = window.ms365StudentListImport;
                if (!api || typeof api.downloadCsvTemplate !== 'function') {
                    toast('Vorlagen-Modul fehlt (tenant-settings-ui.js).');
                    return;
                }
                if (!api.downloadCsvTemplate()) {
                    toast('CSV-Vorlage konnte nicht erzeugt werden.');
                }
            });
        wireSchuldatenMasterDownloadClick('swBtnSchuldatenMasterTpl4');
        wireSchuldatenMasterDownloadClick('swBtnSchuldatenMasterTpl5');
        const swStudentsImportFile = document.getElementById('swStudentsImportFile');
        if (swStudentsImportFile) {
            swStudentsImportFile.addEventListener('change', function (e) {
                const api = window.ms365StudentListImport;
                const f = e.target.files && e.target.files[0];
                if (!f || !api || typeof api.importFile !== 'function') return;
                api.importFile(
                    f,
                    function (lines) {
                        const ta = document.getElementById('swStudentsLines');
                        if (ta) {
                            const cur = normStr(ta.value);
                            ta.value = cur ? cur + '\n' + lines : lines;
                        }
                        renderSwStudentsTableFromTextarea();
                        toast('Import in die Schülerliste übernommen (noch nicht in „Schülerliste speichern“ geschrieben).');
                    },
                    function (err) {
                        toast(err || 'Import fehlgeschlagen.');
                    }
                );
                swStudentsImportFile.value = '';
            });
        }
        const swSchuldatenMasterImportFile = document.getElementById('swSchuldatenMasterImportFile');
        if (swSchuldatenMasterImportFile) {
            swSchuldatenMasterImportFile.addEventListener('change', function (e) {
                const api = window.ms365SchuldatenMasterImport;
                const f = e.target.files && e.target.files[0];
                if (!f || !api || typeof api.importFile !== 'function') return;
                api.importFile(
                    f,
                    function (payload) {
                        applySchuldatenMasterImportPayload(payload);
                    },
                    function (err) {
                        toast(err || 'Gesamt-Import fehlgeschlagen.');
                    }
                );
                swSchuldatenMasterImportFile.value = '';
            });
        }
        const swTaStudents = document.getElementById('swStudentsLines');
        if (swTaStudents) {
            swTaStudents.addEventListener('input', function () {
                renderSwStudentsTableFromTextarea();
            });
        }
        const swTaSga = document.getElementById('swSgaLines');
        if (swTaSga) {
            swTaSga.addEventListener('input', function () {
                renderSwSgaTableFromTextarea();
            });
        }
        const swTaStudentCouncil = document.getElementById('swStudentCouncilLines');
        if (swTaStudentCouncil) {
            swTaStudentCouncil.addEventListener('input', function () {
                renderSwStudentCouncilTableFromTextarea();
            });
        }
        document.getElementById('swBtnStudentsAddRow') &&
            document.getElementById('swBtnStudentsAddRow').addEventListener('click', function () {
                const all = getSwStudentsFromTextarea();
                all.push({ klasse: '1x', name: '', email: '' });
                setSwStudentsTextareaFromRows(all);
                renderSwStudentsTableFromTextarea();
            });
        document.getElementById('swBtnSgaAddRow') &&
            document.getElementById('swBtnSgaAddRow').addEventListener('click', function () {
                const all = getSwSgaFromTextarea();
                all.push({ scope: '', name: '', email: '' });
                setSwSgaTextareaFromRows(all);
                renderSwSgaTableFromTextarea();
            });
        document.getElementById('swBtnStudentCouncilAddRow') &&
            document.getElementById('swBtnStudentCouncilAddRow').addEventListener('click', function () {
                const all = getSwStudentCouncilFromTextarea();
                all.push({ klasse: '', name: '', email: '' });
                setSwStudentCouncilTextareaFromRows(all);
                renderSwStudentCouncilTableFromTextarea();
            });
        document.getElementById('swBtnVerifyTeachersGraph') &&
            document.getElementById('swBtnVerifyTeachersGraph').addEventListener('click', function () {
                runVerifyTeachersGraph();
            });
        document.getElementById('swBtnCreateMissingTeachers') &&
            document.getElementById('swBtnCreateMissingTeachers').addEventListener('click', function () {
                runCreateMissingTeachersEntra();
            });
        document.getElementById('swBtnVerifyStudentsGraph') &&
            document.getElementById('swBtnVerifyStudentsGraph').addEventListener('click', function () {
                runVerifyStudentsGraph();
            });
        document.getElementById('swBtnVerifySgaGraph') &&
            document.getElementById('swBtnVerifySgaGraph').addEventListener('click', function () {
                runVerifySgaGraph();
            });
        document.getElementById('swBtnVerifyStudentCouncilGraph') &&
            document.getElementById('swBtnVerifyStudentCouncilGraph').addEventListener('click', function () {
                runVerifyStudentCouncilGraph();
            });
        document.getElementById('swBtnCreateMissingStudents') &&
            document.getElementById('swBtnCreateMissingStudents').addEventListener('click', function () {
                runCreateMissingStudentsEntra();
            });
        document.getElementById('swBtnCreateMissingSga') &&
            document.getElementById('swBtnCreateMissingSga').addEventListener('click', function () {
                runCreateMissingSgaEntra();
            });
        document.getElementById('swBtnCreateMissingStudentCouncil') &&
            document.getElementById('swBtnCreateMissingStudentCouncil').addEventListener('click', function () {
                runCreateMissingStudentCouncilEntra();
            });
        document.getElementById('swBtnVerifySgaGroup') &&
            document.getElementById('swBtnVerifySgaGroup').addEventListener('click', async function () {
                try {
                    await verifyWizardSchoolGroup('sga');
                } catch (e) {
                    toast('SGA-Gruppe prüfen: ' + (e.message || e));
                }
            });
        document.getElementById('swBtnCreateSgaGroup') &&
            document.getElementById('swBtnCreateSgaGroup').addEventListener('click', async function () {
                try {
                    await createWizardSchoolGroup('sga');
                } catch (e) {
                    toast('SGA-Gruppe anlegen: ' + (e.message || e));
                }
            });
        document.getElementById('swBtnVerifyStudentCouncilGroup') &&
            document.getElementById('swBtnVerifyStudentCouncilGroup').addEventListener('click', async function () {
                try {
                    await verifyWizardSchoolGroup('studentCouncil');
                } catch (e) {
                    toast('Schülervertretung prüfen: ' + (e.message || e));
                }
            });
        document.getElementById('swBtnCreateStudentCouncilGroup') &&
            document.getElementById('swBtnCreateStudentCouncilGroup').addEventListener('click', async function () {
                try {
                    await createWizardSchoolGroup('studentCouncil');
                } catch (e) {
                    toast('Schülervertretung anlegen: ' + (e.message || e));
                }
            });
        document.getElementById('swBtnSyncMembersSga') &&
            document.getElementById('swBtnSyncMembersSga').addEventListener('click', async function () {
                const gid = getGroupIdForKind('sga');
                const emails = collectEmails(getSwSgaFromTextarea());
                if (!gid) return toast('Zuerst die SGA-Gruppe prüfen oder anlegen.');
                if (!emails.length) return toast('Keine E‑Mails in der SGA-Liste (oben speichern).');
                try {
                    const token = await G().getGraphToken();
                    const logEl = document.getElementById('swSyncLogSga');
                    if (logEl) logEl.replaceChildren();
                    await G().syncEmailsToGroup(token, gid, emails, 'SGA', function (msg, kind) {
                        appendSwLogToSuffix('Sga', msg, kind);
                    });
                    await ensureOwnersForVerwaltungWizard(token, gid);
                    swMatchLiveSlot('sga').fullyLoaded = false;
                    loadSwMatchLiveDetails('sga', { force: true });
                    toast('SGA synchronisiert.');
                } catch (e) {
                    toast('SGA synchronisieren: ' + (e.message || e));
                }
            });
        document.getElementById('swBtnSyncMembersStudentCouncil') &&
            document.getElementById('swBtnSyncMembersStudentCouncil').addEventListener('click', async function () {
                const gid = getGroupIdForKind('studentCouncil');
                const emails = collectEmails(getSwStudentCouncilFromTextarea());
                if (!gid) return toast('Zuerst die Schülervertretungs-Gruppe prüfen oder anlegen.');
                if (!emails.length) return toast('Keine E‑Mails in der Schülervertretungsliste (oben speichern).');
                try {
                    const token = await G().getGraphToken();
                    const logEl = document.getElementById('swSyncLogStudentCouncil');
                    if (logEl) logEl.replaceChildren();
                    await G().syncEmailsToGroup(token, gid, emails, 'Schülervertretung', function (msg, kind) {
                        appendSwLogToSuffix('StudentCouncil', msg, kind);
                    });
                    await ensureOwnersForSlgKind(token, gid, 'schueler');
                    swMatchLiveSlot('studentCouncil').fullyLoaded = false;
                    loadSwMatchLiveDetails('studentCouncil', { force: true });
                    toast('Schülervertretung synchronisiert.');
                } catch (e) {
                    toast('Schülervertretung synchronisieren: ' + (e.message || e));
                }
            });
        const swSgaMode = document.getElementById('swSgaMode');
        if (swSgaMode) {
            swSgaMode.addEventListener('change', function () {
                setGroupIdForKind('sga', null);
                persistMatched();
                setWizardGroupStatus('sga', null);
            });
        }
        document.getElementById('swBtnSaveSubjectsBulk') &&
            document.getElementById('swBtnSaveSubjectsBulk').addEventListener('click', saveSubjectsBulk);
        const swTaSubjects = document.getElementById('swSubjectsBulk');
        if (swTaSubjects) {
            swTaSubjects.addEventListener('input', function () {
                fillCatalogSlice('subject');
            });
        }
        const swPrefSub = document.getElementById('swSubjectGroupPrefix');
        if (swPrefSub) {
            swPrefSub.addEventListener('input', function () {
                persistSubjectGroupPrefixFromDom();
                fillCatalogSlice('subject');
            });
        }
        const swPrefArge = document.getElementById('swArgeGroupPrefix');
        if (swPrefArge) {
            swPrefArge.addEventListener('input', function () {
                persistArgeGroupPrefixFromDom();
                fillCatalogSlice('arge');
            });
        }
        const swTaArges = document.getElementById('swArgesBulk');
        if (swTaArges) {
            swTaArges.addEventListener('input', function () {
                fillCatalogSlice('arge');
            });
        }
        document.getElementById('swBtnArgesAddRow') &&
            document.getElementById('swBtnArgesAddRow').addEventListener('click', function () {
                const all = getSwArgesFromTextarea();
                all.push({ code: 'AG', name: '', subjects: [] });
                setSwArgesTextareaFromRows(all);
                fillCatalogSlice('arge');
            });
        document.getElementById('swBtnArgesAutoMatch') &&
            document.getElementById('swBtnArgesAutoMatch').addEventListener('click', function () {
                runArgesAutomatch();
            });
        document.getElementById('swBtnArgesBulkCreateM365') &&
            document.getElementById('swBtnArgesBulkCreateM365').addEventListener('click', function () {
                runCatalogBulkCreateMissing('arge');
            });
        document.getElementById('swBtnSubjectsAddRow') &&
            document.getElementById('swBtnSubjectsAddRow').addEventListener('click', function () {
                const all = getSwSubjectsFromTextarea();
                all.push({ code: 'XX', name: '' });
                setSwSubjectsTextareaFromRows(all);
                fillCatalogSlice('subject');
            });
        document.getElementById('swBtnSubjectsAutoMatch') &&
            document.getElementById('swBtnSubjectsAutoMatch').addEventListener('click', function () {
                runSubjectsAutomatch();
            });
        document.getElementById('swBtnSubjectsBulkCreateM365') &&
            document.getElementById('swBtnSubjectsBulkCreateM365').addEventListener('click', function () {
                runCatalogBulkCreateMissing('subject');
            });
        document.getElementById('swBtnSaveArgesBulk') &&
            document.getElementById('swBtnSaveArgesBulk').addEventListener('click', saveArgesBulk);
        document.getElementById('swBtnSkipSubjects') &&
            document.getElementById('swBtnSkipSubjects').addEventListener('click', function () {
                showStep(9);
            });
        document.getElementById('swBtnSkipArges') &&
            document.getElementById('swBtnSkipArges').addEventListener('click', function () {
                showStep(10);
            });
        document.getElementById('swBtnSaveClassesBulk') &&
            document.getElementById('swBtnSaveClassesBulk').addEventListener('click', saveClassesBulk);
        const swTaClasses = document.getElementById('swClassesBulk');
        if (swTaClasses) {
            swTaClasses.addEventListener('input', function () {
                renderClassesTable();
            });
        }
        document.getElementById('swBtnClassesAddRow') &&
            document.getElementById('swBtnClassesAddRow').addEventListener('click', function () {
                const all = getSwClassesFromTextarea();
                all.push({ code: '1A', year: '', name: '', headName: '', headEmail: '', stableMailNickname: '' });
                setSwClassesTextareaFromRows(all);
                renderClassesTable();
            });
        document.getElementById('swBtnStep10Next') &&
            document.getElementById('swBtnStep10Next').addEventListener('click', function () {
                showStep(11);
            });
        document.getElementById('swBtnFinish') &&
            document.getElementById('swBtnFinish').addEventListener('click', function () {
                try {
                    if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                        const steps = [
                            'auth',
                            'domain',
                            'verwaltung',
                            'teachers',
                            'students',
                            'sga',
                            'studentCouncil',
                            'subjects',
                            'arges',
                            'classes'
                        ];
                        window.ms365AppDataV2.patchSetup({
                            completedSteps: steps,
                            finishedAt: new Date().toISOString()
                        });
                    }
                } catch {
                    // ignore
                }
                toast('Einrichtung abgeschlossen (lokal gespeichert).');
            });
    }

    function init() {
        syncSetupFromAppData();
        wire();
        const hash = (window.location.hash || '').replace(/^#/, '');
        let start = 1;
        if (hash === 'step2') start = 2;
        else if (hash === 'step3') start = 3;
        else if (hash === 'step4') start = 4;
        else if (hash === 'step5') start = 5;
        else if (hash === 'step6') start = 6;
        else if (hash === 'step7') start = 7;
        else if (hash === 'step8') start = 8;
        else if (hash === 'step9') start = 9;
        else if (hash === 'step10') start = 10;
        else if (hash === 'step11') start = 11;
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const ws = window.ms365AppDataV2.getSetup().wizardStep;
                if (ws >= 1 && ws <= 11) start = ws;
            }
        } catch {
            // ignore
        }
        if (!swGroupFormCacheBootstrapped) {
            initSwGroupFormCacheFromSlgDraft();
            initSwVerwaltungFormCacheFromDraft();
            swGroupFormCacheBootstrapped = true;
        }
        fillCatalogTextareas();
        readGroupPrefixesFromSetupToDom();
        refreshSwAuthStatus();
        window.addEventListener('ms365-auth-widget-ready', refreshSwAuthStatus);
        showStep(start);
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
