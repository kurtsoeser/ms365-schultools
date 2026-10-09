(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-lizenzverwaltung-v1';
    const ENTRA_GROUP =
        'https://entra.microsoft.com/#view/Microsoft_AAD_IAM/GroupDetailsMenuBlade/~/Members/groupId/';

    const SLOT_DEFS = {
        schueler: {
            label: 'Schüler:innen',
            defaultName: 'Lizenz Schüler',
            defaultNick: 'lizenz-schueler',
            audience: 'student',
            sammelKey: 'schuelerGroupId',
            deptHint: 'Schüler'
        },
        lehrer: {
            label: 'Lehrkräfte',
            defaultName: 'Lizenz Lehrkräfte',
            defaultNick: 'lizenz-lehrkraefte',
            audience: 'faculty',
            sammelKey: 'lehrerGroupId',
            deptHint: 'Lehrer'
        },
        schulleitung: {
            label: 'Schulleitung',
            defaultName: 'Lizenz Schulleitung',
            defaultNick: 'lizenz-schulleitung',
            audience: 'other',
            sammelKey: 'schulleitungGroupId',
            deptHint: 'Schulleitung'
        },
        verwaltung: {
            label: 'Verwaltung (Personal)',
            defaultName: 'Lizenz Verwaltung Personal',
            defaultNick: 'lizenz-verwaltung',
            audience: 'other',
            sammelKey: 'verwaltungGroupId',
            deptHint: 'Verwaltung'
        }
    };

    /** @type {string} */
    let activeSlot = 'schueler';
    /** @type {Record<string, any>} */
    let state = loadState();
    /** @type {any[]} */
    let subscribedSkus = [];
    /** @type {boolean} */
    let skusOk = false;
    /** @type {any|null} */
    let currentGroup = null;

    function $(id) {
        return document.getElementById(id);
    }

    function gug() {
        const api = window.ms365GraphUnifiedGroups;
        if (!api) throw new Error('Graph-Modul fehlt – Seite neu laden.');
        return api;
    }

    function licApi() {
        return window.ms365GraphLicenses || null;
    }

    function toast(msg) {
        const el = $('toast');
        if (!el) return;
        el.textContent = String(msg || '');
        el.classList.add('show');
        clearTimeout(toast._t);
        toast._t = setTimeout(function () {
            el.classList.remove('show');
        }, 3200);
    }

    function log(msg, kind) {
        const el = $('lvLog');
        if (!el) return;
        const line = document.createElement('div');
        const prefix = kind === 'ok' ? '✓ ' : kind === 'warn' ? '! ' : kind === 'err' ? '✗ ' : '';
        line.textContent = new Date().toLocaleTimeString() + '  ' + prefix + String(msg || '');
        if (kind === 'err') line.style.color = '#b00020';
        else if (kind === 'ok') line.style.color = '#0d8050';
        else if (kind === 'warn') line.style.color = '#856404';
        el.appendChild(line);
        el.scrollTop = el.scrollHeight;
    }

    function escapeHtml(s) {
        return String(s || '')
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function normStr(s) {
        return String(s == null ? '' : s).trim();
    }

    function loadState() {
        try {
            const raw = localStorage.getItem(STORAGE_KEY);
            const obj = raw ? JSON.parse(raw) : null;
            const slots = obj && obj.slots && typeof obj.slots === 'object' ? obj.slots : {};
            const custom = Array.isArray(obj && obj.custom) ? obj.custom : [];
            return { slots: slots, custom: custom };
        } catch {
            return { slots: {}, custom: [] };
        }
    }

    function saveState() {
        try {
            localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
        } catch {
            /* ignore */
        }
    }

    function slotDef(id) {
        if (SLOT_DEFS[id]) return SLOT_DEFS[id];
        const c = (state.custom || []).find(function (x) {
            return x && x.id === id;
        });
        if (c) {
            return {
                label: c.label || 'Eigenes Profil',
                defaultName: c.label || 'Lizenz Gruppe',
                defaultNick: 'lizenz-' + String(c.id || 'custom').slice(0, 40),
                audience: 'other',
                sammelKey: '',
                deptHint: ''
            };
        }
        return SLOT_DEFS.schueler;
    }

    function getSlotGroupId(id) {
        if (SLOT_DEFS[id]) {
            const s = state.slots[id] || {};
            return normStr(s.groupId);
        }
        const c = (state.custom || []).find(function (x) {
            return x && x.id === id;
        });
        return c ? normStr(c.groupId) : '';
    }

    function setSlotGroupId(id, groupId) {
        const gid = normStr(groupId) || null;
        if (SLOT_DEFS[id]) {
            state.slots[id] = Object.assign({}, state.slots[id] || {}, { groupId: gid });
        } else {
            state.custom = (state.custom || []).map(function (c) {
                if (!c || c.id !== id) return c;
                return Object.assign({}, c, { groupId: gid });
            });
        }
        saveState();
    }

    function getMatchedSammelId(key) {
        if (!key) return '';
        try {
            const api = window.ms365AppDataV2;
            const setup = api && typeof api.getSetup === 'function' ? api.getSetup() : null;
            const matched = setup && setup.matched && typeof setup.matched === 'object' ? setup.matched : {};
            return normStr(matched[key] || '');
        } catch {
            return '';
        }
    }

    function renderSlotList() {
        const list = $('lvSlotList');
        if (!list) return;
        const builtin = ['schueler', 'lehrer', 'schulleitung', 'verwaltung'];
        list.replaceChildren();
        builtin.forEach(function (id) {
            list.appendChild(makeSlotButton(id, slotDef(id).label, getSlotGroupId(id)));
        });
        (state.custom || []).forEach(function (c) {
            if (!c || !c.id) return;
            list.appendChild(makeSlotButton(c.id, c.label || 'Eigenes Profil', normStr(c.groupId)));
        });
        updateSlotMetas();
    }

    function makeSlotButton(id, label, groupId) {
        const li = document.createElement('li');
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'lv-slot-btn';
        btn.setAttribute('data-lv-slot', id);
        btn.setAttribute('aria-current', id === activeSlot ? 'true' : 'false');
        btn.innerHTML =
            '<span class="lv-slot-title">' +
            escapeHtml(label) +
            '</span><span class="lv-slot-meta" data-lv-meta="' +
            escapeHtml(id) +
            '">' +
            (groupId ? 'Gruppe zugeordnet' : 'Noch keine Lizenzgruppe') +
            '</span>';
        btn.addEventListener('click', function () {
            selectSlot(id);
        });
        li.appendChild(btn);
        return li;
    }

    function updateSlotMetas() {
        document.querySelectorAll('[data-lv-meta]').forEach(function (el) {
            const id = el.getAttribute('data-lv-meta');
            const gid = getSlotGroupId(id);
            el.textContent = gid
                ? currentGroup && String(currentGroup.id) === gid
                    ? normStr(currentGroup.displayName) || 'Gruppe zugeordnet'
                    : 'Gruppe zugeordnet'
                : 'Noch keine Lizenzgruppe';
        });
        const map = {
            schueler: 'lvMetaSchueler',
            lehrer: 'lvMetaLehrer',
            schulleitung: 'lvMetaSchulleitung',
            verwaltung: 'lvMetaVerwaltung'
        };
        Object.keys(map).forEach(function (k) {
            const el = $(map[k]);
            if (!el) return;
            const gid = getSlotGroupId(k);
            el.textContent = gid ? 'Gruppe zugeordnet' : 'Noch keine Lizenzgruppe';
        });
    }

    async function selectSlot(id) {
        activeSlot = id;
        document.querySelectorAll('.lv-slot-btn').forEach(function (b) {
            b.setAttribute('aria-current', b.getAttribute('data-lv-slot') === id ? 'true' : 'false');
        });
        const def = slotDef(id);
        const title = $('lvDetailTitle');
        if (title) {
            title.innerHTML = '<i class="bi bi-key" style="margin-right:8px;"></i>' + escapeHtml(def.label);
        }
        const hint = $('lvGroupHint');
        if (hint) {
            hint.textContent =
                'Reine Sicherheitsgruppe (ohne E‑Mail) für „' +
                def.label +
                '“. Empfohlener Name: ' +
                def.defaultName +
                '.';
        }
        currentGroup = null;
        await refreshDetail();
    }

    function setGroupUi(group) {
        currentGroup = group && group.id ? group : null;
        const st = $('lvGroupStatus');
        const btnUn = $('lvBtnUnmatch');
        const btnEntra = $('lvBtnOpenEntra');
        const licSec = $('lvLicenseSection');
        const memSec = $('lvMembersSection');
        const dynSec = $('lvDynamicSection');
        if (!currentGroup) {
            if (st) {
                st.className = 'lv-status lv-status--warn';
                st.innerHTML = '<i class="bi bi-x-circle"></i> nicht zugeordnet';
            }
            if (btnUn) btnUn.hidden = true;
            if (btnEntra) btnEntra.hidden = true;
            if (licSec) licSec.hidden = true;
            if (memSec) memSec.hidden = true;
            if (dynSec) dynSec.hidden = true;
            updateSlotMetas();
            return;
        }
        if (st) {
            st.className = 'lv-status lv-status--ok';
            st.innerHTML =
                '<i class="bi bi-check-circle"></i> ' +
                escapeHtml(normStr(currentGroup.displayName) || 'zugeordnet');
            st.title =
                (currentGroup.mailNickname ? 'Alias: ' + currentGroup.mailNickname + '\n' : '') +
                'ID: ' +
                currentGroup.id;
        }
        if (btnUn) btnUn.hidden = false;
        if (btnEntra) btnEntra.hidden = false;
        if (licSec) licSec.hidden = false;
        if (memSec) memSec.hidden = false;
        if (dynSec) dynSec.hidden = false;
        updateSlotMetas();
    }

    async function refreshDetail() {
        const gid = getSlotGroupId(activeSlot);
        setGroupUi(null);
        if (!gid) {
            renderAssignedSkus([]);
            renderMembers([]);
            return;
        }
        try {
            const token = await gug().getGraphToken();
            await ensureSkus(token);
            const g = await fetchGroupForLicensing(token, gid);
            setGroupUi(g);
            renderAssignedSkus((g && g.assignedLicenses) || []);
            fillSkuSelect((g && g.assignedLicenses) || []);
            const rule = $('lvDynamicRule');
            if (rule) rule.value = normStr(g && g.membershipRule);
            await loadMembers(token, gid);
        } catch (e) {
            log('Gruppe laden: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Gruppe konnte nicht geladen werden.');
        }
    }

    async function fetchGroupForLicensing(token, id) {
        const path =
            '/groups/' +
            encodeURIComponent(id) +
            '?$select=' +
            encodeURIComponent(
                'id,displayName,mail,mailNickname,description,securityEnabled,mailEnabled,groupTypes,assignedLicenses,membershipRule,membershipRuleProcessingState'
            );
        return gug().graphJson('GET', path, token);
    }

    async function ensureSkus(token) {
        if (skusOk && subscribedSkus.length) return;
        try {
            const data = await gug().graphJson(
                'GET',
                '/subscribedSkus?$select=' +
                    encodeURIComponent('skuId,skuPartNumber,prepaidUnits,consumedUnits,capabilityStatus'),
                token
            );
            subscribedSkus = Array.isArray(data.value) ? data.value : [];
            skusOk = true;
        } catch (e) {
            subscribedSkus = [];
            skusOk = false;
            log('SKUs nicht lesbar: ' + (e && e.message ? e.message : String(e)), 'warn');
        }
    }

    function skuLabel(skuId, part) {
        const api = licApi();
        if (api && typeof api.resolveSku === 'function') {
            const info = api.resolveSku(skuId, part);
            if (info) return info.shortLabel || info.name || part || skuId;
        }
        return part || skuId;
    }

    function renderAssignedSkus(list) {
        const host = $('lvAssignedSkus');
        if (!host) return;
        host.replaceChildren();
        const arr = Array.isArray(list) ? list : [];
        if (!arr.length) {
            host.innerHTML = '<span class="muted">Noch keine Lizenz an dieser Gruppe.</span>';
            return;
        }
        arr.forEach(function (lic) {
            const skuId = String((lic && lic.skuId) || '').toLowerCase();
            const part = findPartNumber(skuId);
            const chip = document.createElement('span');
            chip.className = 'lv-sku-chip';
            chip.innerHTML =
                '<strong>' +
                escapeHtml(skuLabel(skuId, part)) +
                '</strong> <code style="font-size:0.85em;">' +
                escapeHtml(part || skuId.slice(0, 8)) +
                '</code>';
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-danger small-btn';
            btn.innerHTML = '<i class="bi bi-dash-circle"></i>Entziehen';
            btn.addEventListener('click', function () {
                void removeSku(skuId);
            });
            chip.appendChild(btn);
            host.appendChild(chip);
        });
    }

    function findPartNumber(skuId) {
        const hit = subscribedSkus.find(function (s) {
            return String((s && s.skuId) || '').toLowerCase() === String(skuId || '').toLowerCase();
        });
        return hit ? String(hit.skuPartNumber || '') : '';
    }

    function fillSkuSelect(assigned) {
        const sel = $('lvSkuSelect');
        if (!sel) return;
        const assignedIds = (Array.isArray(assigned) ? assigned : []).map(function (l) {
            return String((l && l.skuId) || '').toLowerCase();
        });
        const api = licApi();
        let opts = [];
        if (api && typeof api.buildAssignableSkuOptions === 'function') {
            opts = api.buildAssignableSkuOptions(subscribedSkus, assignedIds, {
                fallbackCatalog: !skusOk
            });
        } else {
            opts = subscribedSkus
                .filter(function (s) {
                    const id = String((s && s.skuId) || '').toLowerCase();
                    return id && assignedIds.indexOf(id) === -1;
                })
                .map(function (s) {
                    return {
                        skuId: String(s.skuId).toLowerCase(),
                        name: s.skuPartNumber || s.skuId,
                        remaining: null,
                        disabled: false
                    };
                });
        }
        // Audience-Hinweis: passendere SKUs nach oben (student/faculty)
        const aud = slotDef(activeSlot).audience;
        if (aud === 'student' || aud === 'faculty') {
            opts.sort(function (a, b) {
                const la = String(a.name || a.shortLabel || '').toLowerCase();
                const lb = String(b.name || b.shortLabel || '').toLowerCase();
                const ka =
                    aud === 'student'
                        ? /schüler|student/.test(la)
                            ? 0
                            : 1
                        : /lehr|faculty/.test(la)
                          ? 0
                          : 1;
                const kb =
                    aud === 'student'
                        ? /schüler|student/.test(lb)
                            ? 0
                            : 1
                        : /lehr|faculty/.test(lb)
                          ? 0
                          : 1;
                return ka - kb;
            });
        }
        sel.replaceChildren();
        const o0 = document.createElement('option');
        o0.value = '';
        o0.textContent = opts.length ? '(Lizenz wählen)' : '(keine freie Lizenz)';
        sel.appendChild(o0);
        opts.forEach(function (o) {
            const opt = document.createElement('option');
            opt.value = o.skuId;
            const rest = o.remaining == null ? '' : ' · ' + o.remaining + ' frei';
            opt.textContent = (o.name || o.shortLabel || o.skuId) + rest;
            opt.disabled = !!o.disabled;
            sel.appendChild(opt);
        });
    }

    async function createLicenseGroup() {
        const def = slotDef(activeSlot);
        try {
            const token = await gug().getGraphToken();
            const nick =
                typeof gug().sanitizeMailNickname === 'function'
                    ? gug().sanitizeMailNickname(def.defaultNick)
                    : def.defaultNick;
            const body = {
                displayName: def.defaultName,
                description: 'MS365-Schul-Tools – Lizenzgruppe ' + def.label,
                mailEnabled: false,
                mailNickname: nick,
                securityEnabled: true,
                groupTypes: []
            };
            const created = await gug().graphJson('POST', '/groups', token, body);
            setSlotGroupId(activeSlot, created.id);
            log('Sicherheitsgruppe angelegt: ' + (created.displayName || created.id), 'ok');
            toast('Lizenzgruppe angelegt.');
            await refreshDetail();
        } catch (e) {
            log('Anlegen: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Anlegen fehlgeschlagen.');
        }
    }

    async function searchSecurityGroups() {
        const inp = $('lvGroupSearch');
        const host = $('lvGroupSearchResults');
        const q = inp && inp.value ? String(inp.value).trim() : '';
        if (!q) {
            toast('Bitte Suchbegriff eingeben.');
            return;
        }
        if (!host) return;
        host.style.display = 'block';
        host.innerHTML = '<div class="lv-result-row muted">Suche …</div>';
        try {
            const token = await gug().getGraphToken();
            const esc = typeof gug().odataEscape === 'function' ? gug().odataEscape(q) : q.replace(/'/g, "''");
            let list = [];
            try {
                const phrase = '"' + q.replace(/"/g, '') + '"';
                const path =
                    '/groups?$search=' +
                    encodeURIComponent('"displayName:' + phrase + ' OR mailNickname:' + phrase + '"') +
                    '&$select=' +
                    encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled') +
                    '&$top=25';
                const data = await gug().graphJson('GET', path, token, undefined, {
                    ConsistencyLevel: 'eventual'
                });
                list = (data && data.value) || [];
            } catch {
                const filter =
                    "(startswith(displayName,'" +
                    esc +
                    "') or startswith(mailNickname,'" +
                    esc +
                    "'))";
                const path =
                    '/groups?$filter=' +
                    encodeURIComponent(filter) +
                    '&$select=' +
                    encodeURIComponent('id,displayName,mail,mailNickname,groupTypes,securityEnabled,mailEnabled') +
                    '&$top=25';
                const data = await gug().graphJson('GET', path, token);
                list = (data && data.value) || [];
            }
            const filtered = list.filter(function (g) {
                if (!g || !g.securityEnabled) return false;
                const types = Array.isArray(g.groupTypes) ? g.groupTypes : [];
                if (types.indexOf('Unified') !== -1) return false;
                return true;
            });
            host.replaceChildren();
            if (!filtered.length) {
                host.innerHTML =
                    '<div class="lv-result-row muted">Keine passende Sicherheitsgruppe gefunden.</div>';
                return;
            }
            filtered.forEach(function (g) {
                const row = document.createElement('div');
                row.className = 'lv-result-row';
                const left = document.createElement('div');
                left.innerHTML =
                    '<div style="font-weight:700;">' +
                    escapeHtml(g.displayName || '(ohne Namen)') +
                    '</div><div class="muted" style="font-size:0.9em;">Alias: <code>' +
                    escapeHtml(g.mailNickname || '–') +
                    '</code>' +
                    (g.mailEnabled ? ' · mail-fähig' : '') +
                    '</div>';
                const btn = document.createElement('button');
                btn.type = 'button';
                btn.className = 'btn btn-primary';
                btn.innerHTML = '<i class="bi bi-link-45deg"></i> Zuordnen';
                btn.addEventListener('click', function () {
                    setSlotGroupId(activeSlot, g.id);
                    host.style.display = 'none';
                    host.replaceChildren();
                    log('Gruppe zugeordnet: ' + (g.displayName || g.id), 'ok');
                    toast('Sicherheitsgruppe zugeordnet.');
                    void refreshDetail();
                });
                row.appendChild(left);
                row.appendChild(btn);
                host.appendChild(row);
            });
        } catch (e) {
            host.innerHTML =
                '<div class="lv-result-row" style="color:#b02a37;">' +
                escapeHtml(e && e.message ? e.message : String(e)) +
                '</div>';
        }
    }

    async function assignSku() {
        const sel = $('lvSkuSelect');
        const skuId = sel && sel.value ? String(sel.value).trim() : '';
        const gid = getSlotGroupId(activeSlot);
        if (!gid || !skuId) {
            toast('Bitte Gruppe und Lizenz wählen.');
            return;
        }
        try {
            const token = await gug().getGraphToken();
            await gug().graphJson('POST', '/groups/' + encodeURIComponent(gid) + '/assignLicense', token, {
                addLicenses: [{ skuId: skuId, disabledPlans: [] }],
                removeLicenses: []
            });
            log('Lizenz zugewiesen: ' + skuLabel(skuId, findPartNumber(skuId)), 'ok');
            toast('Lizenz der Gruppe zugewiesen.');
            await refreshDetail();
        } catch (e) {
            log('Zuweisen: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Zuweisen fehlgeschlagen (Rechte / freie Lizenzen prüfen).');
        }
    }

    async function removeSku(skuId) {
        const gid = getSlotGroupId(activeSlot);
        if (!gid || !skuId) return;
        const ok =
            typeof window.ms365AppDialogConfirm === 'function'
                ? await window.ms365AppDialogConfirm(
                      'Lizenz wirklich von der Gruppe entziehen? Mitglieder verlieren die Lizenz über diese Gruppe.',
                      { title: 'Lizenz entziehen', okText: 'Entziehen' }
                  )
                : window.confirm('Lizenz von der Gruppe entziehen?');
        if (!ok) return;
        try {
            const token = await gug().getGraphToken();
            await gug().graphJson('POST', '/groups/' + encodeURIComponent(gid) + '/assignLicense', token, {
                addLicenses: [],
                removeLicenses: [skuId]
            });
            log('Lizenz entzogen: ' + skuLabel(skuId, findPartNumber(skuId)), 'ok');
            toast('Lizenz entzogen.');
            await refreshDetail();
        } catch (e) {
            log('Entziehen: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Entziehen fehlgeschlagen.');
        }
    }

    function groupMembersArray(mem) {
        if (Array.isArray(mem)) return mem;
        if (mem && Array.isArray(mem.items)) return mem.items;
        return [];
    }

    function renderMembers(list) {
        const tbody = $('lvMembersBody');
        const meta = $('lvMembersMeta');
        if (!tbody) return;
        tbody.replaceChildren();
        const arr = Array.isArray(list) ? list : [];
        if (meta) meta.textContent = arr.length ? arr.length + ' Mitglied(er)' : 'Keine Mitglieder';
        if (!arr.length) {
            const tr = document.createElement('tr');
            tr.innerHTML = '<td colspan="2" class="muted">Keine Mitglieder.</td>';
            tbody.appendChild(tr);
            return;
        }
        arr.forEach(function (u) {
            const tr = document.createElement('tr');
            const name = normStr(u.displayName) || '–';
            const mail = normStr(u.userPrincipalName || u.mail) || '–';
            tr.innerHTML =
                '<td>' + escapeHtml(name) + '</td><td><code>' + escapeHtml(mail) + '</code></td>';
            tbody.appendChild(tr);
        });
    }

    async function loadMembers(token, gid) {
        const meta = $('lvMembersMeta');
        if (meta) meta.textContent = 'Lade Mitglieder …';
        try {
            const members = await gug().fetchGroupMembers(token, gid);
            const users = groupMembersArray(members).filter(function (m) {
                const t = String((m && m['@odata.type']) || '');
                return !t || t.indexOf('user') !== -1 || m.userPrincipalName || m.mail;
            });
            renderMembers(users);
        } catch (e) {
            renderMembers([]);
            if (meta) meta.textContent = 'Mitglieder nicht lesbar';
            log('Mitglieder: ' + (e && e.message ? e.message : String(e)), 'warn');
        }
    }

    async function syncFromSammel() {
        const def = slotDef(activeSlot);
        const targetId = getSlotGroupId(activeSlot);
        if (!targetId) {
            toast('Zuerst Lizenzgruppe zuordnen oder anlegen.');
            return;
        }
        const sourceId = getMatchedSammelId(def.sammelKey);
        if (!sourceId) {
            toast(
                'Keine Sammelgruppe in den Stammdaten gematcht. Bitte unter „Alle Schüler und Lehrkräfte“ bzw. Verwaltung zuerst matchen.'
            );
            return;
        }
        try {
            const token = await gug().getGraphToken();
            const memberList = groupMembersArray(await gug().fetchGroupMembers(token, sourceId));
            let ok = 0;
            let skip = 0;
            let fail = 0;
            for (let i = 0; i < memberList.length; i++) {
                const m = memberList[i];
                const uid = m && m.id ? String(m.id) : '';
                if (!uid) continue;
                try {
                    await gug().graphAddMember(token, targetId, uid);
                    ok++;
                } catch (e) {
                    if (gug().isDuplicateMemberError && gug().isDuplicateMemberError(e)) skip++;
                    else fail++;
                }
            }
            log(
                'Übernahme aus Sammelgruppe: +' + ok + ', schon drin ' + skip + ', Fehler ' + fail,
                fail ? 'warn' : 'ok'
            );
            toast('Mitglieder übernommen (' + ok + ' neu).');
            await loadMembers(token, targetId);
        } catch (e) {
            log('Übernahme: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Übernahme fehlgeschlagen.');
        }
    }

    async function enableDynamic() {
        const gid = getSlotGroupId(activeSlot);
        const ruleEl = $('lvDynamicRule');
        const rule = ruleEl && ruleEl.value ? String(ruleEl.value).trim() : '';
        if (!gid || !rule) {
            toast('Gruppe und Regel erforderlich.');
            return;
        }
        try {
            const token = await gug().getGraphToken();
            await gug().graphJson('PATCH', '/groups/' + encodeURIComponent(gid), token, {
                groupTypes: ['DynamicMembership'],
                membershipRule: rule,
                membershipRuleProcessingState: 'On'
            });
            log('Dynamische Mitgliedschaft aktiviert.', 'ok');
            toast('Regel aktiviert (Verarbeitung kann etwas dauern).');
            await refreshDetail();
        } catch (e) {
            log('Dynamisch: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Regel fehlgeschlagen – oft fehlt Entra ID P1 oder die Syntax.');
        }
    }

    async function disableDynamic() {
        const gid = getSlotGroupId(activeSlot);
        if (!gid) return;
        try {
            const token = await gug().getGraphToken();
            await gug().graphJson('PATCH', '/groups/' + encodeURIComponent(gid), token, {
                membershipRuleProcessingState: 'Paused'
            });
            log('Dynamische Mitgliedschaft pausiert.', 'ok');
            toast('Regel pausiert.');
            await refreshDetail();
        } catch (e) {
            log('Pausieren: ' + (e && e.message ? e.message : String(e)), 'err');
            toast('Pausieren fehlgeschlagen.');
        }
    }

    function applyDeptTemplate() {
        const def = slotDef(activeSlot);
        const hint = def.deptHint || def.label;
        const ruleEl = $('lvDynamicRule');
        if (ruleEl) {
            ruleEl.value = '(user.department -eq "' + hint + '")';
        }
        toast('Vorlage gesetzt – bitte an eure Abteilungsnamen anpassen.');
    }

    async function addCustomProfile() {
        const name =
            typeof window.ms365AppDialogPrompt === 'function'
                ? await window.ms365AppDialogPrompt('Name für das Lizenz-Profil', 'Lizenz Zusatz', {
                      title: 'Eigenes Profil',
                      inputLabel: 'Bezeichnung',
                      okText: 'Anlegen'
                  })
                : window.prompt('Name für das Lizenz-Profil', 'Lizenz Zusatz');
        if (!name || !String(name).trim()) return;
        const id = 'c' + Date.now().toString(36);
        state.custom = state.custom || [];
        state.custom.push({ id: id, label: String(name).trim(), groupId: null });
        saveState();
        renderSlotList();
        selectSlot(id);
    }

    function wire() {
        renderSlotList();
        $('lvBtnCreateGroup')?.addEventListener('click', function () {
            void createLicenseGroup();
        });
        $('lvBtnSearchGroup')?.addEventListener('click', function () {
            void searchSecurityGroups();
        });
        $('lvGroupSearch')?.addEventListener('keydown', function (ev) {
            if (ev.key === 'Enter') {
                ev.preventDefault();
                void searchSecurityGroups();
            }
        });
        $('lvBtnUnmatch')?.addEventListener('click', function () {
            setSlotGroupId(activeSlot, null);
            currentGroup = null;
            setGroupUi(null);
            renderAssignedSkus([]);
            renderMembers([]);
            toast('Zuordnung gelöst.');
        });
        $('lvBtnOpenEntra')?.addEventListener('click', function () {
            const gid = getSlotGroupId(activeSlot);
            if (gid) window.open(ENTRA_GROUP + encodeURIComponent(gid), '_blank', 'noopener');
        });
        $('lvBtnAssignSku')?.addEventListener('click', function () {
            void assignSku();
        });
        $('lvBtnReloadSkus')?.addEventListener('click', async function () {
            try {
                skusOk = false;
                const token = await gug().getGraphToken();
                await ensureSkus(token);
                fillSkuSelect((currentGroup && currentGroup.assignedLicenses) || []);
                toast('SKUs aktualisiert.');
            } catch (e) {
                toast('SKU-Laden fehlgeschlagen.');
            }
        });
        $('lvBtnReloadMembers')?.addEventListener('click', async function () {
            const gid = getSlotGroupId(activeSlot);
            if (!gid) return;
            try {
                const token = await gug().getGraphToken();
                await loadMembers(token, gid);
            } catch (e) {
                toast('Mitglieder laden fehlgeschlagen.');
            }
        });
        $('lvBtnSyncFromSammel')?.addEventListener('click', function () {
            void syncFromSammel();
        });
        $('lvBtnDynTemplateDept')?.addEventListener('click', applyDeptTemplate);
        $('lvBtnEnableDynamic')?.addEventListener('click', function () {
            void enableDynamic();
        });
        $('lvBtnDisableDynamic')?.addEventListener('click', function () {
            void disableDynamic();
        });
        $('lvBtnAddCustom')?.addEventListener('click', function () {
            void addCustomProfile();
        });
        void selectSlot('schueler');
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wire);
    } else {
        wire();
    }
})();
