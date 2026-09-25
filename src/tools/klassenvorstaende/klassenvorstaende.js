(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-klassenvorstaende-v1';

    function gug() {
        const G = window.ms365GraphUnifiedGroups;
        if (!G) throw new Error('graph-unified-groups.js muss vor diesem Skript geladen werden.');
        return G;
    }

    function live() {
        const L = window.ms365SlgLiveDetails;
        if (!L) throw new Error('slg-live-details.js muss vor diesem Skript geladen werden.');
        return L;
    }

    function gd() {
        const G = window.ms365GroupDetail;
        if (!G) throw new Error('group-detail.js muss vor diesem Skript geladen werden.');
        return G;
    }

    async function getGraphToken() {
        return gug().getGraphToken();
    }

    /** @type {string|null} */
    let matchedGroupId = null;

    /** @type {{ emails: string[], direktion: string[], rows: { code: string, name: string, year: string, headName: string, headEmail: string }[] }} */
    let listCache = { emails: [], direktion: [], rows: [] };

    /** @type {Record<string, number>} */
    let graphMemberCounts = {};
    let countsFetchGen = 0;

    /** @type {{ gid: string, diff: object, graphByEmail: Map<string, object> }|null} */
    let deviationReviewState = null;

    function toast(msg) {
        const el = document.getElementById('toast');
        if (el) {
            el.textContent = msg;
            el.classList.add('show');
            clearTimeout(toast._t);
            toast._t = setTimeout(function () {
                el.classList.remove('show');
            }, 3800);
        } else if (typeof window.ms365ToastOrAlert === 'function') {
            window.ms365ToastOrAlert(msg);
        } else {
            window.alert(msg);
        }
    }

    function dlgConfirm(message, options) {
        if (typeof window.ms365AppDialogConfirm === 'function') {
            return window.ms365AppDialogConfirm(message, options || {});
        }
        return Promise.resolve(window.confirm(message));
    }

    function normStr(v) {
        return String(v ?? '').trim();
    }
    function normEmail(v) {
        return normStr(v).toLowerCase();
    }

    function escapeHtml(s) {
        return String(s)
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function loadTenantSettings() {
        if (typeof window.ms365TenantSettingsLoad !== 'function') return null;
        return window.ms365TenantSettingsLoad();
    }

    function isDirektionRole(roleRaw) {
        const r = normStr(roleRaw).toLowerCase();
        if (!r) return false;
        return r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1;
    }

    function collectDirektionOwnerEmails(settings) {
        const out = [];
        const seen = new Set();
        const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
        admin.forEach(function (row) {
            if (!isDirektionRole(row && row.role)) return;
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function collectKlassenvorstaende(settings) {
        const classes = settings && Array.isArray(settings.classes) ? settings.classes : [];
        const rows = [];
        const emails = [];
        const seen = new Set();
        classes.forEach(function (c) {
            if (!c) return;
            const code = normStr(c.code);
            const name = normStr(c.name);
            const year = normStr(c.year);
            const headName = normStr(c.headName);
            const headEmail = normEmail(c.headEmail);
            rows.push({ code: code, name: name, year: year, headName: headName, headEmail: headEmail });
            if (!headEmail || headEmail.indexOf('@') === -1) return;
            if (seen.has(headEmail)) return;
            seen.add(headEmail);
            emails.push(headEmail);
        });
        rows.sort(function (a, b) {
            return (a.code || a.name).localeCompare(b.code || b.name, 'de');
        });
        return { rows: rows, emails: emails };
    }

    function readLists() {
        const settings = loadTenantSettings();
        const kv = collectKlassenvorstaende(settings);
        listCache.rows = kv.rows;
        listCache.emails = kv.emails;
        listCache.direktion = collectDirektionOwnerEmails(settings);
    }

    async function ensureOwners(token, groupId) {
        return gug().ensureOwners(token, groupId, listCache.direktion || []);
    }

    function appendSyncLog(msg, kind) {
        const el = document.getElementById('slgSyncLog');
        if (!el) return;
        const line = document.createElement('div');
        line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
        if (kind === 'err') line.style.color = '#b00020';
        else if (kind === 'ok') line.style.color = '#0d8050';
        else if (kind === 'warn') line.style.color = '#856404';
        el.appendChild(line);
        el.scrollTop = el.scrollHeight;
    }

    function clearSyncLog() {
        const el = document.getElementById('slgSyncLog');
        if (el) el.replaceChildren();
    }

    function graphCountFor(groupId) {
        const id = String(groupId || '').trim();
        if (!id) return null;
        const n = graphMemberCounts[id];
        return typeof n === 'number' && n >= 0 ? n : null;
    }

    function rememberGraphMemberCount(groupId, count) {
        const id = String(groupId || '').trim();
        if (!id || typeof count !== 'number' || count < 0) return;
        graphMemberCounts[id] = count;
        paintCounts();
        updateMismatchUi();
    }

    function paintCounts() {
        const listN = (listCache.emails || []).length;
        const gid = matchedGroupId;
        const groupN = graphCountFor(gid);
        const wrap = document.getElementById('kvGroupCounts');
        const listEl = document.getElementById('kvListCount');
        const groupEl = document.getElementById('kvGroupCount');
        const line = document.getElementById('kvGroupLine');
        if (listEl) listEl.textContent = String(listN);
        if (groupEl) groupEl.textContent = gid ? (groupN === null ? '–' : String(groupN)) : '–';
        if (line) line.textContent = gid ? 'Gematcht: ' + gid : 'Noch kein Match';
        if (!wrap) return;
        wrap.classList.remove('is-match', 'is-mismatch');
        const known = gid && groupN !== null;
        if (known) {
            const same = listN === groupN;
            wrap.classList.add(same ? 'is-match' : 'is-mismatch');
            wrap.title = same
                ? 'KV-Liste und Gruppe: je ' + listN + ' – Anzahl stimmt überein.'
                : 'KV-Liste: ' + listN + ' · Gruppe: ' + groupN + ' Mitglieder.';
        } else {
            wrap.title = gid
                ? 'KV-Liste: ' + listN + ' E-Mails. Mitgliederzahl wird aus Microsoft Graph geladen.'
                : 'KV-Liste: ' + listN + ' eindeutige E-Mails. Noch keine Gruppe gematcht.';
        }
    }

    function updateMismatchUi() {
        paintCounts();
        const bar = document.getElementById('kvMismatchBar');
        const actions = document.getElementById('kvMismatchActions');
        const listN = (listCache.emails || []).length;
        const gid = matchedGroupId;
        const groupN = graphCountFor(gid);
        const mismatch = !!(gid && groupN !== null && listN !== groupN);
        if (bar) bar.hidden = !mismatch;
        if (!actions) return;
        actions.replaceChildren();
        if (!mismatch) return;
        const btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'btn btn-sm';
        btn.innerHTML = '<i class="bi bi-arrow-left-right"></i>Abgleich öffnen';
        btn.addEventListener('click', function () {
            void loadMembershipReview();
        });
        actions.appendChild(btn);
    }

    async function refreshGraphMemberCounts() {
        const gid = matchedGroupId;
        if (!gid) {
            updateMismatchUi();
            return;
        }
        const gen = ++countsFetchGen;
        try {
            const token = await getGraphToken();
            if (gen !== countsFetchGen) return;
            const n = await gug().fetchGroupMemberCount(token, gid);
            if (typeof n === 'number' && n >= 0) graphMemberCounts[gid] = n;
            if (gen !== countsFetchGen) return;
            updateMismatchUi();
        } catch {
            updateMismatchUi();
        }
    }

    function renderClassTable() {
        const body = document.getElementById('kvClassBody');
        const summary = document.getElementById('kvClassSummary');
        const rows = listCache.rows || [];
        const withMail = rows.filter(function (r) {
            return r.headEmail && r.headEmail.indexOf('@') !== -1;
        }).length;
        if (summary) {
            summary.textContent =
                rows.length +
                ' Klasse' +
                (rows.length === 1 ? '' : 'n') +
                ' · ' +
                withMail +
                ' mit KV‑E‑Mail · ' +
                listCache.emails.length +
                ' eindeutige Adresse' +
                (listCache.emails.length === 1 ? '' : 'n');
        }
        if (!body) return;
        body.replaceChildren();
        if (!rows.length) {
            const tr = document.createElement('tr');
            tr.innerHTML = '<td colspan="3" class="muted">Keine Klassen im aktuellen Schuljahr.</td>';
            body.appendChild(tr);
            return;
        }
        rows.forEach(function (r) {
            const tr = document.createElement('tr');
            const klasse = escapeHtml(r.code || r.name || '–');
            const kv = escapeHtml(r.headName || '–');
            const em = r.headEmail
                ? '<span class="kv-ok">' + escapeHtml(r.headEmail) + '</span>'
                : '<span class="kv-missing">fehlt</span>';
            tr.innerHTML = '<td>' + klasse + '</td><td>' + kv + '</td><td>' + em + '</td>';
            body.appendChild(tr);
        });
    }

    function renderOwnerPreview() {
        const el = document.getElementById('slgOwnerPreview');
        if (!el) return;
        el.replaceChildren();
        const list = listCache.direktion || [];
        if (!list.length) {
            el.textContent = 'Keine Direktion in den Stammdaten – beim Anlegen wird der angemeldete Benutzer Besitzer.';
            return;
        }
        list.forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
    }

    function renderMemberPreview() {
        const el = document.getElementById('slgMemberPreview');
        if (!el) return;
        el.replaceChildren();
        const list = listCache.emails || [];
        if (!list.length) {
            el.textContent = 'Keine KV‑E‑Mails in der Klassenliste.';
            return;
        }
        list.slice(0, 30).forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        if (list.length > 30) {
            const more = document.createElement('div');
            more.className = 'muted';
            more.style.padding = '6px 0 0';
            more.textContent = '… und ' + (list.length - 30) + ' weitere';
            el.appendChild(more);
        }
    }

    function updateLeftListUi() {
        paintCounts();
        renderClassTable();
        updateMismatchUi();
    }

    function getActiveMatchedId() {
        return matchedGroupId;
    }

    function setActiveMatchedId(id) {
        matchedGroupId = id ? String(id) : null;
        live().resetCaches();
    }

    function applyCreateDefaults() {
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        const mail = document.getElementById('slgNewCreateTargetMail');
        const team = document.getElementById('slgNewCreateTargetTeam');
        let draft = null;
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                draft = su && su.kvDraft ? su.kvDraft : null;
            }
        } catch {
            draft = null;
        }
        if (dn && !dn.value) dn.value = (draft && draft.kvNewDisplayName) || 'Klassenvorstände';
        if (nn && !nn.value) nn.value = (draft && draft.kvNewMailNick) || 'klassenvorstaende';
        if (desc && !desc.value) {
            desc.value =
                (draft && draft.kvNewDescription) ||
                'Klassenvorstände aus den Stammdaten (MS365-Schul-Tools)';
        }
        const wantTeam = draft && draft.kvNewCreateTarget === 'team';
        if (team && mail) {
            team.checked = !!wantTeam;
            mail.checked = !wantTeam;
        }
    }

    function readCreateTarget() {
        const team = document.getElementById('slgNewCreateTargetTeam');
        return team && team.checked ? 'team' : 'mail';
    }

    async function runSyncMembers() {
        const gid = getActiveMatchedId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const emails = listCache.emails || [];
        if (!emails.length) {
            toast('Keine KV‑E‑Mails in der Klassenliste.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Klassenvorstände (' + emails.length + ' Adressen) …', '');
        try {
            const token = await getGraphToken();
            let joinEmails = emails;
            let leaveEmails = [];
            if (typeof gug().fetchGroupMembers === 'function') {
                const mem = await gug().fetchGroupMembers(token, gid);
                const current = (mem.items || [])
                    .map(function (m) {
                        return String((m && (m.mail || m.userPrincipalName)) || '')
                            .trim()
                            .toLowerCase();
                    })
                    .filter(function (em) {
                        return em.indexOf('@') !== -1;
                    });
                const lc = window.ms365StudentClassLifecycle;
                if (lc && typeof lc.reconcileSammelgruppe === 'function') {
                    const rec = lc.reconcileSammelgruppe(emails, current);
                    joinEmails = rec.join;
                    leaveEmails = rec.leave;
                } else {
                    const M = window.ms365MembershipReconcile;
                    if (M && typeof M.diffMemberships === 'function') {
                        const diff = M.diffMemberships(emails, current);
                        joinEmails = diff.onlyLocal;
                        leaveEmails = diff.onlyGraph;
                    }
                }
                appendSyncLog('Abgleich mit KV‑Liste: +' + joinEmails.length + ' / −' + leaveEmails.length + '.', '');
            }
            if (joinEmails.length) {
                const r = await gug().syncEmailsToGroup(token, gid, joinEmails, 'KV', appendSyncLog);
                appendSyncLog('Aufnehmen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (leaveEmails.length && typeof gug().removeEmailsFromGroup === 'function') {
                const r = await gug().removeEmailsFromGroup(token, gid, leaveEmails, 'KV', appendSyncLog);
                appendSyncLog('Entfernen: ' + r.ok + ' OK, übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (!joinEmails.length && !leaveEmails.length) {
                appendSyncLog('Keine Änderungen gegenüber der KV‑Liste.', 'ok');
            }
            await ensureOwners(token, gid);
            live().invalidateMembership();
            await live().loadMembers();
            await refreshGraphMemberCounts();
            if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                window.ms365ActionLog.append({
                    tool: 'klassenvorstaende',
                    action: 'sync-members',
                    target: gid,
                    summary: 'KV-Mitglieder synchronisiert (+' + joinEmails.length + '/−' + leaveEmails.length + ')'
                });
            }
            toast('Synchronisation abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    function renderDeviationReviewPanel(state) {
        const body = document.getElementById('kvDeviationBody');
        const summaryEl = document.getElementById('kvDeviationSummary');
        const actions = document.getElementById('kvDeviationActions');
        const R = window.ms365MembershipReviewRender;
        if (!body || !summaryEl || !state || !R) return;

        summaryEl.textContent =
            'Markieren Sie Personen und wählen Sie unten in der passenden Spalte die Aktion.';
        if (actions) {
            R.collectMembershipReviewActionButtons(actions);
            actions.hidden = false;
            actions.setAttribute('aria-hidden', 'false');
        }

        body.replaceChildren();
        body.appendChild(
            R.buildMembershipReviewBody({
                diff: state.diff,
                graphByEmail: state.graphByEmail,
                listCount: (listCache.emails || []).length,
                labels: {
                    onlyLocalTitle: 'Nur in der KV‑Liste',
                    onlyLocalHint: 'In den Stammdaten (Klassen → KV), aber nicht in der Microsoft-365-Gruppe.',
                    onlyGraphTitle: 'Nur in der Microsoft-365-Gruppe',
                    onlyGraphHint: 'In der Gruppe online, aber nicht als Klassenvorstand in den Stammdaten.'
                }
            })
        );
        R.attachMembershipReviewSectionActions(body, {
            onlyLocal: state.diff.onlyLocal.length,
            onlyGraph: state.diff.onlyGraph.length
        });
    }

    function hideDeviationPanel() {
        const panel = document.getElementById('kvDeviationPanel');
        if (panel) panel.hidden = true;
        deviationReviewState = null;
    }

    function getSelectedEmails(sectionKey) {
        const out = [];
        document.querySelectorAll('input[data-mr-section="' + sectionKey + '"]:checked').forEach(function (cb) {
            const em = normEmail(cb.getAttribute('data-mr-email'));
            if (em.indexOf('@') !== -1) out.push(em);
        });
        return out;
    }

    function mr() {
        const M = window.ms365MembershipReconcile;
        if (!M) throw new Error('membership-reconcile.js muss vor diesem Skript geladen werden.');
        return M;
    }

    async function loadMembershipReview() {
        const gid = getActiveMatchedId();
        const panel = document.getElementById('kvDeviationPanel');
        const body = document.getElementById('kvDeviationBody');
        const titleEl = document.getElementById('kvDeviationTitle');
        const summaryEl = document.getElementById('kvDeviationSummary');
        const actions = document.getElementById('kvDeviationActions');
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        if (!panel || !body || !titleEl || !summaryEl) return;

        titleEl.textContent = 'Mitglieder-Abgleich: Klassenvorstände';
        summaryEl.textContent = 'Vergleich wird geladen …';
        if (actions && window.ms365MembershipReviewRender) {
            window.ms365MembershipReviewRender.collectMembershipReviewActionButtons(actions);
        }
        body.replaceChildren();
        panel.hidden = false;
        panel.scrollIntoView({ behavior: 'smooth', block: 'nearest' });

        try {
            const token = await getGraphToken();
            const mem = await gug().fetchGroupMembers(token, gid);
            const graphEmails = (mem.items || [])
                .map(function (m) {
                    return mr().memberEmailFromGraph(m);
                })
                .filter(function (em) {
                    return em.indexOf('@') !== -1;
                });
            const diff = mr().diffMemberships(listCache.emails || [], graphEmails);
            deviationReviewState = {
                gid: gid,
                diff: diff,
                graphByEmail: mr().indexGraphMembersByEmail(mem.items || [])
            };
            renderDeviationReviewPanel(deviationReviewState);
        } catch (e) {
            deviationReviewState = null;
            summaryEl.textContent = '';
            const p = document.createElement('p');
            p.className = 'slg-deviation-panel__error';
            p.textContent = 'Abgleich konnte nicht geladen werden: ' + (e.message || e);
            body.replaceChildren(p);
        }
    }

    async function applyMrAddToGroup() {
        if (!deviationReviewState) return;
        const emails = getSelectedEmails('onlyLocal');
        if (!emails.length) {
            toast('Bitte Adressen auswählen.');
            return;
        }
        try {
            const token = await getGraphToken();
            await gug().syncEmailsToGroup(token, deviationReviewState.gid, emails, 'KV', appendSyncLog);
            toast(emails.length + ' aufgenommen.');
            await loadMembershipReview();
            await refreshGraphMemberCounts();
            live().invalidateMembership();
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function applyMrRemoveFromGroup() {
        if (!deviationReviewState) return;
        const emails = getSelectedEmails('onlyGraph');
        if (!emails.length) {
            toast('Bitte Adressen auswählen.');
            return;
        }
        const ok = await dlgConfirm(emails.length + ' Person(en) aus der Gruppe entfernen?', {
            title: 'Aus Gruppe entfernen'
        });
        if (!ok) return;
        try {
            const token = await getGraphToken();
            await gug().removeEmailsFromGroup(token, deviationReviewState.gid, emails, 'KV', appendSyncLog);
            toast(emails.length + ' entfernt.');
            await loadMembershipReview();
            await refreshGraphMemberCounts();
            live().invalidateMembership();
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function buildStateObject() {
        return {
            kind: STORAGE_KEY,
            savedAt: new Date().toISOString(),
            matched: { kvGroupId: matchedGroupId },
            kvNewDisplayName: document.getElementById('slgNewDisplayName')
                ? document.getElementById('slgNewDisplayName').value
                : '',
            kvNewMailNick: document.getElementById('slgNewMailNick')
                ? document.getElementById('slgNewMailNick').value
                : '',
            kvNewDescription: document.getElementById('slgNewDescription')
                ? document.getElementById('slgNewDescription').value
                : '',
            kvNewCreateTarget: readCreateTarget()
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        if (o.matched && typeof o.matched === 'object') {
            matchedGroupId = o.matched.kvGroupId ? String(o.matched.kvGroupId) : null;
        } else if (o.kvGroupId) {
            matchedGroupId = String(o.kvGroupId);
        }
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const mail = document.getElementById('slgNewCreateTargetMail');
        const team = document.getElementById('slgNewCreateTargetTeam');
        if (dn && o.kvNewDisplayName !== undefined) dn.value = String(o.kvNewDisplayName || '');
        if (nn && o.kvNewMailNick !== undefined) nn.value = String(o.kvNewMailNick || '');
        if (dd && o.kvNewDescription !== undefined) dd.value = String(o.kvNewDescription || '');
        if (mail && team && o.kvNewCreateTarget !== undefined) {
            const wantTeam = String(o.kvNewCreateTarget) === 'team';
            team.checked = wantTeam;
            mail.checked = !wantTeam;
        }
        const gid = getActiveMatchedId();
        live().setMatchedMode(!!gid);
        live().fillForm(gid ? { id: gid } : null);
        updateLeftListUi();
        gd().setTab('general');
    }

    function saveState() {
        try {
            const obj = buildStateObject();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { kvGroupId: matchedGroupId },
                    kvDraft: {
                        kvNewDisplayName: obj.kvNewDisplayName,
                        kvNewMailNick: obj.kvNewMailNick,
                        kvNewDescription: obj.kvNewDescription,
                        kvNewCreateTarget: obj.kvNewCreateTarget
                    }
                });
            }
        } catch {
            // ignore
        }
    }

    function loadState() {
        let rawLocal = null;
        try {
            rawLocal = localStorage.getItem(STORAGE_KEY);
        } catch {
            rawLocal = null;
        }
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                const hasId = su && su.matched && !!su.matched.kvGroupId;
                if (hasId || !rawLocal) {
                    const d = su.kvDraft || {};
                    applyStateObject({
                        matched: su.matched,
                        kvNewDisplayName: d.kvNewDisplayName,
                        kvNewMailNick: d.kvNewMailNick,
                        kvNewDescription: d.kvNewDescription,
                        kvNewCreateTarget: d.kvNewCreateTarget
                    });
                    return;
                }
            }
        } catch {
            // ignore
        }
        try {
            if (!rawLocal) return;
            applyStateObject(JSON.parse(rawLocal));
        } catch {
            // ignore
        }
    }

    function clearStorage() {
        try {
            localStorage.removeItem(STORAGE_KEY);
            matchedGroupId = null;
            graphMemberCounts = {};
            countsFetchGen += 1;
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { kvGroupId: null }
                });
            }
            saveState();
            live().loadGroup({ silent: true });
            updateLeftListUi();
            toast('Zurückgesetzt.');
        } catch (e) {
            toast('Löschen fehlgeschlagen: ' + (e.message || e));
        }
    }

    function onClick(id, fn) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', fn);
    }

    function mountDetail() {
        gd().mount('#groupDetailHost', {
            title: 'Klassenvorstände',
            searchPlaceholder: 'z. B. klassenvorstand oder @schule.at',
            unmatchedCreateHint:
                'Wählen Sie das Ziel: nur E‑Mail‑Verteiler oder Microsoft 365‑Gruppe inkl. Team. Mitglieder kommen aus den KV‑E‑Mails der Klassenliste.',
            membersUnmatchedHint:
                'Mitglieder = eindeutige Klassenvorstand‑E‑Mails aus den Stammdaten. Nach dem Match können Sie live verwalten und synchronisieren.',
            membersUnmatchedTitle: 'Vorschau KV‑Liste (erste 30)',
            membersMatchedHint:
                'Live aus Microsoft Graph. „Mitglieder synchronisieren“ gleicht die Gruppe mit den KV‑Adressen ab: fehlende hinzufügen, nicht gelistete entfernen.',
            features: { syncMembers: true, membershipReview: true, createTargetChoice: true },
            live: {
                toast: toast,
                dlgConfirm: dlgConfirm,
                getGroupId: getActiveMatchedId,
                ensureDirektionOwners: function (token, gid) {
                    if (!(listCache.direktion && listCache.direktion.length)) {
                        throw new Error('Keine Direktion‑Adressen in den Stammdaten.');
                    }
                    return ensureOwners(token, gid);
                },
                onUnmatched: function () {
                    renderOwnerPreview();
                    renderMemberPreview();
                    updateLeftListUi();
                },
                onAfterLoad: function () {
                    updateLeftListUi();
                    return refreshGraphMemberCounts();
                },
                onMembersCount: function (groupId, count) {
                    rememberGraphMemberCount(groupId, count);
                }
            },
            match: {
                persistMatch: function (g) {
                    setActiveMatchedId(String(g.id));
                    saveState();
                },
                persistUnmatch: function () {
                    setActiveMatchedId(null);
                    saveState();
                },
                ensureOwners: function (token, gid) {
                    return ensureOwners(token, gid);
                },
                afterMatch: function () {
                    updateLeftListUi();
                    refreshGraphMemberCounts();
                },
                afterCreate: async function (token, g) {
                    const emails = listCache.emails || [];
                    if (!emails.length) return;
                    appendSyncLog('Mitglieder nach Anlegen: ' + emails.length + ' KV‑Adressen …', '');
                    await gug().syncEmailsToGroup(token, g.id, emails, 'KV', appendSyncLog);
                }
            },
            onTabUnmatched: function (tab) {
                if (tab === 'owners') renderOwnerPreview();
                if (tab === 'members') renderMemberPreview();
            }
        });
    }

    function wire() {
        onClick('kvBtnReloadLists', function () {
            readLists();
            updateLeftListUi();
            renderOwnerPreview();
            renderMemberPreview();
            toast('Listen neu eingelesen.');
        });
        onClick('slgBtnSync', function () {
            runSyncMembers();
        });
        onClick('slgBtnMembershipReview', function () {
            void loadMembershipReview();
        });
        onClick('kvMrAddToGroup', function () {
            void applyMrAddToGroup();
        });
        onClick('kvMrRemoveFromGroup', function () {
            void applyMrRemoveFromGroup();
        });
        onClick('kvMrReload', function () {
            void loadMembershipReview();
        });
        onClick('kvDeviationClose', function () {
            hideDeviationPanel();
        });
        onClick('kvBtnSaveState', function () {
            saveState();
            toast('Gespeichert.');
        });
        onClick('kvBtnLoadState', function () {
            loadState();
            toast('Geladen.');
            if (getActiveMatchedId()) live().loadGroup({ silent: true });
        });
        onClick('kvBtnClearStorage', function () {
            clearStorage();
        });
    }

    function init() {
        mountDetail();
        readLists();
        loadState();
        updateLeftListUi();
        renderOwnerPreview();
        renderMemberPreview();
        wire();
        if (!getActiveMatchedId()) {
            live().setMatchedMode(false);
            applyCreateDefaults();
        } else {
            live().loadGroup({ silent: true });
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
