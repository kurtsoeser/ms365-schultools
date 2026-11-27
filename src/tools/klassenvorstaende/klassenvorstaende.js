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
        try {
            const h = window.ms365MembershipHygiene;
            if (h && typeof h.patchHygieneScanGroupCount === 'function') {
                const listCount = (listCache.emails || []).length;
                h.patchHygieneScanGroupCount(id, count, { listCount: listCount });
                h.patchHygieneScanGroupCount('klassenvorstaende', count, { listCount: listCount });
            }
        } catch {
            /* Dashboard-Cache optional */
        }
    }

    function paintCounts() {
        const listN = (listCache.emails || []).length;
        const gid = matchedGroupId;
        const groupN = graphCountFor(gid);
        const wrap = document.getElementById('kvGroupCounts');
        const line = document.getElementById('kvGroupLine');
        if (line) {
            line.textContent = '';
            const badge = document.createElement('span');
            badge.className =
                'jg-match-badge gd-sync-badge ' +
                (gid ? 'is-ok gd-sync-badge--ok' : 'is-warn gd-sync-badge--warn');
            badge.innerHTML = gid
                ? '<i class="bi bi-check-circle-fill" aria-hidden="true"></i>Verknüpft'
                : '<i class="bi bi-exclamation-circle" aria-hidden="true"></i>Nicht verknüpft';
            line.appendChild(badge);
        }
        const G = window.ms365GroupDetail;
        const painted =
            G && typeof G.paintMasterMemberSyncCounts === 'function'
                ? G.paintMasterMemberSyncCounts(wrap, {
                      listN: listN,
                      groupN: groupN,
                      linked: !!gid,
                      memberWord: 'Mitglieder'
                  })
                : null;
        const syncLine = document.getElementById('slgSyncStatusLine');
        const preview = document.getElementById('slgSyncPreview');
        if (syncLine && !(preview && !preview.hidden)) {
            if (painted && painted.known && painted.same) {
                syncLine.textContent =
                    listN === 1 ? '✓ 1 Mitglied synchron' : '✓ ' + String(listN) + ' Mitglieder synchron';
            } else if (painted && painted.known) {
                syncLine.textContent =
                    '⚠ ' +
                    painted.delta +
                    (painted.delta === 1 ? ' Abweichung' : ' Abweichungen') +
                    ' (' +
                    listN +
                    ' erwartet · ' +
                    groupN +
                    ' in M365)';
            } else {
                syncLine.textContent = 'Mitglieder mit Stammdaten abgleichen.';
            }
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
            if (gen !== countsFetchGen) return;
            if (typeof n === 'number' && n >= 0) rememberGraphMemberCount(gid, n);
            else updateMismatchUi();
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

    /** @type {{ gid: string, joinEmails: string[], leaveEmails: string[], stayCount: number }|null} */
    let syncPreviewState = null;

    function hideSyncPreview() {
        syncPreviewState = null;
        const G = window.ms365GroupDetail;
        if (G && typeof G.setSyncPreviewVisible === 'function') G.setSyncPreviewVisible(false);
        else {
            const el = document.getElementById('slgSyncPreview');
            if (el) el.hidden = true;
        }
    }

    function showSyncPreview(state) {
        syncPreviewState = state;
        const G = window.ms365GroupDetail;
        if (G && typeof G.fillSyncPreviewLines === 'function') {
            G.fillSyncPreviewLines(state.joinEmails.length, state.leaveEmails.length, state.stayCount);
        }
        if (G && typeof G.setSyncPreviewVisible === 'function') G.setSyncPreviewVisible(true);
        else {
            const el = document.getElementById('slgSyncPreview');
            if (el) el.hidden = false;
        }
    }

    async function computeKvSyncDiff(token, gid, emails) {
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
        }
        const allowLeaveEl = document.getElementById('slgSyncAllowLeave');
        const allowLeave = !!(allowLeaveEl && allowLeaveEl.checked);
        if (leaveEmails.length && !allowLeave) leaveEmails = [];
        return {
            joinEmails: joinEmails,
            leaveEmails: leaveEmails,
            stayCount: Math.max(0, emails.length - joinEmails.length)
        };
    }

    async function runSyncPreview() {
        const gid = getActiveMatchedId();
        if (!gid) {
            toast('Zuerst eine Gruppe mit Microsoft 365 verknüpfen oder anlegen.');
            return;
        }
        const emails = listCache.emails || [];
        if (!emails.length) {
            toast('Keine KV‑E‑Mails in der Klassenliste.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Prüfen: Klassenvorstände (' + emails.length + ' Adressen) …', '');
        try {
            const token = await getGraphToken();
            const diff = await computeKvSyncDiff(token, gid, emails);
            appendSyncLog(
                'Vorschau: +' + diff.joinEmails.length + ' / −' + diff.leaveEmails.length + ' (noch keine Graph-Änderungen).',
                'ok'
            );
            showSyncPreview({
                gid: gid,
                joinEmails: diff.joinEmails,
                leaveEmails: diff.leaveEmails,
                stayCount: diff.stayCount
            });
            if (!diff.joinEmails.length && !diff.leaveEmails.length) {
                toast('Keine Änderungen gegenüber der KV‑Liste.');
            }
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
            hideSyncPreview();
        }
    }

    async function runSyncApply() {
        if (!syncPreviewState) {
            toast('Zuerst „Änderungen prüfen“.');
            return;
        }
        const state = syncPreviewState;
        const gid = state.gid;
        const joinEmails = state.joinEmails.slice();
        let leaveEmails = state.leaveEmails.slice();
        hideSyncPreview();
        appendSyncLog('Übernehmen: +' + joinEmails.length + ' / −' + leaveEmails.length + ' …', '');
        try {
            const token = await getGraphToken();
            const allowLeaveEl = document.getElementById('slgSyncAllowLeave');
            if (leaveEmails.length && !(allowLeaveEl && allowLeaveEl.checked)) leaveEmails = [];
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
            toast('Änderungen übernommen.');
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
            'Wer steht in der Schul-Liste, wer in der M365-Gruppe? Markieren und darunter die Sync-Richtung wählen.';
        if (actions) R.collectMembershipReviewActionButtons(actions);

        body.replaceChildren();
        body.appendChild(
            R.buildMembershipReviewBody({
                diff: state.diff,
                graphByEmail: state.graphByEmail,
                listCount: (listCache.emails || []).length,
                labels: {
                    onlyLocalTitle: 'Nur in der Schul-Liste',
                    onlyLocalHint:
                        'In den Stammdaten (Klassen → Klassenvorstand), aber nicht in der Microsoft-365-Gruppe.',
                    onlyGraphTitle: 'Nur in der Microsoft-365-Gruppe',
                    onlyGraphHint:
                        'In der Gruppe online, aber nicht als Klassenvorstand in den Stammdaten.'
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
            toast('Zuerst eine Gruppe mit Microsoft 365 verknüpfen oder anlegen.');
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
            if (mr().graphMemberCountFromDiff) {
                rememberGraphMemberCount(gid, mr().graphMemberCountFromDiff(diff));
            }
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
            live().invalidateMembership();
            await refreshGraphMemberCounts();
            await loadMembershipReview();
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
            live().invalidateMembership();
            await refreshGraphMemberCounts();
            await loadMembershipReview();
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
        let localObj = null;
        try {
            rawLocal = localStorage.getItem(STORAGE_KEY);
            if (rawLocal) localObj = JSON.parse(rawLocal);
        } catch {
            rawLocal = null;
            localObj = null;
        }
        const localKv =
            localObj && localObj.matched && localObj.matched.kvGroupId
                ? String(localObj.matched.kvGroupId).trim()
                : localObj && localObj.kvGroupId
                  ? String(localObj.kvGroupId).trim()
                  : '';
        try {
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.getSetup === 'function') {
                const su = window.ms365AppDataV2.getSetup();
                const hasId = su && su.matched && !!su.matched.kvGroupId;
                if (hasId || localKv || !rawLocal) {
                    const d = su.kvDraft || {};
                    const merged = su.matched && typeof su.matched === 'object' ? { ...su.matched } : {};
                    if (localKv && !merged.kvGroupId) merged.kvGroupId = localKv;
                    applyStateObject({
                        matched: merged,
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
            if (!localObj) return;
            applyStateObject(localObj);
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
            void runSyncPreview();
        });
        onClick('slgBtnSyncApply', function () {
            void runSyncApply();
        });
        onClick('slgBtnSyncCancel', function () {
            hideSyncPreview();
            updateMismatchUi();
        });
        onClick('slgBtnShowSyncLog', function () {
            const ov = document.getElementById('slgSyncOverflow');
            if (ov) ov.open = false;
            const G = window.ms365GroupDetail;
            if (G && typeof G.setSyncLogVisible === 'function') G.setSyncLogVisible(true);
            else {
                const wrap = document.getElementById('slgSyncLogWrap');
                if (wrap) wrap.hidden = false;
            }
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
            void (async function () {
                const ok = await dlgConfirm(
                    'Verknüpfung und Entwurf dieses Werkzeugs zurücksetzen?\n\nStammdaten und SharePoint bleiben unberührt.',
                    { title: 'Zwischenstand zurücksetzen', okText: 'Zurücksetzen', danger: true }
                );
                if (!ok) return;
                clearStorage();
            })();
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
