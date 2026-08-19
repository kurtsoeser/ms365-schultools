(function () {
    'use strict';

    const STORAGE_KEY = 'ms365-schueler-lehrer-gruppen-v2';

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

    /** @type {'schueler' | 'lehrer'} */
    let activeKind = 'schueler';

    /** @type {{ schuelerGroupId: string|null, lehrerGroupId: string|null }} */
    let matched = { schuelerGroupId: null, lehrerGroupId: null };

    /** @type {{ students: string[], teachers: string[], direktion: string[] }} */
    let listCache = { students: [], teachers: [], direktion: [] };

    /** Graph-Mitgliederzahl je Gruppen-ID; fehlt der Eintrag, ist die Zahl noch unbekannt. */
    /** @type {Record<string, number>} */
    let graphMemberCounts = {};
    let countsFetchGen = 0;

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
        } else if (typeof window.ms365ShowToast === 'function') {
            window.ms365ShowToast(msg);
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

    function collectEmails(arr) {
        const out = [];
        const seen = new Set();
        (Array.isArray(arr) ? arr : []).forEach(function (row) {
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1) return;
            if (seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function readLists() {
        const settings = loadTenantSettings();
        listCache.students = collectEmails(settings && settings.students);
        listCache.teachers = collectEmails(settings && settings.teachers);
        listCache.direktion = collectDirektionOwnerEmails(settings);
    }

    function graphCountFor(groupId) {
        const id = String(groupId || '').trim();
        if (!id) return null;
        const n = graphMemberCounts[id];
        return typeof n === 'number' && n >= 0 ? n : null;
    }

    function paintKindCounts(kind) {
        const isLehrer = kind === 'lehrer';
        const listN = isLehrer ? listCache.teachers.length : listCache.students.length;
        const gid = isLehrer ? matched.lehrerGroupId : matched.schuelerGroupId;
        const groupN = graphCountFor(gid);
        const wrap = document.getElementById(isLehrer ? 'slgLehrerCounts' : 'slgSchuelerCounts');
        const listEl = document.getElementById(isLehrer ? 'slgLehrerCount' : 'slgSchuelerCount');
        const groupEl = document.getElementById(isLehrer ? 'slgLehrerGroupCount' : 'slgSchuelerGroupCount');
        if (listEl) listEl.textContent = String(listN);
        if (groupEl) groupEl.textContent = gid ? (groupN === null ? '–' : String(groupN)) : '–';
        if (!wrap) return;
        wrap.classList.remove('is-match', 'is-mismatch');
        const known = gid && groupN !== null;
        if (known) {
            const same = listN === groupN;
            wrap.classList.add(same ? 'is-match' : 'is-mismatch');
            wrap.title = same
                ? 'Schul-Liste und Gruppenmitglieder: je ' + listN + ' – Anzahl stimmt überein.'
                : 'Schul-Liste: ' + listN + ' E-Mails · Gruppe: ' + groupN + ' Mitglieder. Die Anzahlen unterscheiden sich.';
            wrap.setAttribute(
                'aria-label',
                (isLehrer ? 'Lehrer:innen' : 'Schüler:innen') +
                    ': Liste ' +
                    listN +
                    ', Gruppe ' +
                    groupN +
                    (same ? ', gleich' : ', abweichend')
            );
        } else {
            wrap.title = gid
                ? 'Schul-Liste: ' + listN + ' E-Mails. Mitgliederzahl der gematchten Gruppe wird aus Microsoft Graph geladen.'
                : 'Schul-Liste: ' + listN + ' E-Mails. Noch keine Microsoft-365-Gruppe gematcht.';
            wrap.setAttribute(
                'aria-label',
                (isLehrer ? 'Lehrer:innen' : 'Schüler:innen') + ': Liste ' + listN + ', Gruppe unbekannt'
            );
        }
    }

    /** @type {{ kind: 'schueler'|'lehrer', gid: string, diff: object, graphByEmail: Map<string, object> }|null} */
    let deviationReviewState = null;

    function mr() {
        const M = window.ms365MembershipReconcile;
        if (!M) throw new Error('membership-reconcile.js muss vor diesem Skript geladen werden.');
        return M;
    }

    function logMembershipAction(action, target, summary, result) {
        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
            window.ms365ActionLog.append({
                tool: 'slg',
                action: action,
                target: target,
                summary: summary,
                result: result || 'ok'
            });
        }
    }

    async function openLocalImportWizard(kind, emails) {
        const Ui = window.ms365MembershipImportUi;
        if (!Ui || typeof Ui.openMembershipImportDialog !== 'function') {
            throw new Error('membership-import-ui.js muss vor diesem Skript geladen werden.');
        }
        return Ui.openMembershipImportDialog({
            kind: kind,
            emails: emails,
            getGraphToken: getGraphToken,
            loadSettings: function () {
                return typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
            },
            saveSettings: function (settings) {
                if (typeof window.ms365TenantSettingsSave === 'function') {
                    window.ms365TenantSettingsSave(settings);
                }
            },
            toast: toast,
            dlgConfirm: dlgConfirm,
            logAction: function (entry) {
                logMembershipAction(entry.action, entry.target, entry.summary, entry.result);
            },
            onApplied: async function () {
                readLists();
                updateLeftListUi();
                renderMemberPreview();
                if (deviationReviewState) await loadMembershipReview(deviationReviewState.kind);
            }
        });
    }

    function kindHasMismatch(kind) {
        const isLehrer = kind === 'lehrer';
        const listN = isLehrer ? listCache.teachers.length : listCache.students.length;
        const gid = isLehrer ? matched.lehrerGroupId : matched.schuelerGroupId;
        const groupN = graphCountFor(gid);
        return !!(gid && groupN !== null && listN !== groupN);
    }

    function getSelectedEmails(sectionKey) {
        const out = [];
        document.querySelectorAll('input[data-mr-section="' + sectionKey + '"]:checked').forEach(function (cb) {
            const em = normEmail(cb.getAttribute('data-mr-email'));
            if (em.indexOf('@') !== -1) out.push(em);
        });
        return out;
    }

    function renderDeviationReviewPanel(state) {
        const body = document.getElementById('slgDeviationBody');
        const summaryEl = document.getElementById('slgDeviationSummary');
        const actions = document.getElementById('slgDeviationActions');
        const R = window.ms365MembershipReviewRender;
        if (!body || !summaryEl || !state || !R) return;

        summaryEl.textContent =
            'Markieren Sie Personen und wählen Sie unten in der passenden Spalte die Aktion.';
        if (actions) R.collectMembershipReviewActionButtons(actions);

        body.replaceChildren();
        body.appendChild(
            R.buildMembershipReviewBody({
                diff: state.diff,
                graphByEmail: state.graphByEmail,
                listCount: state.kind === 'lehrer' ? listCache.teachers.length : listCache.students.length,
                labels: {
                    onlyGraphHint:
                        'In der Gruppe online, aber nicht in den lokalen Stammdaten. Import mit Lizenzprüfung und bearbeitbarer Vorschau.'
                }
            })
        );
        R.attachMembershipReviewSectionActions(body, {
            onlyLocal: state.diff.onlyLocal.length,
            onlyGraph: state.diff.onlyGraph.length
        });
    }

    function hideDeviationPanel() {
        const panel = document.getElementById('slgDeviationPanel');
        if (panel) panel.hidden = true;
        deviationReviewState = null;
    }

    async function loadMembershipReview(kind) {
        const isLehrer = kind === 'lehrer';
        const label = isLehrer ? 'Lehrer:innen' : 'Schüler:innen';
        const gid = isLehrer ? matched.lehrerGroupId : matched.schuelerGroupId;
        const localEmails = isLehrer ? listCache.teachers : listCache.students;
        const panel = document.getElementById('slgDeviationPanel');
        const body = document.getElementById('slgDeviationBody');
        const titleEl = document.getElementById('slgDeviationTitle');
        const summaryEl = document.getElementById('slgDeviationSummary');
        const actions = document.getElementById('slgDeviationActions');
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        if (!panel || !body || !titleEl || !summaryEl) return;

        titleEl.textContent = 'Mitglieder-Abgleich: ' + label;
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
            const diff = mr().diffMemberships(localEmails, graphEmails);
            deviationReviewState = {
                kind: kind,
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
            body.appendChild(p);
        }
    }

    async function applyMrAddToGroup() {
        if (!deviationReviewState) return;
        const emails = getSelectedEmails('onlyLocal');
        if (!emails.length) {
            toast('Keine Einträge aus „Nur in der Schul-Liste“ ausgewählt.');
            return;
        }
        const ok = await dlgConfirm(
            emails.length +
                (emails.length === 1 ? ' Person in' : ' Personen in') +
                ' die Microsoft-365-Gruppe aufnehmen?',
            { title: 'In Gruppe aufnehmen' }
        );
        if (!ok) return;
        try {
            const token = await getGraphToken();
            const label = deviationReviewState.kind === 'lehrer' ? 'Lehrer' : 'Schüler';
            const r = await gug().syncEmailsToGroup(token, deviationReviewState.gid, emails, label, appendSyncLog);
            logMembershipAction(
                'membership-add',
                deviationReviewState.gid,
                label + ': +' + r.ok + ' in Gruppe (' + emails.length + ' ausgewählt)'
            );
            live().invalidateMembership();
            await live().loadMembers();
            await refreshGraphMemberCounts();
            readLists();
            toast(r.ok + ' in Gruppe aufgenommen.');
            await loadMembershipReview(deviationReviewState.kind);
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function applyMrRemoveFromGroup() {
        if (!deviationReviewState) return;
        const emails = getSelectedEmails('onlyGraph');
        if (!emails.length) {
            toast('Keine Einträge aus „Nur in der Microsoft-365-Gruppe“ ausgewählt.');
            return;
        }
        const ok = await dlgConfirm(
            emails.length +
                (emails.length === 1 ? ' Person aus' : ' Personen aus') +
                ' der Microsoft-365-Gruppe entfernen?',
            { title: 'Aus Gruppe entfernen', danger: true }
        );
        if (!ok) return;
        try {
            const token = await getGraphToken();
            const label = deviationReviewState.kind === 'lehrer' ? 'Lehrer' : 'Schüler';
            const r = await gug().removeEmailsFromGroup(token, deviationReviewState.gid, emails, label, appendSyncLog);
            logMembershipAction(
                'membership-remove',
                deviationReviewState.gid,
                label + ': −' + r.ok + ' aus Gruppe (' + emails.length + ' ausgewählt)'
            );
            live().invalidateMembership();
            await live().loadMembers();
            await refreshGraphMemberCounts();
            toast(r.ok + ' aus Gruppe entfernt.');
            await loadMembershipReview(deviationReviewState.kind);
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    async function applyMrImportLocal() {
        if (!deviationReviewState) {
            toast('Abgleich ist nicht aktiv – bitte zuerst „Abgleich öffnen“.');
            return;
        }
        const emails = getSelectedEmails('onlyGraph');
        if (!emails.length) {
            toast('Keine Einträge aus „Nur in der Microsoft-365-Gruppe“ ausgewählt.');
            return;
        }
        try {
            const result = await openLocalImportWizard(deviationReviewState.kind, emails);
            if (result && !result.cancelled && result.added > 0) {
                await refreshGraphMemberCounts();
            }
        } catch (e) {
            toast('Fehler: ' + (e.message || e));
        }
    }

    function updateMismatchBar() {
        const bar = document.getElementById('slgMismatchBar');
        const actions = document.getElementById('slgMismatchActions');
        if (!bar || !actions) return;

        const kinds = [];
        if (kindHasMismatch('schueler')) kinds.push('schueler');
        if (kindHasMismatch('lehrer')) kinds.push('lehrer');

        if (!kinds.length) {
            bar.hidden = true;
            actions.replaceChildren();
            return;
        }

        bar.hidden = false;
        actions.replaceChildren();
        kinds.forEach(function (kind) {
            const isLehrer = kind === 'lehrer';
            const label = isLehrer ? 'Lehrer:innen' : 'Schüler:innen';
            const listN = isLehrer ? listCache.teachers.length : listCache.students.length;
            const groupN = graphCountFor(isLehrer ? matched.lehrerGroupId : matched.schuelerGroupId);
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.className = 'btn btn-sm slg-mismatch-bar__btn';
            btn.innerHTML =
                '<i class="bi bi-intersect"></i>' +
                label +
                ': Liste ' +
                listN +
                ' / Gruppe ' +
                groupN +
                ' – Abgleich öffnen';
            btn.addEventListener('click', function (ev) {
                ev.preventDefault();
                ev.stopPropagation();
                void loadMembershipReview(kind);
            });
            actions.appendChild(btn);
        });
    }

    function updateLeftListUi() {
        paintKindCounts('schueler');
        paintKindCounts('lehrer');
        updateMismatchBar();

        const sLine = document.getElementById('slgSchuelerLine');
        const tLine = document.getElementById('slgLehrerLine');
        if (sLine) sLine.textContent = matched.schuelerGroupId ? 'Gematcht: ' + matched.schuelerGroupId : 'Noch kein Match';
        if (tLine) tLine.textContent = matched.lehrerGroupId ? 'Gematcht: ' + matched.lehrerGroupId : 'Noch kein Match';
        updateListPhotoThumbs();
    }

    function mountListPhotoThumb(btn, groupId, displayName) {
        if (!btn) return;
        const existing = btn.querySelector('[data-slg-list-thumb]');
        if (existing) existing.remove();
        const gid = String(groupId || '').trim();
        if (!gid) return;
        const T = window.ms365GroupPhotoThumb;
        if (!T || typeof T.createThumb !== 'function') return;
        const thumb = T.createThumb({
            groupId: gid,
            displayName: String(displayName || '').trim(),
            size: 'list'
        });
        thumb.setAttribute('data-slg-list-thumb', '1');
        const main = btn.querySelector('.slg-side-main');
        if (main) btn.insertBefore(thumb, main);
        else btn.insertBefore(thumb, btn.firstChild);
    }

    function updateListPhotoThumbs() {
        mountListPhotoThumb(
            document.querySelector('#slgListItems button[data-slg-kind="schueler"]'),
            matched.schuelerGroupId,
            'Schüler:innen'
        );
        mountListPhotoThumb(
            document.querySelector('#slgListItems button[data-slg-kind="lehrer"]'),
            matched.lehrerGroupId,
            'Lehrer:innen'
        );
        const host = document.getElementById('slgListItems');
        if (host && window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.hydrate === 'function') {
            window.ms365GroupPhotoThumb.hydrate(host);
        }
    }

    function rememberGraphMemberCount(groupId, count) {
        const id = String(groupId || '').trim();
        if (!id) return;
        const n = typeof count === 'number' ? count : -1;
        if (n < 0) return;
        graphMemberCounts[id] = n;
        updateLeftListUi();
    }

    async function refreshGraphMemberCounts() {
        const ids = [];
        if (matched.schuelerGroupId) ids.push(String(matched.schuelerGroupId));
        if (matched.lehrerGroupId) ids.push(String(matched.lehrerGroupId));
        if (!ids.length) {
            updateLeftListUi();
            return;
        }
        const gen = ++countsFetchGen;
        try {
            const token = await getGraphToken();
            if (gen !== countsFetchGen) return;
            await Promise.all(
                ids.map(async function (id) {
                    try {
                        const n = await gug().fetchGroupMemberCount(token, id);
                        if (typeof n === 'number' && n >= 0) graphMemberCounts[id] = n;
                    } catch {
                        /* Zahl bleibt unbekannt */
                    }
                })
            );
            if (gen !== countsFetchGen) return;
            updateLeftListUi();
        } catch {
            /* Anmeldung abgebrochen oder Graph nicht erreichbar */
        }
    }

    function renderOwnerPreview() {
        const el = document.getElementById('slgOwnerPreview');
        if (!el) return;
        el.replaceChildren();
        const list = listCache.direktion || [];
        if (!list.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine Direktion‑Besitzer in den Schul‑Einstellungen gefunden.';
            el.appendChild(p);
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
        const list = activeKind === 'schueler' ? listCache.students : listCache.teachers;
        const first = list.slice(0, 30);
        if (!first.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = 'Keine E‑Mails in dieser Liste.';
            el.appendChild(p);
            return;
        }
        first.forEach(function (em) {
            const d = document.createElement('div');
            d.textContent = em;
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        if (list.length > first.length) {
            const more = document.createElement('div');
            more.className = 'muted';
            more.style.paddingTop = '8px';
            more.textContent = '… und ' + String(list.length - first.length) + ' weitere.';
            el.appendChild(more);
        }
    }

    function applyCreateDefaults() {
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        if (activeKind === 'schueler') {
            if (dn && !dn.value) dn.value = 'Schüler:innen';
            if (nn && !nn.value) nn.value = 'schueler';
            if (desc && !desc.value) desc.value = 'Alle Schüler:innen (MS365-Schulverwaltung / Schul‑Liste)';
        } else {
            if (dn && !dn.value) dn.value = 'Lehrer:innen';
            if (nn && !nn.value) nn.value = 'lehrer';
            if (desc && !desc.value) desc.value = 'Alle Lehrer:innen (MS365-Schulverwaltung / Schul‑Liste)';
        }
    }

    function getActiveMatchedId() {
        return activeKind === 'schueler' ? matched.schuelerGroupId : matched.lehrerGroupId;
    }

    function setActiveMatchedId(id) {
        if (activeKind === 'schueler') matched.schuelerGroupId = id;
        else matched.lehrerGroupId = id;
        live().resetCaches();
    }

    function setActiveKind(kind) {
        activeKind = kind === 'lehrer' ? 'lehrer' : 'schueler';
        const title = document.getElementById('slgDetailTitle');
        if (title) title.textContent = activeKind === 'schueler' ? 'Schüler:innen' : 'Lehrer:innen';

        document.querySelectorAll('button[data-slg-kind]').forEach(function (btn) {
            const on = btn.getAttribute('data-slg-kind') === activeKind;
            btn.setAttribute('aria-current', on ? 'true' : 'false');
        });

        applyCreateDefaults();
        live().resetCaches();
        gd().clearSearchResults();
        const gid = getActiveMatchedId();
        live().setMatchedMode(!!gid);
        live().fillForm(gid ? { id: gid } : null);
        updateLeftListUi();
        gd().setTab('general');
    }

    async function runSyncMembers() {
        const gid = getActiveMatchedId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const emails = activeKind === 'schueler' ? listCache.students : listCache.teachers;
        if (!emails.length) {
            toast('Keine E‑Mails in dieser Liste.');
            return;
        }
        clearSyncLog();
        appendSyncLog(
            'Start: ' + (activeKind === 'schueler' ? 'Schüler:innen' : 'Lehrer:innen') + ' (' + emails.length + ' Adressen) …',
            ''
        );
        try {
            const token = await getGraphToken();
            const label = activeKind === 'schueler' ? 'Schüler' : 'Lehrer';
            const lc = window.ms365StudentClassLifecycle;
            let joinEmails = emails;
            let leaveEmails = [];
            if (lc && typeof lc.reconcileSammelgruppe === 'function' && typeof gug().fetchGroupMembers === 'function') {
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
                const rec = lc.reconcileSammelgruppe(emails, current);
                joinEmails = rec.join;
                leaveEmails = rec.leave;
                appendSyncLog('Abgleich mit Stammliste: +' + joinEmails.length + ' / −' + leaveEmails.length + '.', '');
            }
            if (joinEmails.length) {
                const r = await gug().syncEmailsToGroup(token, gid, joinEmails, label, appendSyncLog);
                appendSyncLog('Aufnehmen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (leaveEmails.length && typeof gug().removeEmailsFromGroup === 'function') {
                const r = await gug().removeEmailsFromGroup(token, gid, leaveEmails, label, appendSyncLog);
                appendSyncLog('Entfernen: ' + r.ok + ' OK, übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
            }
            if (!joinEmails.length && !leaveEmails.length) {
                appendSyncLog('Keine Änderungen gegenüber der Stammliste.', 'ok');
            }
            await ensureOwners(token, gid);
            live().invalidateMembership();
            await live().loadMembers();
            await refreshGraphMemberCounts();
            toast('Synchronisation abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    function buildStateObject() {
        return {
            kind: STORAGE_KEY,
            savedAt: new Date().toISOString(),
            activeKind: activeKind,
            matched: {
                schuelerGroupId: matched.schuelerGroupId,
                lehrerGroupId: matched.lehrerGroupId
            },
            slgNewDisplayName: document.getElementById('slgNewDisplayName')
                ? document.getElementById('slgNewDisplayName').value
                : '',
            slgNewMailNick: document.getElementById('slgNewMailNick')
                ? document.getElementById('slgNewMailNick').value
                : '',
            slgNewDescription: document.getElementById('slgNewDescription')
                ? document.getElementById('slgNewDescription').value
                : '',
            slgNewCreateTeam: document.getElementById('slgNewCreateTeam')
                ? !!document.getElementById('slgNewCreateTeam').checked
                : false
        };
    }

    function applyStateObject(o) {
        if (!o || typeof o !== 'object') return;
        if (o.matched && typeof o.matched === 'object') {
            matched.schuelerGroupId = o.matched.schuelerGroupId ? String(o.matched.schuelerGroupId) : null;
            matched.lehrerGroupId = o.matched.lehrerGroupId ? String(o.matched.lehrerGroupId) : null;
        }
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const dd = document.getElementById('slgNewDescription');
        const ct = document.getElementById('slgNewCreateTeam');
        if (dn && o.slgNewDisplayName !== undefined) dn.value = String(o.slgNewDisplayName || '');
        if (nn && o.slgNewMailNick !== undefined) nn.value = String(o.slgNewMailNick || '');
        if (dd && o.slgNewDescription !== undefined) dd.value = String(o.slgNewDescription || '');
        if (ct && o.slgNewCreateTeam !== undefined) ct.checked = !!o.slgNewCreateTeam;
        setActiveKind(o.activeKind === 'lehrer' ? 'lehrer' : 'schueler');
    }

    function saveState() {
        try {
            const obj = buildStateObject();
            localStorage.setItem(STORAGE_KEY, JSON.stringify(obj));
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: obj.matched,
                    slgDraft: {
                        activeKind: obj.activeKind,
                        slgNewDisplayName: obj.slgNewDisplayName,
                        slgNewMailNick: obj.slgNewMailNick,
                        slgNewDescription: obj.slgNewDescription,
                        slgNewCreateTeam: obj.slgNewCreateTeam
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
                const hasIds = su && su.matched && !!(su.matched.schuelerGroupId || su.matched.lehrerGroupId);
                if (hasIds || !rawLocal) {
                    const d = su.slgDraft || {};
                    applyStateObject({
                        matched: su.matched,
                        activeKind: d.activeKind === 'lehrer' ? 'lehrer' : 'schueler',
                        slgNewDisplayName: d.slgNewDisplayName,
                        slgNewMailNick: d.slgNewMailNick,
                        slgNewDescription: d.slgNewDescription,
                        slgNewCreateTeam: d.slgNewCreateTeam
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
            matched = { schuelerGroupId: null, lehrerGroupId: null };
            graphMemberCounts = {};
            countsFetchGen += 1;
            if (window.ms365AppDataV2 && typeof window.ms365AppDataV2.patchSetup === 'function') {
                window.ms365AppDataV2.patchSetup({
                    matched: { schuelerGroupId: null, lehrerGroupId: null }
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
            title: 'Schüler:innen',
            searchPlaceholder: 'z. B. lehrer oder @schule.at',
            unmatchedCreateHint:
                'Legt eine Microsoft 365‑Gruppe (Unified) an. Optional auch als Team bereitstellen.',
            membersUnmatchedHint:
                'Mitglieder kommen aus der Schul‑Liste. Nach dem Match können Sie live verwalten und die Liste synchronisieren.',
            membersUnmatchedTitle: 'Vorschau Schul‑Liste (erste 30)',
            membersMatchedHint:
                'Live aus Microsoft Graph. „Mitglieder synchronisieren“ gleicht die Gruppe mit der Schul‑Liste ab: fehlende Adressen werden hinzugefügt, Mitglieder die nicht in der Liste stehen werden entfernt.',
            features: { syncMembers: true, membershipReview: true },
            live: {
                toast: toast,
                dlgConfirm: dlgConfirm,
                getGroupId: getActiveMatchedId,
                ensureDirektionOwners: function (token, gid) {
                    if (!(listCache.direktion && listCache.direktion.length)) {
                        throw new Error('Keine Direktion‑Adressen in den Schul‑Einstellungen.');
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
                }
            },
            onTabUnmatched: function (tab) {
                if (tab === 'owners') renderOwnerPreview();
                if (tab === 'members') renderMemberPreview();
            }
        });
    }

    function wire() {
        const listHost = document.getElementById('slgListItems');
        if (listHost) {
            listHost.addEventListener('click', function (ev) {
                const t = ev.target;
                if (!t || !t.closest) return;
                const item = t.closest('button[data-slg-kind]');
                if (!item) return;
                const kind = item.getAttribute('data-slg-kind');
                setActiveKind(kind === 'lehrer' ? 'lehrer' : 'schueler');
                saveState();
                if (getActiveMatchedId()) live().loadGroup({ silent: true });
            });
        }

        onClick('slgBtnReloadLists', function () {
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
            void loadMembershipReview(activeKind);
        });
        onClick('slgMrAddToGroup', function () {
            void applyMrAddToGroup();
        });
        onClick('slgMrImportLocal', function () {
            void applyMrImportLocal();
        });
        onClick('slgMrRemoveFromGroup', function () {
            void applyMrRemoveFromGroup();
        });
        onClick('slgMrReload', function () {
            if (deviationReviewState) void loadMembershipReview(deviationReviewState.kind);
        });
        onClick('slgBtnSaveState', function () {
            saveState();
            toast('Gespeichert.');
        });
        onClick('slgBtnLoadState', function () {
            loadState();
            toast('Geladen.');
            if (getActiveMatchedId()) live().loadGroup({ silent: true });
        });
        onClick('slgBtnClearStorage', function () {
            clearStorage();
        });
        onClick('slgDeviationClose', function () {
            hideDeviationPanel();
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
        }
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
