(function () {
    'use strict';

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

    function dataV2() {
        return window.ms365AppDataV2 || null;
    }

    let activeKey = '';
    let listFilter = '';
    /** @type {'create' | 'edit'} */
    let classModalMode = 'create';
    /** Original-Kürzel beim Bearbeiten (für Umbenennung / Match-Migration). */
    let classEditOriginalCode = '';
    /** Original-Abschlussjahr beim Bearbeiten. */
    let classEditOriginalYear = '';
    /** Original stableMailNickname beim Bearbeiten (Identität beibehalten). */
    let classEditOriginalNick = '';
    /** @type {{ code: string, name: string, year: string, headName?: string, headEmail?: string, stableMailNickname?: string }[]} */
    let classes = [];
    /** @type {{ klasse: string, name: string, email: string }[]} */
    let students = [];
    /** @type {string[]} */
    let direktion = [];
    let schoolYearLabel = '';
    /** @type {Set<string>} */
    let selectedKeys = new Set();
    /** @type {Record<string, number>} */
    let graphMemberCounts = {};
    let countsFetchGen = 0;
    /** @type {ReturnType<import('../../shared/membership-review-ui.js').createMembershipReview>|null} */
    let membershipReview = null;

    function graphCountFor(groupId) {
        const id = String(groupId || '').trim();
        if (!id) return null;
        const n = graphMemberCounts[id];
        return typeof n === 'number' && n >= 0 ? n : null;
    }

    function updateActiveClassCounts() {
        const row = getActiveRow();
        const gid = getActiveGroupId();
        const listN = row ? emailsForClass(row).length : 0;
        const groupN = graphCountFor(gid);
        const wrap = document.getElementById('jgActiveClassCounts');
        const listEl = document.getElementById('jgActiveListCount');
        const groupEl = document.getElementById('jgActiveGroupCount');
        if (wrap) wrap.hidden = !row;
        if (listEl) listEl.textContent = String(listN);
        if (groupEl) groupEl.textContent = gid ? (groupN === null ? '–' : String(groupN)) : '–';
        if (wrap) {
            wrap.classList.remove('is-match', 'is-mismatch');
            const known = gid && groupN !== null;
            if (known) {
                const same = listN === groupN;
                wrap.classList.add(same ? 'is-match' : 'is-mismatch');
                wrap.title = same
                    ? 'Klassenliste und Gruppe: je ' + listN + ' – Anzahl stimmt überein.'
                    : 'Klassenliste: ' + listN + ' · Gruppe: ' + groupN + ' Mitglieder.';
            } else {
                wrap.title = gid
                    ? 'Klassenliste: ' + listN + ' E-Mails. Mitgliederzahl wird aus Microsoft Graph geladen.'
                    : 'Noch keine Microsoft-365-Gruppe gematcht.';
            }
        }
        if (membershipReview) {
            if (gid && groupN !== null && listN !== groupN && row) {
                membershipReview.updateMismatchBar([
                    {
                        key: rowKey(row),
                        label: 'Klasse ' + (row.name || row.code || ''),
                        listN: listN,
                        groupN: groupN,
                        gid: gid
                    }
                ]);
            } else {
                membershipReview.updateMismatchBar([]);
            }
        }
    }

    async function refreshGraphMemberCounts() {
        const gid = getActiveGroupId();
        if (!gid) {
            updateActiveClassCounts();
            return;
        }
        const gen = ++countsFetchGen;
        try {
            const token = await gug().getGraphToken();
            if (gen !== countsFetchGen) return;
            const n = await gug().fetchGroupMemberCount(token, gid);
            if (typeof n === 'number' && n >= 0) graphMemberCounts[gid] = n;
            if (gen !== countsFetchGen) return;
            updateActiveClassCounts();
        } catch {
            updateActiveClassCounts();
        }
    }

    function initMembershipReview() {
        const R = window.ms365MembershipReviewUi;
        if (!R || typeof R.createMembershipReview !== 'function') return;
        membershipReview = R.createMembershipReview({
            mode: 'class',
            tool: 'jg',
            syncLabel: 'Klasse',
            getGraphToken: function () {
                return gug().getGraphToken();
            },
            getGroupId: function () {
                return getActiveGroupId();
            },
            getLocalEmails: function () {
                return emailsForClass(getActiveRow());
            },
            getAllStudentEmails: collectStudentEmails,
            getActiveReviewKey: function () {
                return rowKey(getActiveRow());
            },
            getReviewTitle: function () {
                const row = getActiveRow();
                if (!row) return 'Mitglieder-Abgleich';
                return 'Mitglieder-Abgleich: ' + (row.name || row.code || 'Klasse');
            },
            toast: toast,
            dlgConfirm: dlgConfirm,
            appendSyncLog: appendSyncLog,
            live: {
                invalidateMembership: function () {
                    live().invalidateMembership();
                },
                loadMembers: function () {
                    return live().loadMembers();
                }
            },
            refreshCounts: refreshGraphMemberCounts,
            onAfterChange: async function () {
                readLists();
                renderLeftList();
                renderMemberPreview();
                updateActiveClassCounts();
            },
            openImport: async function (emails) {
                const Ui = window.ms365MembershipImportUi;
                if (!Ui || typeof Ui.openMembershipImportDialog !== 'function') {
                    throw new Error('membership-import-ui.js fehlt.');
                }
                const row = getActiveRow();
                return Ui.openMembershipImportDialog({
                    kind: 'schueler',
                    emails: emails,
                    importOptions: { defaultClass: row && row.code ? row.code : '' },
                    getGraphToken: function () {
                        return gug().getGraphToken();
                    },
                    loadSettings: function () {
                        return typeof window.ms365TenantSettingsLoad === 'function'
                            ? window.ms365TenantSettingsLoad()
                            : null;
                    },
                    saveSettings: function (settings) {
                        if (typeof window.ms365TenantSettingsSave === 'function') {
                            window.ms365TenantSettingsSave(settings);
                        }
                        readLists();
                    },
                    toast: toast,
                    dlgConfirm: dlgConfirm,
                    logAction: function (entry) {
                        if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                            window.ms365ActionLog.append(
                                Object.assign({ tool: 'jg' }, entry || {})
                            );
                        }
                    },
                    onApplied: async function () {
                        renderMemberPreview();
                        updateActiveClassCounts();
                        if (membershipReview && membershipReview.getState()) {
                            await membershipReview.loadReview(rowKey(getActiveRow()));
                        }
                    }
                });
            },
            logAction: function (action, target, summary) {
                if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                    window.ms365ActionLog.append({
                        tool: 'jg',
                        action: action,
                        target: target,
                        summary: summary,
                        result: 'ok'
                    });
                }
            },
            labels: {
                onlyLocalTitle: 'Nur in der Klassenliste',
                onlyLocalHint:
                    'Schüler:innen dieser Klasse in den Stammdaten, aber nicht in der Microsoft-365-Gruppe.',
                onlyGraphTitle: 'Nur in der Microsoft-365-Gruppe',
                onlyGraphHint:
                    'In der Gruppe, aber nicht als Schüler:in dieser Klasse geführt – importieren oder aus Gruppe entfernen.'
            }
        });
        membershipReview.wire();
    }

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
    function normCode(v) {
        return normStr(v).toUpperCase();
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

    const JG_CREATE_TEAM_KEY = 'ms365-class-create-team-v1';

    function getJgCreateTeam() {
        try {
            const raw = localStorage.getItem(JG_CREATE_TEAM_KEY);
            if (raw === '0' || raw === 'false') return false;
            if (raw === '1' || raw === 'true') return true;
        } catch {
            /* ignore */
        }
        return true;
    }

    function saveJgCreateTeam(on) {
        try {
            localStorage.setItem(JG_CREATE_TEAM_KEY, on ? '1' : '0');
        } catch {
            /* ignore */
        }
    }

    function syncJgCreateTeamUi() {
        const on = getJgCreateTeam();
        const bulk = document.getElementById('jgBulkCreateTeam');
        const single = document.getElementById('slgNewCreateTeam');
        if (bulk) bulk.checked = on;
        if (single) single.checked = on;
    }

    function sanitizeNick(raw) {
        return String(raw || '')
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9-]/g, '')
            .replace(/-+/g, '-')
            .replace(/^-|-$/g, '')
            .slice(0, 60);
    }

    function rowKey(row) {
        return normCode(row && row.code) || normStr(row && row.name).toUpperCase();
    }

    function getActiveRow() {
        const key = normCode(activeKey) || normStr(activeKey).toUpperCase();
        for (let i = 0; i < classes.length; i++) {
            if (rowKey(classes[i]) === key) return classes[i];
        }
        return null;
    }

    function listClassTeams() {
        const api = dataV2();
        if (!api || typeof api.getContainer !== 'function') return [];
        const c = api.getContainer();
        const raw = c && c.core && Array.isArray(c.core.classTeams) ? c.core.classTeams : [];
        if (typeof api.normalizeCoreClassTeams === 'function') return api.normalizeCoreClassTeams(raw);
        return raw;
    }

    function deriveNick(row) {
        if (!row) return '';
        if (typeof window.ms365DeriveClassStableMailNickname === 'function') {
            const d = sanitizeNick(
                window.ms365DeriveClassStableMailNickname(row.year || '', row.code || '', row)
            );
            if (d) return d;
        }
        const fromRow = sanitizeNick(row.stableMailNickname);
        if (fromRow) return fromRow;
        const y = normStr(row.year);
        const yy = /^\d{4}$/.test(y) ? y : '';
        const tail = String(normCode(row.code) || '')
            .replace(/[^0-9A-Za-z]/g, '')
            .toLowerCase()
            .slice(0, 24);
        if (yy && tail) return ('jg' + yy + tail).toLowerCase().slice(0, 60);
        if (tail) return ('jg' + tail).toLowerCase().slice(0, 60);
        return '';
    }

    function findClassTeam(row) {
        if (!row) return null;
        const teams = listClassTeams();
        const code = normCode(row.code);
        const year = normStr(row.year);
        if (code) {
            for (let i = 0; i < teams.length; i++) {
                if (normCode(teams[i].classCode) !== code) continue;
                if (year && teams[i].abschlussJahr && String(teams[i].abschlussJahr) !== year) continue;
                return teams[i];
            }
        }
        const nick = deriveNick(row);
        if (nick) {
            for (let i = 0; i < teams.length; i++) {
                if (sanitizeNick(teams[i].stableMailNickname) === nick) return teams[i];
            }
        }
        return null;
    }

    function persistNickForRow(row) {
        const existing = findClassTeam(row);
        // Nur bereits gematchte Gruppen behalten ihren echten Graph-Alias.
        // Ungematchte Klassen nehmen immer das aktuelle Nomenklatur-Schema.
        if (existing && existing.graphGroupId) {
            const pretty = graphMailNick(existing.mailNickname);
            if (pretty) return pretty;
            const stable = sanitizeNick(existing.stableMailNickname);
            if (stable) return stable;
        }
        const derived = deriveNick(row);
        if (derived) return graphMailNick(derived) || derived;
        return sanitizeNick(row && row.stableMailNickname);
    }

    function getActiveGroupId() {
        const row = getActiveRow();
        const team = findClassTeam(row);
        const id = team && team.graphGroupId ? String(team.graphGroupId).trim() : '';
        return id || null;
    }

    function getSchoolDomainNoAt() {
        try {
            if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
                const d = String(window.ms365GetSchoolDomainNoAt() || '')
                    .replace(/^@+/, '')
                    .trim();
                if (d) return d;
            }
        } catch {
            /* ignore */
        }
        try {
            if (typeof window.ms365TenantSettingsLoad === 'function') {
                const s = window.ms365TenantSettingsLoad();
                const d = String((s && s.domain) || '')
                    .replace(/^@+/, '')
                    .trim();
                if (d) return d;
            }
        } catch {
            /* ignore */
        }
        try {
            const api = dataV2();
            const c = api && typeof api.getContainer === 'function' ? api.getContainer() : null;
            return String((c && c.core && c.core.domain) || '')
                .replace(/^@+/, '')
                .trim();
        } catch {
            return '';
        }
    }

    function domainFromMail(mail) {
        const m = String(mail || '')
            .trim()
            .toLowerCase();
        const i = m.lastIndexOf('@');
        return i >= 0 ? m.slice(i + 1) : '';
    }

    function psEscapeSingle(s) {
        return String(s || '').replace(/'/g, "''");
    }

    function collectSmtpScriptItems(onlyActive) {
        const domain = getSchoolDomainNoAt();
        const items = [];
        const rows = onlyActive ? [getActiveRow()].filter(Boolean) : classes.slice();
        rows.forEach(function (row) {
            if (!row) return;
            const team = findClassTeam(row);
            const id = team && team.graphGroupId ? String(team.graphGroupId).trim() : '';
            if (!id) return;
            let nick = '';
            if (onlyActive) {
                const aliasEl = document.getElementById('slgLiveAlias');
                nick = graphMailNick(aliasEl && aliasEl.value);
            }
            if (!nick) {
                nick =
                    graphMailNick(team && team.mailNickname) ||
                    graphMailNick(team && team.stableMailNickname) ||
                    deriveNick(row);
            }
            if (!nick) return;
            items.push({
                id: id,
                name: normStr(row.name) || normStr(row.code) || nick,
                nick: nick,
                smtp: domain ? nick + '@' + domain : ''
            });
        });
        return { domain: domain, items: items };
    }

    function buildClassSmtpPs1(items, domain) {
        const stamp = new Date().toISOString();
        const lines = [];
        lines.push('#Requires -Version 5.1');
        lines.push('# Klassengruppen: primäre SMTP auf die Schul-Domain setzen.');
        lines.push('# Microsoft Graph kann die Domain nicht aendern – dafuer Exchange Online (Set-UnifiedGroup).');
        lines.push('# Erzeugt in der Browser-App am ' + stamp);
        lines.push('# Schul-Domain: ' + domain);
        lines.push('');
        lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
        lines.push('$ErrorActionPreference = "Continue"');
        lines.push('');
        lines.push('if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {');
        lines.push('    Write-Host "Installiere ExchangeOnlineManagement (einmalig) ..." -ForegroundColor Yellow');
        lines.push('    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber');
        lines.push('}');
        lines.push('Import-Module ExchangeOnlineManagement -ErrorAction Stop');
        lines.push('Connect-ExchangeOnline -ShowBanner:$false');
        lines.push('');
        lines.push('$items = @(');
        items.forEach(function (it, idx) {
            const comma = idx < items.length - 1 ? ',' : '';
            lines.push(
                "    [pscustomobject]@{ Id = '" +
                    psEscapeSingle(it.id) +
                    "'; Name = '" +
                    psEscapeSingle(it.name) +
                    "'; Smtp = '" +
                    psEscapeSingle(it.smtp) +
                    "' }" +
                    comma
            );
        });
        lines.push(')');
        lines.push('');
        lines.push('foreach ($r in $items) {');
        lines.push('    try {');
        lines.push('        Set-UnifiedGroup -Identity $r.Id -PrimarySmtpAddress $r.Smtp -ErrorAction Stop');
        lines.push('        Write-Host ("OK  {0} -> {1}" -f $r.Name, $r.Smtp) -ForegroundColor Green');
        lines.push('    } catch {');
        lines.push('        Write-Warning ("Fehler {0}: {1}" -f $r.Name, $_.Exception.Message)');
        lines.push('    }');
        lines.push('}');
        lines.push('');
        lines.push('Write-Host "Fertig. In Entra/Exchange kann die neue Adresse kurz brauchen." -ForegroundColor Cyan');
        return lines.join('\r\n');
    }

    function downloadText(filename, text) {
        const blob = new Blob([text], { type: 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        a.remove();
        setTimeout(function () {
            URL.revokeObjectURL(a.href);
        }, 1500);
    }

    function showSmtpScript(text) {
        const drop = document.getElementById('jgSmtpDrop');
        if (drop) drop.open = true;
        const ta = document.getElementById('jgSmtpScript');
        if (ta) {
            ta.value = text;
            ta.style.display = 'block';
        }
        if (navigator.clipboard && typeof navigator.clipboard.writeText === 'function') {
            navigator.clipboard.writeText(text).then(
                function () {
                    toast('Exchange‑Skript kopiert (und zum Download angeboten).');
                },
                function () {
                    toast('Skript angezeigt – bitte manuell kopieren.');
                }
            );
        } else {
            toast('Skript angezeigt – bitte manuell kopieren.');
        }
    }

    function graphMailNick(raw) {
        const s = String(raw || '').trim();
        if (!s) return '';
        try {
            if (typeof gug().sanitizeUnifiedGroupMailNickname === 'function') {
                return gug().sanitizeUnifiedGroupMailNickname(s);
            }
        } catch {
            /* ignore */
        }
        try {
            const api = dataV2();
            if (api && typeof api.mailNicknamePrefixSanitize === 'function') {
                return api.mailNicknamePrefixSanitize(s, 60);
            }
        } catch {
            /* ignore */
        }
        return sanitizeNick(s);
    }

    function previewMailFromAlias() {
        const mailEl = document.getElementById('slgLiveMail');
        const aliasEl = document.getElementById('slgLiveAlias');
        if (!mailEl || !aliasEl || aliasEl.readOnly) return '';
        const nick = graphMailNick(aliasEl.value);
        const school = getSchoolDomainNoAt();
        const graphMail = String(mailEl.getAttribute('data-graph-mail') || '').trim();
        const domain = school || domainFromMail(graphMail);
        if (!nick || !domain) return '';
        const wanted = nick + '@' + domain;
        mailEl.value = wanted;
        return wanted;
    }

    function refreshSmtpHint() {
        previewMailFromAlias();
        const el = document.getElementById('jgSmtpHint');
        if (!el) return;
        const domain = getSchoolDomainNoAt();
        const mailEl = document.getElementById('slgLiveMail');
        const aliasEl = document.getElementById('slgLiveAlias');
        const actual = String((mailEl && mailEl.getAttribute('data-graph-mail')) || '').trim();
        const nick = graphMailNick(aliasEl && aliasEl.value) || deriveNick(getActiveRow());
        const wanted = nick && domain ? nick + '@' + domain : String((mailEl && mailEl.value) || '').trim();
        const actDom = domainFromMail(actual);
        if (!domain) {
            el.innerHTML =
                'Keine Schul‑Domain gespeichert. Bitte in den <a href="../tenant.html">Stammdaten</a> oder in der Einrichtung (Schritt 2) setzen.';
            return;
        }
        let html =
            'Ziel‑SMTP (Schul‑Domain): <code>' +
            escapeHtml(wanted || '–') +
            '</code><br>Aktuell in Graph: <code>' +
            escapeHtml(actual || '–') +
            '</code>';
        if (wanted && actDom && actDom !== domain.toLowerCase()) {
            html +=
                '<br><span style="color:#856404;">Weicht ab – Graph kann die Domain nicht setzen. Exchange‑Skript unten verwenden (Domain muss in Microsoft 365 verifiziert sein).</span>';
        } else if (wanted && actual && actual.toLowerCase() === wanted.toLowerCase()) {
            html += '<br>Die Gruppenadresse entspricht bereits der Schul‑Domain.';
        }
        el.innerHTML = html;
    }

    async function runSmtpScript(onlyActive) {
        const pack = collectSmtpScriptItems(onlyActive);
        if (!pack.domain) {
            toast('Bitte zuerst die Schul‑Domain in den Stammdaten oder in der Einrichtung (Schritt 2) speichern.');
            return;
        }
        if (!pack.items.length) {
            toast(onlyActive ? 'Diese Klasse ist nicht mit einer Microsoft‑365‑Gruppe gematcht.' : 'Keine gematchten Klassengruppen.');
            return;
        }
        const preview = pack.items
            .slice(0, 8)
            .map(function (it) {
                return it.name + ' → ' + it.smtp;
            })
            .join('\n');
        const extra = pack.items.length > 8 ? '\n… insgesamt ' + pack.items.length + ' Gruppen' : '';
        const ok = await dlgConfirm(
            'Graph kann die E‑Mail‑Domain nicht ändern.\n\nEs wird ein Exchange‑Online‑Skript erzeugt (Set-UnifiedGroup) für:\n\n' +
                preview +
                extra +
                '\n\nDie Domain ' +
                pack.domain +
                ' muss in Microsoft 365 verifiziert sein. Fortfahren?',
            { title: 'SMTP auf Schul‑Domain', okText: 'Skript erzeugen' }
        );
        if (!ok) return;
        const text = buildClassSmtpPs1(pack.items, pack.domain);
        downloadText(onlyActive ? 'klassengruppe-smtp.ps1' : 'klassengruppen-smtp.ps1', text);
        showSmtpScript(text);
    }

    function isDirektionRole(roleRaw) {
        const r = normStr(roleRaw).toLowerCase();
        return !!r && (r.indexOf('direktion') !== -1 || r.indexOf('direktor') !== -1);
    }

    function currentYearFromV2() {
        const api = dataV2();
        if (!api || typeof api.getContainer !== 'function') return '';
        const c = api.getContainer();
        return c && c.years ? String(c.years.current || '').trim() : '';
    }

    function readLists() {
        const settings = typeof window.ms365TenantSettingsLoad === 'function' ? window.ms365TenantSettingsLoad() : null;
        classes = Array.isArray(settings && settings.classes) ? settings.classes.slice() : [];
        students = Array.isArray(settings && settings.students) ? settings.students.slice() : [];
        schoolYearLabel = currentYearFromV2();
        const out = [];
        const seen = new Set();
        const admin = settings && Array.isArray(settings.admin) ? settings.admin : [];
        admin.forEach(function (row) {
            if (!isDirektionRole(row && row.role)) return;
            const em = normEmail(row && row.email);
            if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        direktion = out;
        const yearEl = document.getElementById('jgYearLabel');
        if (yearEl) {
            yearEl.textContent = schoolYearLabel
                ? 'Schuljahr ' +
                  schoolYearLabel +
                  ' – links Klasse wählen bzw. neu anlegen, rechts matchen oder eine Gruppe erstellen.'
                : 'Aktuelles Schuljahr – links Klasse wählen bzw. neu anlegen, rechts matchen oder eine Gruppe erstellen.';
        }
    }

    function dispatchTenantSettingsChanged(saved, reason) {
        try {
            window.dispatchEvent(
                new CustomEvent('ms365-tenant-settings-changed', {
                    detail: { settings: saved, reason: reason || 'jg-class' }
                })
            );
        } catch (_) {
            /* ignore */
        }
    }

    function setClassModalError(msg) {
        const el = document.getElementById('jgClassModalError');
        if (!el) return;
        const text = normStr(msg);
        el.textContent = text;
        el.style.display = text ? '' : 'none';
    }

    function updateClassActionButtons() {
        const has = !!getActiveRow();
        const editBtn = document.getElementById('jgBtnEditClass');
        const delBtn = document.getElementById('jgBtnDeleteClass');
        if (editBtn) editBtn.disabled = !has;
        if (delBtn) delBtn.disabled = !has;
    }

    function classCodeExists(code, exceptCode) {
        const key = normCode(code);
        const skip = normCode(exceptCode);
        return classes.some(function (r) {
            const c = normCode(r.code);
            if (skip && c === skip) return false;
            return c === key;
        });
    }

    function remapStudentKlassen(list, fromCode, toCode) {
        const from = normCode(fromCode);
        const to = normCode(toCode);
        if (!from || from === to) return Array.isArray(list) ? list.slice() : [];
        return (Array.isArray(list) ? list : []).map(function (s) {
            if (normCode(s && s.klasse) !== from) return s;
            return Object.assign({}, s, { klasse: to });
        });
    }

    function openClassModal(mode) {
        const modal = document.getElementById('jgClassModal');
        if (!modal) return;
        classModalMode = mode === 'edit' ? 'edit' : 'create';
        const title = document.getElementById('jgClassModalTitle');
        const hint = document.getElementById('jgClassModalHint');
        const saveBtn = document.getElementById('jgClassModalSave');
        const codeEl = document.getElementById('jgNewCode');
        const nameEl = document.getElementById('jgNewName');
        const yearEl = document.getElementById('jgNewYear');
        const headNameEl = document.getElementById('jgNewHeadName');
        const headEmailEl = document.getElementById('jgNewHeadEmail');
        const row = classModalMode === 'edit' ? getActiveRow() : null;

        if (classModalMode === 'edit' && !row) {
            toast('Bitte zuerst eine Klasse wählen.');
            return;
        }

        classEditOriginalCode = classModalMode === 'edit' ? normCode(row.code) : '';
        classEditOriginalYear = classModalMode === 'edit' ? normStr(row.year) : '';
        classEditOriginalNick = classModalMode === 'edit' ? sanitizeNick(row.stableMailNickname) : '';

        if (title) {
            title.textContent = classModalMode === 'edit' ? 'Klasse bearbeiten' : 'Neue Klasse anlegen';
        }
        if (hint) {
            hint.textContent =
                classModalMode === 'edit'
                    ? 'Änderungen werden in die Stammdaten geschrieben. Bei neuem Kürzel wird ein vorhandenes Match mit umgehängt (Mail‑Nickname bleibt).'
                    : 'Die Klasse wird in die Stammdaten des aktuellen Schuljahrs geschrieben und erscheint danach in der Liste.';
        }
        if (codeEl) codeEl.value = classModalMode === 'edit' && row ? row.code || '' : '';
        if (nameEl) nameEl.value = classModalMode === 'edit' && row ? row.name || '' : '';
        if (yearEl) yearEl.value = classModalMode === 'edit' && row ? row.year || '' : '';
        if (headNameEl) headNameEl.value = classModalMode === 'edit' && row ? row.headName || '' : '';
        if (headEmailEl) headEmailEl.value = classModalMode === 'edit' && row ? row.headEmail || '' : '';
        if (saveBtn) {
            saveBtn.innerHTML =
                classModalMode === 'edit'
                    ? '<i class="bi bi-check-lg"></i>Speichern'
                    : '<i class="bi bi-check-lg"></i>Anlegen';
        }
        setClassModalError('');
        modal.classList.add('open');
        modal.setAttribute('aria-hidden', 'false');
        if (codeEl) {
            setTimeout(function () {
                codeEl.focus();
                try {
                    codeEl.select();
                } catch (_) {
                    /* ignore */
                }
            }, 30);
        }
    }

    function closeClassModal() {
        const modal = document.getElementById('jgClassModal');
        if (!modal) return;
        modal.classList.remove('open');
        modal.setAttribute('aria-hidden', 'true');
        classModalMode = 'create';
        classEditOriginalCode = '';
        classEditOriginalYear = '';
        classEditOriginalNick = '';
        setClassModalError('');
    }

    function persistClassCreate(entry) {
        if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
            throw new Error('Stammdaten nicht verfügbar (tenant-settings-core.js).');
        }
        const current = window.ms365TenantSettingsLoad() || {};
        const next = Object.assign({}, current);
        const list = Array.isArray(current.classes) ? current.classes.slice() : [];
        list.push({
            code: entry.code,
            name: entry.name,
            year: entry.year,
            headName: entry.headName || '',
            headEmail: entry.headEmail || '',
            stableMailNickname: entry.stableMailNickname || ''
        });
        next.classes = list;
        const saved = window.ms365TenantSettingsSave(next);
        dispatchTenantSettingsChanged(saved, 'jg-class-create');
        return saved;
    }

    function persistClassUpdate(originalCode, originalYear, entry) {
        if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
            throw new Error('Stammdaten nicht verfügbar (tenant-settings-core.js).');
        }
        const from = normCode(originalCode);
        const to = normCode(entry.code);
        const current = window.ms365TenantSettingsLoad() || {};
        const next = Object.assign({}, current);
        const list = Array.isArray(current.classes) ? current.classes.slice() : [];
        const idx = list.findIndex(function (r) {
            return normCode(r.code) === from;
        });
        if (idx < 0) throw new Error('Klasse nicht mehr gefunden.');
        const prev = list[idx] || {};
        const nick = sanitizeNick(prev.stableMailNickname) || sanitizeNick(entry.stableMailNickname) || '';
        list[idx] = {
            code: to,
            name: entry.name,
            year: entry.year,
            headName: entry.headName || '',
            headEmail: entry.headEmail || '',
            stableMailNickname: nick
        };
        next.classes = list;
        if (from !== to) {
            next.students = remapStudentKlassen(current.students, from, to);
        }
        const api = dataV2();
        if (api && typeof api.patchClassTeamMeta === 'function') {
            api.patchClassTeamMeta(from, originalYear, {
                classCode: to,
                displayName: entry.name,
                abschlussJahr: entry.year
            });
        }
        const saved = window.ms365TenantSettingsSave(next);
        dispatchTenantSettingsChanged(saved, 'jg-class-update');
        return saved;
    }

    function persistClassDelete(code, year) {
        if (typeof window.ms365TenantSettingsLoad !== 'function' || typeof window.ms365TenantSettingsSave !== 'function') {
            throw new Error('Stammdaten nicht verfügbar (tenant-settings-core.js).');
        }
        const key = normCode(code);
        const current = window.ms365TenantSettingsLoad() || {};
        const next = Object.assign({}, current);
        next.classes = (Array.isArray(current.classes) ? current.classes : []).filter(function (r) {
            return normCode(r.code) !== key;
        });
        const saved = window.ms365TenantSettingsSave(next);
        const api = dataV2();
        if (api && typeof api.removeClassTeamByClassCode === 'function') {
            api.removeClassTeamByClassCode(key, year);
        }
        dispatchTenantSettingsChanged(saved, 'jg-class-delete');
        return saved;
    }

    async function offerCreateM365Group(code) {
        const ok = await dlgConfirm(
            'Klasse „' + code + '“ ist in den Stammdaten. Jetzt auch eine Microsoft‑365‑Gruppe anlegen und matchen?',
            {
                title: 'M365‑Gruppe anlegen?',
                okText: 'Ja, anlegen',
                cancelText: 'Später'
            }
        );
        if (!ok) return;
        const btn = document.getElementById('slgBtnCreate');
        if (btn) {
            btn.click();
            return;
        }
        toast('Bitte rechts unter „Neue Gruppe anlegen“ auf „Anlegen & matchen“ klicken.');
    }

    async function submitClassModal() {
        const codeEl = document.getElementById('jgNewCode');
        const nameEl = document.getElementById('jgNewName');
        const yearEl = document.getElementById('jgNewYear');
        const headNameEl = document.getElementById('jgNewHeadName');
        const headEmailEl = document.getElementById('jgNewHeadEmail');
        const code = normCode(codeEl && codeEl.value);
        const name = normStr(nameEl && nameEl.value);
        const year = normStr(yearEl && yearEl.value);
        const headName = normStr(headNameEl && headNameEl.value);
        const headEmail = normEmail(headEmailEl && headEmailEl.value);
        const editing = classModalMode === 'edit';
        const originalCode = classEditOriginalCode;
        const originalYear = classEditOriginalYear;

        if (!code) {
            setClassModalError('Bitte ein Kürzel eingeben.');
            if (codeEl) codeEl.focus();
            return;
        }
        if (!name) {
            setClassModalError('Bitte einen Namen eingeben.');
            if (nameEl) nameEl.focus();
            return;
        }
        if (!/^\d{4}$/.test(year)) {
            setClassModalError('Bitte ein gültiges Abschlussjahr (4 Ziffern) eingeben.');
            if (yearEl) yearEl.focus();
            return;
        }
        if (headEmail && headEmail.indexOf('@') === -1) {
            setClassModalError('KV‑E‑Mail sieht ungültig aus.');
            if (headEmailEl) headEmailEl.focus();
            return;
        }
        if (classCodeExists(code, editing ? originalCode : '')) {
            setClassModalError('Klasse mit Kürzel „' + code + '“ gibt es bereits.');
            if (codeEl) codeEl.focus();
            return;
        }

        let nick = editing ? classEditOriginalNick : '';
        if (!nick && typeof window.ms365DeriveClassStableMailNickname === 'function') {
            nick = sanitizeNick(
                window.ms365DeriveClassStableMailNickname(year, code, {
                    code: code,
                    year: year,
                    headName: headName,
                    headEmail: headEmail
                })
            );
        }
        if (!nick) nick = sanitizeNick('jg' + year + code);

        const entry = {
            code: code,
            name: name,
            year: year,
            headName: headName,
            headEmail: headEmail,
            stableMailNickname: nick
        };

        try {
            if (editing) persistClassUpdate(originalCode, originalYear, entry);
            else persistClassCreate(entry);
        } catch (e) {
            setClassModalError((e && e.message) || String(e));
            return;
        }

        closeClassModal();
        readLists();
        setActiveKey(code);
        toast('Klasse „' + code + '“ ' + (editing ? 'gespeichert.' : 'angelegt.'));
        if (!editing) await offerCreateM365Group(code);
    }

    async function deleteActiveClass() {
        const row = getActiveRow();
        if (!row) {
            toast('Bitte zuerst eine Klasse wählen.');
            return;
        }
        const code = normCode(row.code);
        const team = findClassTeam(row);
        const matched = !!(team && team.graphGroupId);
        let msg =
            'Klasse „' +
            (row.name || code) +
            '“ (' +
            code +
            ') aus den Stammdaten entfernen?';
        if (matched) {
            msg +=
                '\n\nDie Verknüpfung zur Microsoft‑365‑Gruppe wird gelöst. Die Gruppe selbst bleibt in Entra erhalten.';
        }
        msg +=
            '\n\nSchüler:innen mit dieser Klassenkennung bleiben in der Schülerliste (Zuordnung ggf. manuell anpassen).';
        const ok = await dlgConfirm(msg, {
            title: 'Klasse löschen?',
            okText: 'Löschen',
            cancelText: 'Abbrechen',
            danger: true
        });
        if (!ok) return;
        try {
            persistClassDelete(code, row.year);
        } catch (e) {
            toast((e && e.message) || String(e));
            return;
        }
        selectedKeys.delete(rowKey(row));
        readLists();
        ensureActiveKey();
        renderLeftList();
        applyCreateDefaults();
        refreshMatchUi();
        updateClassActionButtons();
        toast('Klasse „' + code + '“ gelöscht.');
    }

    function studentsForClass(row) {
        if (!row) return [];
        const code = normCode(row.code);
        const name = normStr(row.name).toLowerCase();
        return students.filter(function (s) {
            const k = normStr(s && s.klasse);
            if (!k) return false;
            if (code && normCode(k) === code) return true;
            if (name && k.toLowerCase() === name) return true;
            return false;
        });
    }

    function emailsForClass(row) {
        const seen = new Set();
        const out = [];
        studentsForClass(row).forEach(function (s) {
            const em = normEmail(s && s.email);
            if (!em || em.indexOf('@') === -1 || seen.has(em)) return;
            seen.add(em);
            out.push(em);
        });
        return out;
    }

    function getOwnerOptions() {
        const useKV = !document.getElementById('jgOwnerUseKV') || document.getElementById('jgOwnerUseKV').checked;
        const useDirektion = document.getElementById('jgOwnerUseDirektion') && document.getElementById('jgOwnerUseDirektion').checked;
        return { useKV, useDirektion };
    }

    /** Liefert die Besitzer-E-Mails für eine Klasse laut aktueller Einstellung. */
    function ownersForRow(row) {
        const opts = getOwnerOptions();
        const seen = new Set();
        const out = [];
        function add(em) {
            const e = normEmail(em);
            if (!e || e.indexOf('@') === -1 || seen.has(e)) return;
            seen.add(e);
            out.push(e);
        }
        if (opts.useKV && row && row.headEmail) add(row.headEmail);
        if (opts.useDirektion) direktion.forEach(add);
        return out;
    }

    /** Globale Besitzer (ohne Klassenbezug), z. B. für bestehende gematchte Gruppen. */
    function ownersGlobal() {
        const opts = getOwnerOptions();
        if (opts.useDirektion) return direktion.slice();
        return [];
    }

    function renderOwnerPreview() {
        const el = document.getElementById('slgOwnerPreview');
        if (!el) return;
        el.replaceChildren();
        const row = getActiveRow();
        const owners = ownersForRow(row);
        if (!owners.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent = row && row.headEmail
                ? 'Keine Besitzer ausgewählt (Optionen unten in den Sammelaktionen prüfen).'
                : 'Kein Klassenvorstand in den Stammdaten hinterlegt und keine Direktion ausgewählt.';
            el.appendChild(p);
            return;
        }
        owners.forEach(function (em) {
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
        const row = getActiveRow();
        const list = studentsForClass(row);
        const emails = emailsForClass(row);
        if (!list.length) {
            const p = document.createElement('p');
            p.style.margin = '0';
            p.style.color = '#6c757d';
            p.textContent =
                'Keine Schüler:innen dieser Klasse in den Stammdaten. Nach dem Match Personen über Suche hinzufügen.';
            el.appendChild(p);
            return;
        }
        const first = list.slice(0, 30);
        first.forEach(function (s) {
            const d = document.createElement('div');
            const parts = [];
            if (s.name) parts.push(s.name);
            if (s.email) parts.push(s.email);
            if (!parts.length) parts.push(s.klasse || '–');
            d.textContent = parts.join(' · ');
            d.style.padding = '4px 0';
            d.style.borderBottom = '1px solid #eef1f4';
            el.appendChild(d);
        });
        const more = document.createElement('div');
        more.className = 'muted';
        more.style.paddingTop = '8px';
        more.textContent =
            String(list.length) +
            ' Einträge · ' +
            String(emails.length) +
            ' mit E‑Mail' +
            (list.length > first.length ? ' · Anzeige der ersten 30' : '') +
            '.';
        el.appendChild(more);
    }

    function applyCreateDefaults() {
        const row = getActiveRow();
        const code = row ? row.code : activeKey;
        const name = row && row.name ? row.name : code;
        const nick = persistNickForRow(row);
        const dn = document.getElementById('slgNewDisplayName');
        const nn = document.getElementById('slgNewMailNick');
        const desc = document.getElementById('slgNewDescription');
        if (dn) dn.value = name || ('Klasse ' + (code || ''));
        if (nn) nn.value = nick;
        if (desc) {
            desc.value =
                'Jahrgangsgruppe ' +
                (name || code || '') +
                (row && row.year ? ' / Abschluss ' + row.year : '') +
                ' (MS365-Schulverwaltung)';
        }
        const search = document.getElementById('slgGroupSearch');
        if (search) search.value = nick || name || code || '';
        syncJgCreateTeamUi();
    }

    function refreshMatchUi() {
        const gid = getActiveGroupId();
        const row = getActiveRow();
        const title = document.getElementById('slgDetailTitle');
        if (title) {
            if (!row) title.textContent = 'Jahrgangsgruppe';
            else {
                const bits = [];
                bits.push(row.name || row.code || 'Klasse');
                if (row.code && row.name && row.name !== row.code) bits[0] = row.name + ' (' + row.code + ')';
                title.textContent = bits[0];
            }
        }
        live().resetCaches();
        live().setMatchedMode(!!gid);
        live().fillForm(gid ? { id: gid } : null);
        renderOwnerPreview();
        renderMemberPreview();
        refreshSmtpHint();
        updateActiveClassCounts();
        void refreshGraphMemberCounts();
    }

    function renderLeftList() {
        const host = document.getElementById('jgListItems');
        const summary = document.getElementById('jgListSummary');
        const empty = document.getElementById('jgEmptyHint');
        const wrap = document.getElementById('jgDetailWrap');
        if (!host) return;
        host.replaceChildren();
        const q = listFilter.toLowerCase();
        const all = classes;
        const list = all.filter(function (row) {
            if (!q) return true;
            const nick = persistNickForRow(row);
            const hay = (row.code + ' ' + (row.name || '') + ' ' + (row.year || '') + ' ' + nick).toLowerCase();
            return hay.indexOf(q) !== -1;
        });
        let matchedN = 0;
        all.forEach(function (row) {
            const team = findClassTeam(row);
            if (team && team.graphGroupId) matchedN++;
        });
        if (summary) {
            summary.textContent =
                String(all.length) +
                ' Klassen' +
                (schoolYearLabel ? ' (' + schoolYearLabel + ')' : '') +
                ' · ' +
                String(matchedN) +
                ' gematcht' +
                (q ? ' · Filter: ' + String(list.length) : '');
        }
        const hasRows = all.length > 0;
        if (empty) empty.style.display = hasRows ? 'none' : '';
        if (wrap) wrap.style.display = hasRows ? '' : 'none';

        if (!list.length) {
            const li = document.createElement('li');
            const p = document.createElement('p');
            p.className = 'muted';
            p.style.margin = '10px 12px';
            p.textContent = hasRows ? 'Keine Treffer im Filter.' : 'Liste ist leer.';
            li.appendChild(p);
            host.appendChild(li);
            updateBulkCount();
            updateClassActionButtons();
            return;
        }

        list.forEach(function (row) {
            const team = findClassTeam(row);
            const gid = team && team.graphGroupId ? String(team.graphGroupId) : '';
            const nick = persistNickForRow(row);
            const li = document.createElement('li');
            const btn = document.createElement('button');
            btn.type = 'button';
            btn.setAttribute('data-jg-code', rowKey(row));
            if (rowKey(row) === (normCode(activeKey) || normStr(activeKey).toUpperCase())) {
                btn.setAttribute('aria-current', 'true');
            }
            const main = document.createElement('span');
            main.className = 'slg-side-main';
            const t = document.createElement('span');
            t.className = 'slg-side-title';
            t.textContent = (row.name || row.code) + (row.name && row.code && row.name !== row.code ? ' (' + row.code + ')' : '');
            const meta = document.createElement('span');
            meta.className = 'muted slg-side-meta';
            const prefix = [];
            if (row.year) prefix.push('Abschluss ' + row.year);
            if (nick) prefix.push(nick);
            if (prefix.length) meta.appendChild(document.createTextNode(prefix.join(' · ') + ' · '));
            const badge = document.createElement('span');
            badge.className = 'jg-match-badge ' + (gid ? 'is-ok' : 'is-warn');
            const ico = document.createElement('i');
            ico.className = gid ? 'bi bi-check-circle-fill' : 'bi bi-exclamation-circle-fill';
            ico.setAttribute('aria-hidden', 'true');
            badge.appendChild(ico);
            badge.appendChild(document.createTextNode(gid ? 'Gematcht' : 'Kein Match'));
            meta.appendChild(badge);
            btn.classList.add(gid ? 'is-matched' : 'is-unmatched');
            if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.createThumb === 'function') {
                btn.insertBefore(
                    window.ms365GroupPhotoThumb.createThumb({
                        groupId: gid,
                        displayName: (row.name || row.code || '').trim(),
                        size: 'list'
                    }),
                    btn.firstChild
                );
            }
            main.appendChild(t);
            main.appendChild(meta);
            btn.appendChild(main);
            const pick = document.createElement('label');
            pick.className = 'jg-pick';
            pick.title = 'Für Sammelaktion auswählen';
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.setAttribute('data-jg-pick', rowKey(row));
            cb.checked = selectedKeys.has(rowKey(row));
            cb.disabled = false;
            cb.addEventListener('click', function (ev) {
                ev.stopPropagation();
            });
            cb.addEventListener('change', function () {
                const k = rowKey(row);
                if (cb.checked) selectedKeys.add(k);
                else selectedKeys.delete(k);
                updateBulkCount();
            });
            pick.appendChild(cb);
            li.appendChild(pick);
            li.appendChild(btn);
            host.appendChild(li);
        });
        if (window.ms365GroupPhotoThumb && typeof window.ms365GroupPhotoThumb.hydrate === 'function') {
            window.ms365GroupPhotoThumb.hydrate(host);
        }
        updateBulkCount();
        updateClassActionButtons();
    }

    function ensureActiveKey() {
        if (!classes.length) {
            activeKey = '';
            return;
        }
        const has = classes.some(function (r) {
            return rowKey(r) === (normCode(activeKey) || normStr(activeKey).toUpperCase());
        });
        if (!has) activeKey = rowKey(classes[0]);
    }

    function setActiveKey(code) {
        activeKey = normCode(code) || normStr(code).toUpperCase();
        const search = document.getElementById('slgGroupSearch');
        if (search) search.value = '';
        gd().clearSearchResults();
        renderLeftList();
        applyCreateDefaults();
        gd().setTab('general');
        refreshMatchUi();
        if (getActiveGroupId()) live().loadGroup({ silent: true });
    }

    function persistMatchForRow(row, g, mode) {
        const api = dataV2();
        if (!api || typeof api.upsertClassTeam !== 'function') {
            throw new Error('classTeams-Speicher (app-data-v2) nicht verfügbar.');
        }
        if (!row) return;
        const existing = findClassTeam(row);
        const gid = g && g.id ? String(g.id).trim() : '';
        const displayNick = graphMailNick(g && g.mailNickname);
        const schemaNick = deriveNick(row);
        // Nach Löschen/Unmatchen: alten Graph-Alias nicht wiederverwenden.
        const nick = gid
            ? sanitizeNick(displayNick) || sanitizeNick(existing && existing.stableMailNickname) || schemaNick
            : schemaNick || sanitizeNick(displayNick);
        if (!nick) {
            throw new Error('Kein gültiger Mail‑Nickname für diese Klasse (Kürzel und Abschlussjahr prüfen).');
        }
        const pretty = gid
            ? displayNick || graphMailNick(existing && existing.mailNickname) || nick
            : schemaNick || displayNick || nick;
        api.upsertClassTeam({
            stableMailNickname: nick,
            mailNickname: pretty,
            graphGroupId: gid,
            classCode: row && row.code ? row.code : '',
            displayName: (g && g.displayName) || (row && row.name) || '',
            abschlussJahr: row && row.year ? row.year : '',
            mode: mode
        });
    }

    function persistMatch(g, mode) {
        try {
            persistMatchForRow(getActiveRow(), g, mode);
        } catch (e) {
            toast(e.message || String(e));
            return;
        }
        renderLeftList();
    }

    function syncPersistedAliasFromForm() {
        const gid = getActiveGroupId();
        if (!gid) {
            renderLeftList();
            return;
        }
        const aliasEl = document.getElementById('slgLiveAlias');
        const nameEl = document.getElementById('slgLiveName');
        const team = findClassTeam(getActiveRow());
        persistMatch(
            {
                id: gid,
                displayName: nameEl ? nameEl.value : '',
                mailNickname: aliasEl ? aliasEl.value : ''
            },
            (team && team.mode) || 'matched'
        );
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
        if (!el) return;
        el.replaceChildren();
    }

    function collectStudentEmails() {
        return students
            .map(function (s) {
                return String((s && s.email) || '')
                    .trim()
                    .toLowerCase();
            })
            .filter(function (em) {
                return em.indexOf('@') !== -1;
            });
    }

    /**
     * @returns {Promise<{ empty: boolean, unchanged: boolean, join: number, leave: number, skip: number, fail: number }>}
     */
    async function syncMembersForGroup(token, row, gid, logFn) {
        const log = typeof logFn === 'function' ? logFn : function () {};
        const emails = emailsForClass(row);
        const result = { empty: false, unchanged: false, join: 0, leave: 0, skip: 0, fail: 0 };
        if (!emails.length) {
            result.empty = true;
            return result;
        }
        const lc = window.ms365StudentClassLifecycle;
        let joinEmails = emails;
        let leaveEmails = [];
        if (lc && typeof lc.reconcileClassMembers === 'function' && typeof gug().fetchGroupMembers === 'function') {
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
            const rec = lc.reconcileClassMembers(emails, collectStudentEmails(), current);
            joinEmails = rec.join;
            leaveEmails = rec.leave;
            log(
                'Abgleich: +' + joinEmails.length + ' / −' + leaveEmails.length + ' (Lehrer und andere Mitglieder bleiben).',
                ''
            );
        }
        if (joinEmails.length) {
            const r = await gug().syncEmailsToGroup(token, gid, joinEmails, 'Klasse', log);
            result.join = r.ok || 0;
            result.skip += r.skip || 0;
            result.fail += r.fail || 0;
            log('Aufnehmen: neu ' + r.ok + ', übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
        }
        if (leaveEmails.length && typeof gug().removeEmailsFromGroup === 'function') {
            const r = await gug().removeEmailsFromGroup(token, gid, leaveEmails, 'Klasse', log);
            result.leave = r.ok || 0;
            result.skip += r.skip || 0;
            result.fail += r.fail || 0;
            log('Entfernen: ' + r.ok + ' OK, übersprungen ' + r.skip + ', Fehler ' + r.fail + '.', 'ok');
        }
        if (!joinEmails.length && !leaveEmails.length) {
            result.unchanged = true;
            log('Keine Änderungen gegenüber der Stammliste.', 'ok');
        }
        const dirOwners = ownersGlobal();
        if (dirOwners.length) await gug().ensureOwners(token, gid, dirOwners);
        return result;
    }

    async function runSyncMembers() {
        const gid = getActiveGroupId();
        if (!gid) {
            toast('Zuerst eine Gruppe matchen oder anlegen.');
            return;
        }
        const emails = emailsForClass(getActiveRow());
        if (!emails.length) {
            toast('Keine Schüler‑E‑Mails für diese Klasse in den Stammdaten.');
            return;
        }
        clearSyncLog();
        appendSyncLog('Start: Klasse (' + emails.length + ' Adressen) …', '');
        try {
            const token = await gug().getGraphToken();
            await syncMembersForGroup(token, getActiveRow(), gid, appendSyncLog);
            live().invalidateMembership();
            await live().loadMembers();
            await refreshGraphMemberCounts();
            toast('Synchronisation abgeschlossen.');
        } catch (e) {
            appendSyncLog('Abbruch: ' + (e.message || e), 'err');
            toast('Fehler: ' + (e.message || e));
        }
    }

    function sleep(ms) {
        return new Promise(function (r) {
            setTimeout(r, ms);
        });
    }

    let bulkBusy = false;
    let bulkManualOpen = false;
    let bulkUserCollapsed = false;
    let lastBulkSelectCount = -1;

    function setBulkExpanded(open) {
        const box = document.getElementById('jgBulk');
        const toggle = document.getElementById('jgBulkToggle');
        if (!box) return;
        box.classList.toggle('is-collapsed', !open);
        if (toggle) toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
    }

    function syncBulkExpanded() {
        if (bulkBusy) {
            setBulkExpanded(true);
            return;
        }
        if (bulkUserCollapsed) {
            setBulkExpanded(false);
            return;
        }
        if (bulkManualOpen) {
            setBulkExpanded(true);
            return;
        }
        setBulkExpanded(selectedKeys.size >= 2);
    }

    function setBulkProgress(done, total, label) {
        const wrap = document.getElementById('jgBulkProgress');
        const bar = document.getElementById('jgBulkProgressBar');
        const lab = document.getElementById('jgBulkProgressLabel');
        if (!wrap || !bar) return;
        const max = Math.max(0, Number(total) || 0);
        const cur = Math.max(0, Math.min(max, Number(done) || 0));
        const pct = max ? Math.round((cur / max) * 100) : 0;
        wrap.hidden = false;
        bar.style.width = pct + '%';
        wrap.setAttribute('aria-valuemin', '0');
        wrap.setAttribute('aria-valuemax', String(max || 100));
        wrap.setAttribute('aria-valuenow', String(cur));
        if (lab) {
            lab.textContent = (label || 'Fortschritt') + '  ·  ' + cur + ' / ' + max + '  (' + pct + ' %)';
        }
    }

    function clearBulkProgress() {
        const wrap = document.getElementById('jgBulkProgress');
        const bar = document.getElementById('jgBulkProgressBar');
        if (bar) bar.style.width = '0';
        if (wrap) wrap.hidden = true;
    }

    function setBulkStatus(text, show) {
        const el = document.getElementById('jgBulkStatus');
        if (!el) return;
        if (show === false || !text) {
            el.hidden = true;
            el.textContent = '';
            if (!bulkBusy) clearBulkProgress();
            return;
        }
        el.hidden = false;
        el.textContent = text;
        setBulkExpanded(true);
    }

    function beginBulkJob(total, label) {
        bulkBusy = true;
        setBulkExpanded(true);
        setBulkProgress(0, total, label);
        setBulkStatus(label + ' …');
    }

    function finishBulkJob() {
        bulkBusy = false;
        syncBulkExpanded();
    }

    function pruneSelection() {
        const keep = new Set();
        selectedKeys.forEach(function (key) {
            for (let i = 0; i < classes.length; i++) {
                if (rowKey(classes[i]) === key) {
                    keep.add(key);
                    break;
                }
            }
        });
        selectedKeys = keep;
    }

    function collectSelectedMatched() {
        pruneSelection();
        const out = [];
        selectedKeys.forEach(function (key) {
            let row = null;
            for (let i = 0; i < classes.length; i++) {
                if (rowKey(classes[i]) === key) {
                    row = classes[i];
                    break;
                }
            }
            if (!row) return;
            const team = findClassTeam(row);
            const id = team && team.graphGroupId ? String(team.graphGroupId).trim() : '';
            if (!id) return;
            out.push({
                key: key,
                row: row,
                team: team,
                id: id,
                name: normStr(row.name) || normStr(row.code) || id
            });
        });
        return out;
    }

    function collectSelectedUnmatched() {
        pruneSelection();
        const out = [];
        selectedKeys.forEach(function (key) {
            let row = null;
            for (let i = 0; i < classes.length; i++) {
                if (rowKey(classes[i]) === key) { row = classes[i]; break; }
            }
            if (!row) return;
            const team = findClassTeam(row);
            if (team && team.graphGroupId) return; // already matched
            out.push({ key: key, row: row, name: normStr(row.name) || normStr(row.code) });
        });
        return out;
    }

    function updateBulkCount() {
        pruneSelection();
        const n = selectedKeys.size;
        const el = document.getElementById('jgBulkCount');
        if (el) {
            const label = n === 1 ? '1 Klasse ausgewählt' : String(n) + ' Klassen ausgewählt';
            el.innerHTML = '<i class="bi bi-check2-square" aria-hidden="true"></i>' + label;
            el.classList.toggle('is-active', n > 0);
        }
        if (n !== lastBulkSelectCount) {
            lastBulkSelectCount = n;
            bulkUserCollapsed = false;
            bulkManualOpen = false;
        }
        syncBulkExpanded();
    }

    function visibleMatchedRows() {
        const q = listFilter.toLowerCase();
        return classes.filter(function (row) {
            const team = findClassTeam(row);
            if (!team || !team.graphGroupId) return false;
            if (!q) return true;
            const nick = persistNickForRow(row);
            const hay = (row.code + ' ' + (row.name || '') + ' ' + (row.year || '') + ' ' + nick).toLowerCase();
            return hay.indexOf(q) !== -1;
        });
    }

    function selectVisibleMatched() {
        visibleMatchedRows().forEach(function (row) {
            selectedKeys.add(rowKey(row));
        });
        renderLeftList();
        const n = collectSelectedMatched().length;
        toast(n ? String(n) + ' gematchte Klasse(n) angekreuzt.' : 'Keine gematchten Klassen in der aktuellen Liste.');
    }

    function visibleUnmatchedRows() {
        const q = listFilter.toLowerCase();
        return classes.filter(function (row) {
            const team = findClassTeam(row);
            if (team && team.graphGroupId) return false;
            if (!q) return true;
            const nick = persistNickForRow(row);
            const hay = (row.code + ' ' + (row.name || '') + ' ' + (row.year || '') + ' ' + nick).toLowerCase();
            return hay.indexOf(q) !== -1;
        });
    }

    function selectVisibleUnmatched() {
        visibleUnmatchedRows().forEach(function (row) {
            selectedKeys.add(rowKey(row));
        });
        renderLeftList();
        const n = collectSelectedUnmatched().length;
        toast(n ? String(n) + ' ungematchte Klasse(n) angekreuzt.' : 'Keine ungematchten Klassen in der aktuellen Liste.');
    }

    function clearSelection() {
        selectedKeys = new Set();
        renderLeftList();
    }

    function showBulkOwnerPanel(show) {
        const panel = document.getElementById('jgBulkOwnerPanel');
        if (!panel) return;
        panel.hidden = !show;
        if (show) {
            const inp = document.getElementById('jgBulkOwnerSearch');
            if (inp) inp.focus();
        }
    }

    function fillBulkOwnerSelect(users) {
        const sel = document.getElementById('jgBulkOwnerResults');
        if (!sel) return;
        sel.replaceChildren();
        if (!users || !users.length) {
            const opt = document.createElement('option');
            opt.value = '';
            opt.textContent = '(keine Treffer)';
            sel.appendChild(opt);
            return;
        }
        users.forEach(function (u) {
            const opt = document.createElement('option');
            opt.value = u.id || '';
            opt.textContent = gug().personLabel(u) || (u.id ? String(u.id) : '');
            sel.appendChild(opt);
        });
    }

    async function runBulkOwnerSearch() {
        const inp = document.getElementById('jgBulkOwnerSearch');
        const q = inp ? String(inp.value || '').trim() : '';
        if (!q) {
            toast('Bitte einen Namen oder eine E‑Mail eingeben.');
            return;
        }
        const btn = document.getElementById('jgBulkOwnerSearchBtn');
        if (btn) btn.disabled = true;
        try {
            const token = await gug().getGraphToken();
            const users = await gug().searchUsers(token, q);
            fillBulkOwnerSelect(users);
            toast('Suche: ' + users.length + ' Treffer.');
        } catch (e) {
            toast('Suche: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
        }
    }

    async function runBulkSetOwner() {
        const items = collectSelectedMatched();
        if (!items.length) {
            toast('Bitte zuerst gematchte Klassen ankreuzen.');
            return;
        }
        const sel = document.getElementById('jgBulkOwnerResults');
        const userId = sel && sel.value ? String(sel.value).trim() : '';
        if (!userId) {
            toast('Bitte zuerst einen Benutzer suchen und auswählen.');
            showBulkOwnerPanel(true);
            return;
        }
        const label = sel.options[sel.selectedIndex] ? String(sel.options[sel.selectedIndex].textContent || '').trim() : userId;
        if (
            !(await dlgConfirm(
                '„' +
                    label +
                    '“ als Besitzer zu ' +
                    String(items.length) +
                    ' Gruppe(n) hinzufügen?\n\nBestehende Besitzer bleiben erhalten.',
                { title: 'Besitzer setzen', okText: 'Hinzufügen' }
            ))
        ) {
            return;
        }
        const applyBtn = document.getElementById('jgBulkOwnerApply');
        if (applyBtn) applyBtn.disabled = true;
        let ok = 0;
        let skip = 0;
        let fail = 0;
        const lines = [];
        beginBulkJob(items.length, 'Besitzer wird gesetzt');
        try {
            const token = await gug().getGraphToken();
            for (let i = 0; i < items.length; i++) {
                const it = items[i];
                setBulkProgress(i, items.length, 'Besitzer: ' + it.name);
                try {
                    await gug().addOwnerWithMemberFallback(token, it.id, userId);
                    ok++;
                    lines.push('OK  ' + it.name);
                } catch (e) {
                    if (gug().isDuplicateMemberError(e)) {
                        skip++;
                        lines.push('schon Besitzer  ' + it.name);
                    } else {
                        fail++;
                        lines.push('Fehler  ' + it.name + ': ' + (e.message || e));
                    }
                }
                setBulkProgress(i + 1, items.length, 'Besitzer: ' + it.name);
                if ((i + 1) % 6 === 0) await sleep(120);
            }
            setBulkProgress(items.length, items.length, 'Besitzer fertig');
            setBulkStatus(lines.join('\n'));
            toast('Besitzer: neu ' + ok + ', bereits vorhanden ' + skip + ', Fehler ' + fail + '.');
            if (getActiveGroupId()) {
                try {
                    live().invalidateMembership();
                    if (gd().getActiveTab() === 'owners') await live().loadOwners();
                } catch {
                    /* ignore */
                }
            }
        } catch (e) {
            setBulkStatus('Abbruch: ' + (e.message || e));
            toast('Besitzer setzen: ' + (e.message || e));
        } finally {
            if (applyBtn) applyBtn.disabled = false;
            finishBulkJob();
        }
    }

    async function runBulkSyncMembers() {
        const items = collectSelectedMatched();
        if (!items.length) {
            toast('Bitte zuerst gematchte Klassen ankreuzen.');
            return;
        }
        const preview =
            items
                .slice(0, 12)
                .map(function (it) {
                    return it.name;
                })
                .join('\n') + (items.length > 12 ? '\n…' : '');
        if (
            !(await dlgConfirm(
                String(items.length) +
                    ' Klassengruppe(n) mit der Schülerliste abgleichen?\n\n' +
                    preview +
                    '\n\nFehlende Schüler:innen dieser Klasse werden aufgenommen. Schüler:innen, die laut Stammliste in einer anderen Klasse stehen, werden entfernt. Lehrkräfte und sonstige Mitglieder bleiben.',
                { title: 'Mitglieder synchronisieren', okText: 'Synchronisieren' }
            ))
        ) {
            return;
        }
        const btn = document.getElementById('jgBtnBulkSyncMembers');
        if (btn) btn.disabled = true;
        let ok = 0;
        let empty = 0;
        let unchanged = 0;
        let fail = 0;
        let joinTotal = 0;
        let leaveTotal = 0;
        const lines = [];
        beginBulkJob(items.length, 'Mitglieder werden abgeglichen');
        try {
            const token = await gug().getGraphToken();
            for (let i = 0; i < items.length; i++) {
                const it = items[i];
                setBulkProgress(i, items.length, 'Mitglieder: ' + it.name);
                try {
                    const r = await syncMembersForGroup(token, it.row, it.id, function () {});
                    if (r.empty) {
                        empty++;
                        lines.push('keine Schüler  ' + it.name);
                    } else if (r.fail) {
                        fail++;
                        joinTotal += r.join;
                        leaveTotal += r.leave;
                        lines.push(
                            'teilweise  ' +
                                it.name +
                                ': +' +
                                r.join +
                                ' / −' +
                                r.leave +
                                ', Fehler ' +
                                r.fail
                        );
                    } else if (r.unchanged) {
                        unchanged++;
                        lines.push('unverändert  ' + it.name);
                    } else {
                        ok++;
                        joinTotal += r.join;
                        leaveTotal += r.leave;
                        lines.push('OK  ' + it.name + ': +' + r.join + ' / −' + r.leave);
                    }
                } catch (e) {
                    fail++;
                    lines.push('Fehler  ' + it.name + ': ' + (e.message || e));
                }
                setBulkProgress(i + 1, items.length, 'Mitglieder: ' + it.name);
                if ((i + 1) % 4 === 0) await sleep(200);
            }
            setBulkProgress(items.length, items.length, 'Mitglieder fertig');
            setBulkStatus(lines.join('\n'));
            toast(
                'Mitglieder: ' +
                    ok +
                    ' angepasst, ' +
                    unchanged +
                    ' unverändert' +
                    (empty ? ', ' + empty + ' ohne Schüler' : '') +
                    ', ' +
                    fail +
                    ' Fehler (+' +
                    joinTotal +
                    ' / −' +
                    leaveTotal +
                    ').'
            );
            if (getActiveGroupId()) {
                try {
                    live().invalidateMembership();
                    if (gd().getActiveTab() === 'members') await live().loadMembers();
                } catch {
                    /* ignore */
                }
            }
        } catch (e) {
            setBulkStatus('Abbruch: ' + (e.message || e));
            toast('Mitglieder: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
            finishBulkJob();
        }
    }

    async function runBulkDelete() {
        const items = collectSelectedMatched();
        if (!items.length) {
            toast('Bitte zuerst gematchte Klassen ankreuzen.');
            return;
        }
        const preview =
            items
                .slice(0, 12)
                .map(function (it) {
                    return it.name;
                })
                .join('\n') + (items.length > 12 ? '\n…' : '');
        if (
            !(await dlgConfirm(
                String(items.length) +
                    ' Microsoft‑365‑Gruppe(n) wirklich löschen?\n\n' +
                    preview +
                    '\n\nDie Gruppen verschwinden in Entra/Teams. Das lokale Match wird gelöst.\n\nHinweis: Gelöschte Gruppen liegen oft noch 30 Tage im Entra-Papierkorb. Der alte Alias (z. B. jg20311bk) kann dann noch reserviert sein – in Entra ggf. endgültig löschen, bevor Sie neu anlegen.',
                { title: 'Gruppen löschen', okText: 'Löschen', danger: true }
            ))
        ) {
            return;
        }
        const delBtn = document.getElementById('jgBtnBulkDelete');
        if (delBtn) delBtn.disabled = true;
        let ok = 0;
        let fail = 0;
        const lines = [];
        const deletedKeys = [];
        beginBulkJob(items.length, 'Gruppen werden gelöscht');
        try {
            const token = await gug().getGraphToken();
            for (let i = 0; i < items.length; i++) {
                const it = items[i];
                setBulkProgress(i, items.length, 'Löschen: ' + it.name);
                try {
                    if (typeof gug().deleteUnifiedGroup !== 'function') {
                        throw new Error('deleteUnifiedGroup fehlt.');
                    }
                    await gug().deleteUnifiedGroup(token, it.id);
                    persistMatchForRow(it.row, { id: '', displayName: '', mailNickname: '' }, '');
                    selectedKeys.delete(it.key);
                    deletedKeys.push(it.key);
                    ok++;
                    lines.push('gelöscht  ' + it.name);
                } catch (e) {
                    const msg = String((e && e.message) || e || '');
                    if (/\b404\b/.test(msg) || /Request_ResourceNotFound/i.test(msg)) {
                        persistMatchForRow(it.row, { id: '', displayName: '', mailNickname: '' }, '');
                        selectedKeys.delete(it.key);
                        deletedKeys.push(it.key);
                        ok++;
                        lines.push('bereits weg  ' + it.name);
                    } else {
                        fail++;
                        lines.push('Fehler  ' + it.name + ': ' + msg);
                    }
                }
                setBulkProgress(i + 1, items.length, 'Löschen: ' + it.name);
                if ((i + 1) % 4 === 0) await sleep(200);
            }
            renderLeftList();
            applyCreateDefaults();
            refreshMatchUi();
            if (deletedKeys.indexOf(normCode(activeKey) || normStr(activeKey).toUpperCase()) >= 0) {
                live().loadGroup({ silent: true });
            }
            setBulkProgress(items.length, items.length, 'Löschen fertig');
            setBulkStatus(lines.join('\n'));
            toast('Löschen: ' + ok + ' erledigt, ' + fail + ' Fehler.');
        } catch (e) {
            renderLeftList();
            setBulkStatus('Abbruch: ' + (e.message || e));
            toast('Löschen: ' + (e.message || e));
        } finally {
            if (delBtn) delBtn.disabled = false;
            finishBulkJob();
        }
    }

    async function runBulkCreateAndMatch() {
        const items = collectSelectedUnmatched();
        if (!items.length) {
            toast('Bitte zuerst ungematchte Klassen ankreuzen (rote Badges).');
            return;
        }
        const preview = items.slice(0, 12).map(function (it) { return it.name; }).join('\n') +
            (items.length > 12 ? '\n…' : '');
        if (!(await dlgConfirm(
            String(items.length) + ' ungematchte Klasse(n) als Microsoft‑365‑Gruppe anlegen und matchen?\n\n' +
            preview +
            '\n\nAlias nach aktuellem Schema, z. B. ' +
            (persistNickForRow(items[0].row) || '–') +
            '. Besitzer: ' +
            (getOwnerOptions().useKV ? 'Klassenvorstand' : '') +
            (getOwnerOptions().useKV && getOwnerOptions().useDirektion ? ' + ' : '') +
            (getOwnerOptions().useDirektion ? 'Direktion' : '') +
            (getJgCreateTeam() ? '. Pro Gruppe wird auch ein Microsoft Team bereitgestellt.' : '.'),
            { title: 'Gruppen anlegen & matchen', okText: 'Anlegen' }
        ))) return;
        const btn = document.getElementById('jgBtnBulkCreate');
        const createTeam = getJgCreateTeam();
        if (btn) btn.disabled = true;
        let ok = 0;
        let fail = 0;
        const lines = [];
        beginBulkJob(items.length, 'Gruppen werden angelegt');
        try {
            const token = await gug().getGraphToken();
            for (let i = 0; i < items.length; i++) {
                const it = items[i];
                setBulkProgress(i, items.length, 'Anlegen: ' + it.name);
                const row = it.row;
                const nick = persistNickForRow(row);
                if (!nick) {
                    fail++;
                    lines.push('Fehler  ' + it.name + ': kein Alias ableitbar (Abschlussjahr prüfen).');
                    continue;
                }
                const displayName = normStr(row.name) || normStr(row.code) || nick;
                const desc = 'Jahrgangsgruppe ' + displayName +
                    (row.year ? ' / Abschluss ' + row.year : '') + ' (MS365-Schulverwaltung)';
                try {
                    const g = await gug().createUnifiedGroup(token, displayName, nick, desc);
                    const rowOwners = ownersForRow(row);
                    if (rowOwners.length) await gug().ensureOwners(token, g.id, rowOwners);
                    const emails = emailsForClass(row);
                    if (emails.length) {
                        await gug().syncEmailsToGroup(token, g.id, emails, 'Klasse', function () {});
                    }
                    if (createTeam && typeof gug().provisionTeamForGroup === 'function') {
                        try {
                            await gug().provisionTeamForGroup(token, g.id);
                        } catch (teamErr) {
                            const msg = String((teamErr && teamErr.message) || teamErr || '');
                            if (/\b409\b/.test(msg) || /Conflict|already exists|already provisioned/i.test(msg)) {
                                /* Team bereits vorhanden */
                            } else {
                                lines.push('Team fehlgeschlagen  ' + it.name + ': ' + msg);
                            }
                        }
                    }
                    persistMatchForRow(row, g, 'created');
                    selectedKeys.delete(it.key);
                    ok++;
                    lines.push('OK  ' + it.name + '  →  ' + nick + (createTeam ? ' (+ Team)' : ''));
                    if (window.ms365ActionLog && typeof window.ms365ActionLog.append === 'function') {
                        window.ms365ActionLog.append({
                            tool: 'jg', action: 'createAndMatch', target: nick, summary: displayName, result: 'ok'
                        });
                    }
                } catch (e) {
                    fail++;
                    lines.push('Fehler  ' + it.name + ': ' + (e.message || e));
                }
                setBulkProgress(i + 1, items.length, 'Anlegen: ' + it.name);
                if ((i + 1) % 4 === 0) await sleep(200);
            }
            renderLeftList();
            setBulkProgress(items.length, items.length, 'Anlegen fertig');
            setBulkStatus(lines.join('\n'));
            toast('Anlegen: ' + ok + ' OK, ' + fail + ' Fehler.');
            if (ok > 0 && getActiveGroupId()) {
                try { live().invalidateMembership(); } catch (_) { /* ignore */ }
            }
        } catch (e) {
            renderLeftList();
            setBulkStatus('Abbruch: ' + (e.message || e));
            toast('Anlegen: ' + (e.message || e));
        } finally {
            if (btn) btn.disabled = false;
            finishBulkJob();
        }
    }

    function onClick(id, fn) {
        const el = document.getElementById(id);
        if (el) el.addEventListener('click', fn);
    }

    function mountDetail() {
        gd().mount('#groupDetailHost', {
            title: 'Jahrgangsgruppe',
            searchPlaceholder: 'z. B. 1AK oder jg2030ak',
            unmatchedCreateHint:
                'Legt eine Microsoft 365‑Gruppe (Unified) an und verknüpft sie mit dieser Klasse. Standardmäßig wird auch ein Microsoft Team bereitgestellt (unter Sammelaktionen abschaltbar).',
            membersUnmatchedHint:
                'Wenn in den Stammdaten Schüler:innen mit E‑Mail für diese Klasse hinterlegt sind, können sie nach dem Match additiv synchronisiert werden. Sonst Mitglieder live in Graph pflegen.',
            membersUnmatchedTitle: 'Schüler:innen dieser Klasse',
            membersMatchedHint:
                'Live aus Microsoft Graph. „Mitglieder synchronisieren“ fügt fehlende Schüler‑Adressen dieser Klasse hinzu und entfernt Schüler:innen, die laut Stammliste in einer anderen Klasse stehen (Lehrkräfte bleiben).',
            emptyHintHtml:
                'Keine Klassen in diesem Schuljahr. Legen Sie eine Klasse über <strong>Neu</strong> an oder pflegen Sie Klassen unter <a href="../tenant.html#classes">Stammdaten</a>.',
            features: {
                aliasEditable: true,
                smtpSlot: true,
                syncMembers: true,
                membershipReview: true,
                emptyHint: true
            },
            ids: { emptyHint: 'jgEmptyHint', wrap: 'jgDetailWrap' },
            live: {
                toast: toast,
                dlgConfirm: dlgConfirm,
                getGroupId: getActiveGroupId,
                ensureDirektionOwners: function (token, gid) {
                    const owners = ownersForRow(getActiveRow());
                    if (!owners.length) throw new Error('Kein Klassenvorstand hinterlegt und keine Direktion ausgewählt.');
                    return gug().ensureOwners(token, gid, owners);
                },
                onUnmatched: function () {
                    renderOwnerPreview();
                    renderMemberPreview();
                    renderLeftList();
                },
                onAfterLoad: function () {
                    syncPersistedAliasFromForm();
                    refreshSmtpHint();
                    void refreshGraphMemberCounts();
                },
                onAfterUpdate: function (group) {
                    const aliasEl = document.getElementById('slgLiveAlias');
                    const nick = graphMailNick(aliasEl && aliasEl.value) || (group && group.mailNickname);
                    const team = findClassTeam(getActiveRow());
                    persistMatch(
                        Object.assign({}, group || {}, nick ? { mailNickname: nick } : {}),
                        (team && team.mode) || 'created'
                    );
                    refreshSmtpHint();
                    const mailEl = document.getElementById('slgLiveMail');
                    const wanted = String((mailEl && mailEl.value) || '').trim();
                    const graphMail = String((mailEl && mailEl.getAttribute('data-graph-mail')) || '').trim();
                    if (wanted && graphMail && wanted.toLowerCase() !== graphMail.toLowerCase()) {
                        const pack = collectSmtpScriptItems(true);
                        if (pack.domain && pack.items.length) {
                            showSmtpScript(buildClassSmtpPs1(pack.items, pack.domain));
                        }
                        return (
                            ' Graph ändert die Gruppen-E-Mail nicht (nur den Alias). Exchange-Skript für ' +
                            wanted +
                            ' liegt unter E-Mail (kopiert).'
                        );
                    }
                    return '';
                }
            },
            match: {
                persistMatch: persistMatch,
                persistUnmatch: function () {
                    persistMatch({ id: '', displayName: '', mailNickname: '' }, '');
                },
                canSearch: function () {
                    return activeKey
                        ? { ok: true }
                        : { ok: false, message: 'Bitte zuerst eine Klasse wählen.' };
                },
                canCreate: function () {
                    return activeKey
                        ? { ok: true }
                        : { ok: false, message: 'Bitte zuerst eine Klasse wählen.' };
                },
                ensureOwners: function (token, gid) {
                    const owners = ownersForRow(getActiveRow());
                    if (!owners.length) return Promise.resolve();
                    return gug().ensureOwners(token, gid, owners);
                },
                afterCreate: async function (token, g) {
                    const emails = emailsForClass(getActiveRow());
                    if (emails.length) {
                        await gug().syncEmailsToGroup(token, g.id, emails, 'Klasse', function () {});
                    }
                }
            },
            onTabUnmatched: function (tab) {
                if (tab === 'owners') renderOwnerPreview();
                if (tab === 'members') renderMemberPreview();
            }
        });
    }

    function wire() {
        const listHost = document.getElementById('jgListItems');
        if (listHost) {
            listHost.addEventListener('click', function (ev) {
                const t = ev.target;
                const item = t && t.closest ? t.closest('button[data-jg-code]') : null;
                if (!item) return;
                setActiveKey(item.getAttribute('data-jg-code') || '');
            });
        }
        const filter = document.getElementById('jgListFilter');
        if (filter) {
            filter.addEventListener('input', function () {
                listFilter = String(filter.value || '').trim();
                renderLeftList();
            });
        }
        onClick('slgBtnReloadLists', function () {
            readLists();
            ensureActiveKey();
            renderLeftList();
            applyCreateDefaults();
            refreshMatchUi();
            toast('Listen neu eingelesen.');
        });
        onClick('jgBtnAddClass', function () {
            openClassModal('create');
        });
        onClick('jgBtnEditClass', function () {
            openClassModal('edit');
        });
        onClick('jgBtnDeleteClass', function () {
            void deleteActiveClass();
        });
        onClick('jgClassModalCancel', function () {
            closeClassModal();
        });
        onClick('jgClassModalSave', function () {
            void submitClassModal();
        });
        const classModal = document.getElementById('jgClassModal');
        if (classModal) {
            classModal.addEventListener('click', function (ev) {
                if (ev.target === classModal) closeClassModal();
            });
        }
        ['jgNewCode', 'jgNewName', 'jgNewYear', 'jgNewHeadName', 'jgNewHeadEmail'].forEach(function (id) {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener('keydown', function (ev) {
                if (ev.key !== 'Enter' || ev.shiftKey) return;
                if (!classModal || !classModal.classList.contains('open')) return;
                ev.preventDefault();
                void submitClassModal();
            });
        });
        document.addEventListener('keydown', function (ev) {
            if (ev.key !== 'Escape') return;
            if (classModal && classModal.classList.contains('open')) closeClassModal();
        });
        onClick('slgBtnSync', runSyncMembers);
        onClick('jgBtnSmtpThis', function () {
            runSmtpScript(true);
        });
        onClick('jgBtnSmtpAll', function () {
            runSmtpScript(false);
        });
        const bulkToggle = document.getElementById('jgBulkToggle');
        if (bulkToggle) {
            bulkToggle.addEventListener('click', function () {
                const box = document.getElementById('jgBulk');
                const willOpen = !!(box && box.classList.contains('is-collapsed'));
                bulkManualOpen = willOpen;
                bulkUserCollapsed = !willOpen;
                setBulkExpanded(willOpen);
            });
        }
        onClick('jgBtnSelectMatched', selectVisibleMatched);
        onClick('jgBtnSelectUnmatched', selectVisibleUnmatched);
        onClick('jgBtnSelectNone', clearSelection);
        onClick('jgBtnBulkCreate', function () {
            runBulkCreateAndMatch().catch(function () {});
        });
        onClick('jgBtnBulkSyncMembers', function () {
            runBulkSyncMembers().catch(function () {});
        });
        onClick('jgBtnBulkOwner', function () {
            if (!collectSelectedMatched().length) {
                toast('Bitte zuerst gematchte Klassen ankreuzen.');
                return;
            }
            showBulkOwnerPanel(true);
        });
        onClick('jgBtnBulkDelete', function () {
            runBulkDelete().catch(function () {});
        });
        onClick('jgBulkOwnerSearchBtn', function () {
            runBulkOwnerSearch().catch(function () {});
        });
        onClick('jgBulkOwnerApply', function () {
            runBulkSetOwner().catch(function () {});
        });
        const bulkOwnerSearch = document.getElementById('jgBulkOwnerSearch');
        if (bulkOwnerSearch) {
            bulkOwnerSearch.addEventListener('keydown', function (ev) {
                if (ev.key === 'Enter') {
                    ev.preventDefault();
                    runBulkOwnerSearch().catch(function () {});
                }
            });
        }
        // ── Nomenklatur (Alias-Schema) ───────────────────────────────────────
        function updateNickPreview() {
            const prefix = (document.getElementById('jgNickPrefix')?.value || 'jg').trim().toLowerCase().replace(/[^a-z0-9]/g,'') || 'jg';
            const pattern = document.getElementById('jgNickPattern')?.value || '{prefix}{year}-{suffix}';
            const upper = document.getElementById('jgNickUpper')?.checked;
            const exSuffix = upper ? 'AK' : 'ak';
            const exYear = '2031';
            const exKlasse = upper ? '1AK' : '1ak';
            const raw = pattern
                .replaceAll('{prefix}', prefix)
                .replaceAll('{year}', exYear)
                .replaceAll('{suffix}', exSuffix)
                .replaceAll('{klasse}', exKlasse)
                .replaceAll('{kv}', upper ? 'SCW' : 'scw');
            const sanitized = raw.trim().replace(/\s+/g,'-').replace(/[^a-zA-Z0-9-]/g,'').replace(/-+/g,'-').replace(/^-|-$/g,'').toLowerCase();
            const el = document.getElementById('jgNickPreview');
            if (el) el.textContent = sanitized || '(leer)';
        }
        function saveNickSchema() {
            if (typeof window.ms365SaveClassNickSchema === 'function') {
                window.ms365SaveClassNickSchema({
                    prefix: document.getElementById('jgNickPrefix')?.value || 'jg',
                    pattern: document.getElementById('jgNickPattern')?.value || '{prefix}{year}-{suffix}',
                    upper: !!document.getElementById('jgNickUpper')?.checked
                });
            }
            updateNickPreview();
            renderLeftList();
            applyCreateDefaults();
            refreshSmtpHint();
        }
        ['jgNickPrefix', 'jgNickPattern', 'jgNickUpper'].forEach(function (id) {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener('input', saveNickSchema);
            el.addEventListener('change', saveNickSchema);
        });
        updateNickPreview();

        function onJgCreateTeamChange(ev) {
            const el = ev && ev.target;
            const on = el ? !!el.checked : getJgCreateTeam();
            saveJgCreateTeam(on);
            syncJgCreateTeamUi();
        }
        ['jgBulkCreateTeam', 'slgNewCreateTeam'].forEach(function (id) {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener('change', onJgCreateTeamChange);
        });
        syncJgCreateTeamUi();

        ['jgOwnerUseKV', 'jgOwnerUseDirektion'].forEach(function (id) {
            const el = document.getElementById(id);
            if (!el) return;
            el.addEventListener('change', function () {
                // Mindestens eine Option muss aktiv bleiben
                const kv = document.getElementById('jgOwnerUseKV');
                const dir = document.getElementById('jgOwnerUseDirektion');
                if (kv && dir && !kv.checked && !dir.checked) {
                    el.checked = true; // zurücksetzen
                    toast('Mindestens eine Besitzer-Option muss aktiv sein.');
                    return;
                }
                renderOwnerPreview();
            });
        });
        const aliasInp = document.getElementById('slgLiveAlias');
        if (aliasInp && !aliasInp.readOnly) {
            aliasInp.addEventListener('input', function () {
                refreshSmtpHint();
            });
            aliasInp.addEventListener('blur', function () {
                const n = graphMailNick(aliasInp.value);
                if (n && String(aliasInp.value || '').trim()) aliasInp.value = n;
                refreshSmtpHint();
            });
        }
    }

    function init() {
        mountDetail();
        initMembershipReview();
        readLists();
        ensureActiveKey();
        // Gespeichertes Alias-Schema in die UI laden
        if (typeof window.ms365GetClassNickSchema === 'function') {
            const schema = window.ms365GetClassNickSchema();
            const prefixEl = document.getElementById('jgNickPrefix');
            const patternEl = document.getElementById('jgNickPattern');
            const upperEl = document.getElementById('jgNickUpper');
            if (prefixEl) prefixEl.value = schema.prefix || 'jg';
            if (patternEl) patternEl.value = schema.pattern || '{prefix}{year}-{suffix}';
            if (upperEl) upperEl.checked = !!schema.upper;
        }
        wire();
        syncJgCreateTeamUi();
        // Nomenklatur-Vorschau initialisieren
        const nickPreviewEl = document.getElementById('jgNickPreview');
        if (nickPreviewEl && typeof window.ms365GetClassNickSchema === 'function') {
            const s = window.ms365GetClassNickSchema();
            const prefix = (s.prefix || 'jg').toLowerCase().replace(/[^a-z0-9]/g,'') || 'jg';
            const pattern = s.pattern || '{prefix}{year}-{suffix}';
            const exSuffix = s.upper ? 'AK' : 'ak';
            const raw = pattern.replaceAll('{prefix}',prefix).replaceAll('{year}','2031')
                .replaceAll('{suffix}',exSuffix).replaceAll('{klasse}',s.upper?'1AK':'1ak').replaceAll('{kv}',s.upper?'SCW':'scw');
            nickPreviewEl.textContent = raw.trim().replace(/\s+/g,'-').replace(/[^a-zA-Z0-9-]/g,'').replace(/-+/g,'-').replace(/^-|-$/g,'').toLowerCase() || 'jg2031-ak';
        }
        renderLeftList();
        applyCreateDefaults();
        gd().setTab('general');
        refreshMatchUi();
        if (getActiveGroupId()) live().loadGroup({ silent: true });
    }

    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', init);
    } else {
        init();
    }
})();
