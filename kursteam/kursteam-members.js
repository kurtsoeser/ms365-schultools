(function () {
    'use strict';

    const ns = (window.ms365Kursteam = window.ms365Kursteam || {});

    const GRAPH_SCOPES_MEMBERS = [
        'https://graph.microsoft.com/User.Read',
        'https://graph.microsoft.com/Group.ReadWrite.All',
        'https://graph.microsoft.com/User.Read.All'
    ];

    // ---------------------------
    // State
    // ---------------------------
    ns.studentRosterRaw = ns.studentRosterRaw || '';
    ns.studentRoster = ns.studentRoster || {
        // klasseKey -> Set<upn>
        byClass: {},
        // klasseKey|gruppeKey -> Set<upn>
        byClassGroup: {},
        stats: { validLines: 0, classCount: 0, classGroupCount: 0 }
    };

    function normToken(s) {
        return String(s || '').trim().toUpperCase();
    }

    function normUpn(s) {
        return String(s || '').trim().toLowerCase();
    }

    function splitLineSmart(line) {
        const t = String(line || '').trim();
        if (!t) return [];
        if (t.startsWith('#')) return [];
        if (t.includes(';')) return t.split(/\s*;\s*/).map(x => x.trim()).filter(Boolean);
        if (t.includes('\t')) return t.split(/\t+/).map(x => x.trim()).filter(Boolean);
        if (t.includes('|')) return t.split(/\s*\|\s*/).map(x => x.trim()).filter(Boolean);
        if (/\s{2,}/.test(t)) return t.split(/\s{2,}/).map(x => x.trim()).filter(Boolean);
        return t.split(/\s+/).map(x => x.trim()).filter(Boolean);
    }

    function resetRoster() {
        ns.studentRoster = {
            byClass: {},
            byClassGroup: {},
            stats: { validLines: 0, classCount: 0, classGroupCount: 0 }
        };
    }

    ns.studentRosterTeamSelection = ns.studentRosterTeamSelection || {};

    function ensureSet(map, key) {
        if (!map[key]) map[key] = new Set();
        return map[key];
    }

    function addRosterEntry(klasseRaw, gruppeRaw, upnRaw) {
        const klasse = normToken(klasseRaw);
        const gruppe = normToken(gruppeRaw);
        const upn = normUpn(upnRaw);
        if (!klasse || !upn) return false;

        if (gruppe) {
            ensureSet(ns.studentRoster.byClassGroup, klasse + '|' + gruppe).add(upn);
        } else {
            ensureSet(ns.studentRoster.byClass, klasse).add(upn);
        }
        return true;
    }

    ns.parseStudentRosterFromText = function parseStudentRosterFromText(text) {
        resetRoster();
        const lines = String(text || '').split(/\r?\n/);
        let valid = 0;

        lines.forEach(line => {
            const parts = splitLineSmart(line);
            if (!parts.length) return;
            // 2 Felder: Klasse;UPN
            if (parts.length === 2) {
                if (addRosterEntry(parts[0], '', parts[1])) valid++;
                return;
            }
            // 3+ Felder: Klasse;Gruppe;UPN (Rest wieder zusammen)
            if (parts.length >= 3) {
                const klasse = parts[0];
                const gruppe = parts[1];
                const upn = parts.slice(2).join(' ').trim();
                if (addRosterEntry(klasse, gruppe, upn)) valid++;
                return;
            }
        });

        const classes = Object.keys(ns.studentRoster.byClass);
        const classGroups = Object.keys(ns.studentRoster.byClassGroup);
        ns.studentRoster.stats = {
            validLines: valid,
            classCount: classes.length,
            classGroupCount: classGroups.length
        };
    };

    // ---------------------------
    // File Import (CSV/XLSX/XLS)
    // ---------------------------
    function normalizeImportedRowKeys(row) {
        if (typeof ns.normalizeImportedRowKeys === 'function') return ns.normalizeImportedRowKeys(row);
        const out = {};
        Object.keys(row || {}).forEach(k => {
            const nk = String(k || '').replace(/^\uFEFF/, '').trim();
            out[nk] = row[k];
        });
        return out;
    }

    function pickFirst(row, keys) {
        for (const k of keys) {
            if (row[k] !== undefined && row[k] !== null && String(row[k]).trim() !== '') return row[k];
        }
        return '';
    }

    function parseStudentRosterRows(rows) {
        resetRoster();
        let valid = 0;
        (rows || []).forEach(orig => {
            const row = normalizeImportedRowKeys(orig || {});
            const klasse = pickFirst(row, ['Klasse', 'klasse', 'Class', 'class']);
            const gruppe = pickFirst(row, ['Gruppe', 'gruppe', 'Group', 'group', 'Schülergruppe', 'Schuelergruppe', 'Schüler-Gruppe']);
            const upn = pickFirst(row, ['UPN', 'upn', 'UserPrincipalName', 'userPrincipalName', 'E-Mail', 'Email', 'email', 'Mail', 'mail']);
            if (addRosterEntry(klasse, gruppe, upn)) valid++;
        });
        ns.studentRoster.stats = {
            validLines: valid,
            classCount: Object.keys(ns.studentRoster.byClass).length,
            classGroupCount: Object.keys(ns.studentRoster.byClassGroup).length
        };
    }

    function handleRosterFile(file) {
        const reader = new FileReader();
        reader.onload = e => {
            try {
                const name = (file.name || '').toLowerCase();
                let jsonData;
                if (name.endsWith('.csv')) {
                    const buf = new Uint8Array(e.target.result);
                    let text = new TextDecoder('utf-8').decode(buf);
                    if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
                    let workbook = XLSX.read(text, { type: 'string', FS: ';' });
                    let firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    jsonData = XLSX.utils.sheet_to_json(firstSheet);
                    if (!jsonData.length || Object.keys(jsonData[0] || {}).length < 2) {
                        const wb2 = XLSX.read(text, { type: 'string', FS: ',' });
                        const sh2 = wb2.Sheets[wb2.SheetNames[0]];
                        const j2 = XLSX.utils.sheet_to_json(sh2);
                        if (j2.length) jsonData = j2;
                    }
                } else {
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                    jsonData = XLSX.utils.sheet_to_json(firstSheet);
                }
                parseStudentRosterRows(jsonData);
                ns.studentRosterRaw = ''; // Datei ersetzt Paste-Rohtext
                ns.refreshStudentRosterUI();
                ns.showToast('Schülerliste importiert.');
            } catch (error) {
                ns.showToast('Fehler beim Lesen der Schülerliste: ' + error.message);
            }
        };
        reader.readAsArrayBuffer(file);
    }

    // ---------------------------
    // Mapping Teams -> Members
    // ---------------------------
    function getBool(id, fallback) {
        const el = document.getElementById(id);
        if (!el) return !!fallback;
        return !!el.checked;
    }

    function getMembersForTeam(team) {
        const preferGroup = getBool('studentRosterPreferGroup', true);
        const skipCombined = getBool('studentRosterSkipCombinedClasses', true);

        const klasseRaw = team && (team.originalClass || team.klasseForMembers || '');
        const gruppeRaw = team && (team.gruppe || '');

        if (!klasseRaw) return { members: [], reason: 'no_class' };

        if (skipCombined && String(klasseRaw).includes(',')) {
            return { members: [], reason: 'combined_class' };
        }

        const klasse = normToken(klasseRaw);
        const gruppe = normToken(gruppeRaw);

        if (preferGroup && gruppe) {
            const key = klasse + '|' + gruppe;
            const set = ns.studentRoster.byClassGroup[key];
            if (set && set.size) return { members: Array.from(set), reason: 'class_group' };
        }

        const set2 = ns.studentRoster.byClass[klasse];
        if (set2 && set2.size) return { members: Array.from(set2), reason: 'class' };

        return { members: [], reason: 'no_match' };
    }

    // ---------------------------
    // PowerShell generator
    // ---------------------------
    function buildAddMembersPs1(validTeams) {
        const stamp = new Date().toISOString();

        const assignments = [];
        validTeams.forEach(t => {
            const { members } = getMembersForTeam(t);
            if (!members.length) return;
            const key = String(t.gruppenmail || '').trim();
            const sel = ns.studentRosterTeamSelection && key ? ns.studentRosterTeamSelection[key] !== false : true;
            if (!sel) return;
            assignments.push({
                gruppenmail: t.gruppenmail,
                teamName: t.teamName,
                members
            });
        });

        const lines = [];
        lines.push('#Requires -Version 5.1');
        lines.push('# Kursteam: Schüler als Mitglieder hinzufügen (Microsoft.Graph)');
        lines.push('# Erzeugt in der Browser-App am ' + stamp);
        lines.push('');
        lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
        lines.push('$ErrorActionPreference = "Continue"');
        lines.push('');
        lines.push('if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {');
        lines.push('    Write-Host "Installiere Modul Microsoft.Graph (einmalig)..." -ForegroundColor Yellow');
        lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force');
        lines.push('}');
        lines.push('');
        lines.push('Import-Module Microsoft.Graph -ErrorAction Stop');
        lines.push('');
        lines.push('Write-Host "=== Anmeldung (Microsoft Graph) ===" -ForegroundColor Cyan');
        lines.push('Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All"');
        lines.push('Select-MgProfile -Name "v1.0"');
        lines.push('');
        lines.push('$Assignments = @(');
        lines.push(
            assignments
                .map(a => {
                    const ms = a.members
                        .map(u => "        '" + ns.psEscapeSingle(u) + "'")
                        .join(',\r\n');
                    return (
                        "    [PSCustomObject]@{ Gruppenmail = '" +
                        ns.psEscapeSingle(a.gruppenmail) +
                        "'; TeamName = '" +
                        ns.psEscapeSingle(a.teamName) +
                        "'; Members = @(\r\n" +
                        ms +
                        '\r\n    ) }'
                    );
                })
                .join(',\r\n')
        );
        lines.push(')');
        lines.push('');
        lines.push('if ($Assignments.Count -eq 0) {');
        lines.push('    Write-Host "Keine passenden Schüler-Zuordnungen gefunden (prüfen: Klasse/Gruppe in der Liste)." -ForegroundColor Yellow');
        lines.push('    return');
        lines.push('}');
        lines.push('');
        lines.push('$i = 0');
        lines.push('foreach ($A in $Assignments) {');
        lines.push('    $i++');
        lines.push('    Write-Host ("[{0}/{1}] {2}" -f $i, $Assignments.Count, $A.TeamName) -ForegroundColor Cyan');
        lines.push('    try {');
        lines.push('        $g = Get-MgGroup -Filter ("mailNickname eq ''{0}''" -f $A.Gruppenmail) -ConsistencyLevel eventual');
        lines.push('        if ($null -eq $g) { Write-Warning ("  Gruppe nicht gefunden: {0}" -f $A.Gruppenmail); continue }');
        lines.push('    } catch {');
        lines.push('        Write-Warning ("  Fehler beim Suchen der Gruppe {0}: {1}" -f $A.Gruppenmail, $_.Exception.Message)');
        lines.push('        continue');
        lines.push('    }');
        lines.push('');
        lines.push('    $added = 0');
        lines.push('    foreach ($upn in $A.Members) {');
        lines.push('        if ([string]::IsNullOrWhiteSpace($upn)) { continue }');
        lines.push('        try {');
        lines.push('            $u = Get-MgUser -UserId $upn -ErrorAction Stop');
        lines.push('            New-MgGroupMemberByRef -GroupId $g.Id -BodyParameter @{ "@odata.id" = ("https://graph.microsoft.com/v1.0/directoryObjects/{0}" -f $u.Id) } -ErrorAction Stop | Out-Null');
        lines.push('            $added++');
        lines.push('        } catch {');
        lines.push('            $m = $_.Exception.Message');
        lines.push('            if ($m -match "added object references already exist" -or $m -match "already exist" -or $m -match "already exists") {');
        lines.push('                continue');
        lines.push('            }');
        lines.push('            Write-Warning ("  Mitglied {0}: {1}" -f $upn, $m)');
        lines.push('        }');
        lines.push('        Start-Sleep -Milliseconds 250');
        lines.push('    }');
        lines.push('    Write-Host ("  OK: {0} hinzugefügt (Duplikate ignoriert)" -f $added) -ForegroundColor Green');
        lines.push('    Start-Sleep -Seconds 2');
        lines.push('}');
        lines.push('');
        lines.push('Write-Host "Fertig." -ForegroundColor Cyan');
        lines.push('Read-Host "Enter druecken zum Beenden"');

        return lines.join('\r\n');
    }

    // ---------------------------
    // Online (Graph im Browser)
    // ---------------------------
    let msalMod = null;
    let pca = null;

    async function loadMsal() {
        if (msalMod) return msalMod;
        try {
            msalMod = await import('https://esm.sh/@azure/msal-browser@3.26.1');
        } catch {
            msalMod = await import('https://cdn.jsdelivr.net/npm/@azure/msal-browser@3.26.1/+esm');
        }
        return msalMod;
    }

    function isInteractionRequired(e) {
        return (
            e &&
            (e.name === 'InteractionRequiredAuthError' ||
                e.errorCode === 'interaction_required' ||
                (typeof e.message === 'string' && e.message.indexOf('interaction_required') !== -1))
        );
    }

    function resolveMsalConfig() {
        let cfg = window.MS365_MSAL_CONFIG;
        if (!cfg) cfg = {};
        let id = String(cfg.clientId || '').trim();
        if (!id) {
            const meta = document.querySelector('meta[name="ms365-graph-client-id"]');
            const fromMeta = meta && meta.getAttribute('content') ? meta.getAttribute('content').trim() : '';
            if (fromMeta) id = fromMeta;
        }
        if (!id) {
            throw new Error('Keine clientId (ms365-config.js / meta ms365-graph-client-id).');
        }
        return {
            clientId: id,
            authority: cfg.authority || 'https://login.microsoftonline.com/organizations',
            redirectUri: (cfg.redirectUri || window.location.href.split('#')[0]).trim()
        };
    }

    async function getPca() {
        const m = await loadMsal();
        const PublicClientApplication = m.PublicClientApplication || (m.default && m.default.PublicClientApplication);
        if (!PublicClientApplication) {
            throw new Error('MSAL: PublicClientApplication nicht gefunden (Import).');
        }
        const cfg = resolveMsalConfig();
        if (!pca) {
            pca = new PublicClientApplication({
                auth: {
                    clientId: cfg.clientId,
                    authority: cfg.authority,
                    redirectUri: cfg.redirectUri
                },
                cache: {
                    cacheLocation: 'sessionStorage',
                    storeAuthStateInCookie: true
                }
            });
            await pca.initialize();
            await pca.handleRedirectPromise();
        }
        return pca;
    }

    async function getGraphToken() {
        const instance = await getPca();
        let accounts = instance.getAllAccounts();
        if (!accounts.length) {
            await instance.loginPopup({ scopes: GRAPH_SCOPES_MEMBERS, prompt: 'select_account' });
            accounts = instance.getAllAccounts();
        }
        if (!accounts.length) throw new Error('Anmeldung abgebrochen.');
        const req = { scopes: GRAPH_SCOPES_MEMBERS, account: accounts[0] };
        try {
            return (await instance.acquireTokenSilent(req)).accessToken;
        } catch (e) {
            if (isInteractionRequired(e)) {
                return (await instance.acquireTokenPopup(req)).accessToken;
            }
            throw e;
        }
    }

    function sleep(ms) {
        return new Promise(r => setTimeout(r, ms));
    }

    async function graphRequest(method, path, token, body) {
        const url = path.indexOf('http') === 0 ? path : 'https://graph.microsoft.com/v1.0' + path;
        let attempt = 0;
        while (true) {
            const headers = { Authorization: 'Bearer ' + token };
            if (body !== undefined) headers['Content-Type'] = 'application/json';
            const res = await fetch(url, {
                method,
                headers,
                body: body !== undefined ? JSON.stringify(body) : undefined
            });
            if (res.status === 429 && attempt < 8) {
                const ra = parseInt(res.headers.get('Retry-After') || '5', 10);
                await sleep((isNaN(ra) ? 5 : ra) * 1000);
                attempt++;
                continue;
            }
            return res;
        }
    }

    async function graphJson(method, path, token, body) {
        const res = await graphRequest(method, path, token, body);
        const text = await res.text();
        let data = null;
        if (text) {
            try {
                data = JSON.parse(text);
            } catch {
                data = text;
            }
        }
        if (!res.ok) {
            const msg =
                typeof data === 'object' && data && data.error ? JSON.stringify(data.error) : text || String(res.status);
            throw new Error(method + ' ' + path + ': ' + msg);
        }
        return data || {};
    }

    function isGraphDuplicateRefError(err) {
        const msg = String(err && err.message ? err.message : err);
        return /added object references already exist/i.test(msg) || /already exist/i.test(msg) || /already exists/i.test(msg);
    }

    function appendOnlineLog(msg, kind) {
        const el = document.getElementById('studentRosterOnlineLog');
        if (!el) return;
        const line = document.createElement('div');
        line.textContent = new Date().toLocaleTimeString() + '  ' + msg;
        if (kind === 'err') line.style.color = '#b00020';
        else if (kind === 'ok') line.style.color = '#0d8050';
        else if (kind === 'warn') line.style.color = '#856404';
        else line.style.color = '#212529';
        el.appendChild(line);
        el.scrollTop = el.scrollHeight;
    }

    function clearOnlineLog() {
        const el = document.getElementById('studentRosterOnlineLog');
        if (el) el.replaceChildren();
    }

    function buildAssignmentsForOnline(validTeams) {
        const assignments = [];
        validTeams.forEach(t => {
            const { members } = getMembersForTeam(t);
            if (!members.length) return;
            const key = String(t.gruppenmail || '').trim();
            const sel = ns.studentRosterTeamSelection && key ? ns.studentRosterTeamSelection[key] !== false : true;
            if (!sel) return;
            assignments.push({
                gruppenmail: t.gruppenmail,
                teamName: t.teamName,
                members
            });
        });
        return assignments;
    }

    async function runMembersOnline() {
        const btnLogin = document.getElementById('studentRosterOnlineLogin');
        const btnRun = document.getElementById('studentRosterOnlineRun');

        const rosterStats = ns.studentRoster && ns.studentRoster.stats ? ns.studentRoster.stats : { validLines: 0 };
        if (!rosterStats.validLines) {
            ns.showToast('Bitte zuerst eine Schülerliste übernehmen oder importieren.');
            return;
        }

        const validTeams = (ns.teamsData || []).filter(t => t && t.isValid);
        if (!validTeams.length) {
            ns.showToast('Keine gültigen Teams – zuerst Team-Namen generieren.');
            return;
        }

        const assignments = buildAssignmentsForOnline(validTeams);
        if (!assignments.length) {
            ns.showToast('Keine passenden Zuordnungen gefunden (Klasse/Gruppe prüfen).');
            return;
        }

        if (btnRun) btnRun.disabled = true;
        if (btnLogin) btnLogin.disabled = true;
        clearOnlineLog();
        appendOnlineLog('Start – Schüler als Mitglieder hinzufügen …');

        let token;
        try {
            token = await getGraphToken();
        } catch (e) {
            appendOnlineLog('Anmeldung/Token: ' + (e.message || e), 'err');
            if (btnRun) btnRun.disabled = false;
            if (btnLogin) btnLogin.disabled = false;
            return;
        }

        let aIdx = 0;
        for (const a of assignments) {
            aIdx++;
            try {
                appendOnlineLog('[' + aIdx + '/' + assignments.length + '] ' + a.teamName + ' …');

                // Gruppe per mailNickname finden
                const q = "/groups?$select=id,displayName,mailNickname&$filter=" + encodeURIComponent("mailNickname eq '" + a.gruppenmail + "'");
                const res = await graphJson('GET', q, token, undefined);
                const g = res && res.value && res.value.length ? res.value[0] : null;
                if (!g) {
                    appendOnlineLog('  Gruppe nicht gefunden: ' + a.gruppenmail, 'err');
                    continue;
                }

                let added = 0;
                let skipped = 0;
                for (const upn of a.members) {
                    if (!upn) continue;
                    try {
                        const u = await graphJson('GET', '/users/' + encodeURIComponent(upn) + '?$select=id', token, undefined);
                        await graphJson('POST', '/groups/' + g.id + '/members/$ref', token, {
                            '@odata.id': 'https://graph.microsoft.com/v1.0/directoryObjects/' + u.id
                        });
                        added++;
                    } catch (e) {
                        if (isGraphDuplicateRefError(e)) {
                            skipped++;
                        } else {
                            appendOnlineLog('  ' + upn + ': ' + (e.message || e), 'warn');
                        }
                    }
                    await sleep(250);
                }

                appendOnlineLog('  OK: ' + added + ' hinzugefügt, ' + skipped + ' Duplikat(e) ignoriert.', 'ok');
            } catch (e) {
                appendOnlineLog('Fehler: ' + (e.message || e), 'err');
            }

            await sleep(1200);
            try {
                token = await getGraphToken();
            } catch (e) {
                appendOnlineLog('Token erneuern: ' + (e.message || e), 'err');
                break;
            }
        }

        appendOnlineLog('Fertig.', 'ok');
        if (btnRun) btnRun.disabled = false;
        if (btnLogin) btnLogin.disabled = false;
    }

    async function loginMembersOnly() {
        const btnLogin = document.getElementById('studentRosterOnlineLogin');
        if (btnLogin) btnLogin.disabled = true;
        try {
            await getGraphToken();
            ns.showToast('Microsoft angemeldet – Sie können jetzt Schüler hinzufügen.');
        } catch (e) {
            ns.showToast('Anmeldung: ' + (e.message || e));
        } finally {
            if (btnLogin) btnLogin.disabled = false;
        }
    }

    ns.refreshStudentRosterUI = function refreshStudentRosterUI() {
        const ta = document.getElementById('studentRosterPasteInput');
        if (ta && ns.studentRosterRaw) {
            if (String(ta.value || '').trim() === '') ta.value = ns.studentRosterRaw;
        }

        const s = ns.studentRoster && ns.studentRoster.stats ? ns.studentRoster.stats : { validLines: 0, classCount: 0, classGroupCount: 0 };
        const v = document.getElementById('studentRosterValidLines');
        const c = document.getElementById('studentRosterClassCount');
        const cg = document.getElementById('studentRosterClassGroupCount');
        if (v) v.textContent = String(s.validLines || 0);
        if (c) c.textContent = String(s.classCount || 0);
        if (cg) cg.textContent = String(s.classGroupCount || 0);

        const pre = document.getElementById('studentRosterPowerShellScript');
        if (pre) {
            const validTeams = (ns.teamsData || []).filter(t => t && t.isValid);
            if (!s.validLines) {
                pre.textContent = '# Sobald eine Schülerliste übernommen wurde, erscheint hier das Script.';
            } else if (!validTeams.length) {
                pre.textContent = '# Keine gültigen Teams – zuerst Team-Namen generieren.';
            } else {
                // Script nur für ausgewählte Teams erzeugen (Teams ohne Treffer werden ohnehin rausgefiltert)
                pre.textContent = buildAddMembersPs1(validTeams);
            }
        }

        if (typeof ns.renderStudentRosterTeamsSelection === 'function') {
            ns.renderStudentRosterTeamsSelection();
        }
    };

    function computeMatchedMembersCount(team) {
        const { members } = getMembersForTeam(team);
        return members ? members.length : 0;
    }

    ns.renderStudentRosterTeamsSelection = function renderStudentRosterTeamsSelection() {
        const body = document.getElementById('studentRosterTeamsBody');
        if (!body) return;

        const rosterStats = ns.studentRoster && ns.studentRoster.stats ? ns.studentRoster.stats : { validLines: 0 };
        const validTeams = (ns.teamsData || []).filter(t => t && t.isValid);
        const hideNoMatch = getBool('studentRosterHideNoMatch', true);

        if (!validTeams.length || !rosterStats.validLines) {
            body.replaceChildren();
            const tr = document.createElement('tr');
            const td = document.createElement('td');
            td.colSpan = 4;
            td.style.color = '#6c757d';
            td.textContent = 'Noch keine Teams oder Schülerliste – zuerst oben übernehmen.';
            tr.appendChild(td);
            body.appendChild(tr);
            return;
        }

        // Default: "nur passende" => wenn Key noch nie gesetzt wurde, setzen wir es initial auf (matched>0)
        validTeams.forEach(t => {
            const key = String(t.gruppenmail || '').trim();
            if (!key) return;
            if (ns.studentRosterTeamSelection[key] === undefined) {
                ns.studentRosterTeamSelection[key] = computeMatchedMembersCount(t) > 0;
            }
        });

        const rows = validTeams
            .map(t => ({
                team: t,
                key: String(t.gruppenmail || '').trim(),
                matched: computeMatchedMembersCount(t)
            }))
            .filter(r => (hideNoMatch ? r.matched > 0 : true))
            .sort((a, b) => String(a.team.teamName || '').localeCompare(String(b.team.teamName || ''), 'de'));

        body.replaceChildren();
        if (!rows.length) {
            const tr0 = document.createElement('tr');
            const td0 = document.createElement('td');
            td0.colSpan = 4;
            td0.style.color = '#6c757d';
            td0.textContent = hideNoMatch
                ? 'Keine Teams mit Treffern (oder alle ausgeblendet).'
                : 'Keine Teams in der Liste.';
            tr0.appendChild(td0);
            body.appendChild(tr0);
            return;
        }
        rows.forEach(r => {
            const tr = document.createElement('tr');

            const tdSel = document.createElement('td');
            const cb = document.createElement('input');
            cb.type = 'checkbox';
            cb.checked = r.key ? ns.studentRosterTeamSelection[r.key] !== false : false;
            cb.disabled = !r.key;
            cb.addEventListener('change', () => {
                if (!r.key) return;
                ns.studentRosterTeamSelection[r.key] = !!cb.checked;
                ns.refreshStudentRosterUI();
            });
            tdSel.appendChild(cb);

            const tdName = document.createElement('td');
            tdName.textContent = r.team.teamName || '';

            const tdMail = document.createElement('td');
            tdMail.textContent = r.team.gruppenmail || '';

            const tdCnt = document.createElement('td');
            tdCnt.textContent = String(r.matched);
            tdCnt.style.color = r.matched > 0 ? '#0d8050' : '#6c757d';

            tr.append(tdSel, tdName, tdMail, tdCnt);
            body.appendChild(tr);
        });
    };

    function wireOnce() {
        if (wireOnce._wired) return;
        wireOnce._wired = true;

        const btnApply = document.getElementById('btnStudentRosterApply');
        if (btnApply) {
            btnApply.addEventListener('click', () => {
                const ta = document.getElementById('studentRosterPasteInput');
                const text = ta ? ta.value : '';
                ns.studentRosterRaw = String(text || '');
                ns.parseStudentRosterFromText(ns.studentRosterRaw);
                ns.refreshStudentRosterUI();
                ns.showToast('Schülerliste übernommen.');
            });
        }

        const btnCopy = document.getElementById('btnCopyStudentRosterPowerShell');
        if (btnCopy) {
            btnCopy.addEventListener('click', () => {
                const pre = document.getElementById('studentRosterPowerShellScript');
                const text = pre ? pre.textContent : '';
                if (!text || /^#\s*Sobald/.test(String(text).trim())) {
                    ns.showToast('Noch kein Script: zuerst Schülerliste übernehmen.');
                    return;
                }
                navigator.clipboard.writeText(text).then(
                    () => ns.showToast('PowerShell-Script (Schüler hinzufügen) kopiert.'),
                    () => ns.showToast('Kopieren fehlgeschlagen.')
                );
            });
        }

        const prefer = document.getElementById('studentRosterPreferGroup');
        if (prefer) prefer.addEventListener('change', () => ns.refreshStudentRosterUI());
        const skip = document.getElementById('studentRosterSkipCombinedClasses');
        if (skip) skip.addEventListener('change', () => ns.refreshStudentRosterUI());
        const hide = document.getElementById('studentRosterHideNoMatch');
        if (hide) hide.addEventListener('change', () => ns.renderStudentRosterTeamsSelection());

        const selAll = document.getElementById('studentRosterSelectAll');
        if (selAll) {
            selAll.addEventListener('click', () => {
                const validTeams = (ns.teamsData || []).filter(t => t && t.isValid);
                validTeams.forEach(t => {
                    const key = String(t.gruppenmail || '').trim();
                    if (key) ns.studentRosterTeamSelection[key] = true;
                });
                ns.refreshStudentRosterUI();
            });
        }
        const selNone = document.getElementById('studentRosterSelectNone');
        if (selNone) {
            selNone.addEventListener('click', () => {
                const validTeams = (ns.teamsData || []).filter(t => t && t.isValid);
                validTeams.forEach(t => {
                    const key = String(t.gruppenmail || '').trim();
                    if (key) ns.studentRosterTeamSelection[key] = false;
                });
                ns.refreshStudentRosterUI();
            });
        }
        const selMatched = document.getElementById('studentRosterSelectMatched');
        if (selMatched) {
            selMatched.addEventListener('click', () => {
                const validTeams = (ns.teamsData || []).filter(t => t && t.isValid);
                validTeams.forEach(t => {
                    const key = String(t.gruppenmail || '').trim();
                    if (!key) return;
                    ns.studentRosterTeamSelection[key] = computeMatchedMembersCount(t) > 0;
                });
                ns.refreshStudentRosterUI();
            });
        }
        const selInv = document.getElementById('studentRosterSelectInvert');
        if (selInv) {
            selInv.addEventListener('click', () => {
                const validTeams = (ns.teamsData || []).filter(t => t && t.isValid);
                validTeams.forEach(t => {
                    const key = String(t.gruppenmail || '').trim();
                    if (!key) return;
                    const cur = ns.studentRosterTeamSelection[key] !== false;
                    ns.studentRosterTeamSelection[key] = !cur;
                });
                ns.refreshStudentRosterUI();
            });
        }

        const area = document.getElementById('studentRosterUploadArea');
        const input = document.getElementById('studentRosterFileInput');
        if (area && input) {
            area.addEventListener('click', () => input.click());
            area.addEventListener('dragover', (e) => {
                e.preventDefault();
                area.classList.add('dragover');
            });
            area.addEventListener('dragleave', () => area.classList.remove('dragover'));
            area.addEventListener('drop', (e) => {
                e.preventDefault();
                area.classList.remove('dragover');
                if (e.dataTransfer.files.length > 0) handleRosterFile(e.dataTransfer.files[0]);
            });
            input.addEventListener('change', (e) => {
                if (e.target.files.length > 0) handleRosterFile(e.target.files[0]);
            });
        }
    }

    window.ms365KursteamMembersGraphLogin = loginMembersOnly;
    window.ms365KursteamMembersGraphRun = runMembersOnline;

    // Wire when script loads (DOM is usually ready, but guard anyway)
    if (document.readyState === 'loading') {
        document.addEventListener('DOMContentLoaded', wireOnce);
    } else {
        wireOnce();
    }
})();

