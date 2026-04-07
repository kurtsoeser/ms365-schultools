(function () {
    'use strict';

    let argeCurrentStep = 1;
    /** @type {{ displayName: string, mailNick: string, owner: string, description: string }[]} */
    let argeRows = [];

    const panelW = document.getElementById('panelWebuntis');
    const panelJ = document.getElementById('panelJahrgang');
    const panelA = document.getElementById('panelArge');

    const btnModeW = document.getElementById('modeWebuntis');
    const btnModeJ = document.getElementById('modeJahrgang');
    const btnModeA = document.getElementById('modeArge');

    function showToast(msg) {
        const el = document.getElementById('toast');
        if (!el) return;
        el.textContent = msg;
        el.classList.add('show');
        clearTimeout(showToast._t);
        showToast._t = setTimeout(() => el.classList.remove('show'), 3500);
    }

    function setMode(which) {
        const w = which === 'webuntis';
        const j = which === 'jahrgang';
        const a = which === 'arge';
        if (panelW) panelW.style.display = w ? '' : 'none';
        if (panelJ) panelJ.style.display = j ? '' : 'none';
        if (panelA) panelA.style.display = a ? '' : 'none';
        if (btnModeW) btnModeW.classList.toggle('btn-success', w);
        if (btnModeJ) btnModeJ.classList.toggle('btn-success', j);
        if (btnModeA) btnModeA.classList.toggle('btn-success', a);
        if (a) {
            scheduleArgePreviewRefresh();
        }
    }

    if (btnModeW) btnModeW.addEventListener('click', () => setMode('webuntis'));
    if (btnModeJ) btnModeJ.addEventListener('click', () => setMode('jahrgang'));
    if (btnModeA) btnModeA.addEventListener('click', () => setMode('arge'));

    function applyInitialModeFromUrl() {
        try {
            const mode = new URLSearchParams(window.location.search).get('mode');
            if (!mode) return;
            if (mode.toLowerCase() === 'arge') setMode('arge');
        } catch {
            // ignore
        }
    }

    function argeStepNum(el) {
        const raw = el.getAttribute('data-arge-step');
        const n = parseFloat(String(raw || '').trim());
        return Number.isFinite(n) ? n : NaN;
    }

    function goToArgeStep(step) {
        argeCurrentStep = step;
        document.querySelectorAll('.arge-step-content').forEach(el => {
            el.classList.toggle('active', argeStepNum(el) === step);
        });
        document.querySelectorAll('.arge-steps .step').forEach(el => {
            const s = argeStepNum(el);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
        if (step === 1) {
            scheduleArgePreviewRefresh();
        }
    }

    function getDomain() {
        if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
            return window.ms365GetSchoolDomainNoAt();
        }
        return '';
    }

    function getPrefix() {
        const el = document.getElementById('argeDefaultPrefix');
        const raw = (el && el.value ? el.value : '').trim();
        if (!raw) return '';
        return raw.toLowerCase().replace(/[^a-z0-9]/g, '');
    }

    function toNickBaseFromName(displayName) {
        // sehr einfache Normalisierung (ASCII-ish)
        let s = String(displayName || '').trim();
        s = s.replace(/[äÄ]/g, 'ae').replace(/[öÖ]/g, 'oe').replace(/[üÜ]/g, 'ue').replace(/ß/g, 'ss');
        s = s.replace(/[^A-Za-z0-9]+/g, '-').replace(/-+/g, '-');
        s = s.replace(/^-+|-+$/g, '');
        return s;
    }

    /** Fach-Teil ohne führendes „ARGE “ – für Slug/Mail-Nickname */
    function subjectForSlug(line) {
        let t = String(line || '').trim();
        const stripped = t.replace(/^ARGE\s+/i, '').trim();
        return stripped || t;
    }

    /** Anzeigename der M365-Gruppe aus einer einfachen Fach-Zeile */
    function displayNameFromSubjectLine(line) {
        const t = line.trim();
        if (!t) return '';
        if (/^ARGE\s+/i.test(t)) return t;
        return 'ARGE ' + t;
    }

    function maybeUpper(s) {
        const el = document.getElementById('argeUpperNick');
        const upper = el ? !!el.checked : false;
        return upper ? s.toUpperCase() : s.toLowerCase();
    }

    /** Mail-Nickname nur aus dem Fach (Präfix aus Einstellungen), nicht aus „ARGE …“ doppelt */
    function buildMailNicknameFromSubject(line) {
        const base = toNickBaseFromName(subjectForSlug(line));
        if (!base) return '';
        const pre = getPrefix();
        const combined = pre ? pre + '-' + base : base;
        return maybeUpper(combined).replace(/[^A-Za-z0-9-]/g, '');
    }

    /**
     * Parst die Textarea: eine Zeile pro Fach oder optional Anzeigename;MailNickname.
     * @returns {{ parsed: { displayName: string, mailNick: string, owner: string, description: string, technicalSlug: string }[], errors: string[] }}
     */
    function parseArgeInput() {
        const ta = document.getElementById('argeLines');
        if (!ta) {
            return { parsed: [], errors: [] };
        }
        const lines = ta.value.split(/\r\n|\n|\r/);
        const parsed = [];
        const errors = [];
        const seen = new Set();
        lines.forEach((line, idx) => {
            const t = line.trim();
            if (!t || t.startsWith('#')) return;
            const parts = t.split(/[;\t]/).map(x => x.trim()).filter(Boolean);
            if (!parts.length) return;

            let displayName;
            let mailNick;
            let technicalSlug;

            if (parts.length >= 2) {
                displayName = parts[0];
                const explicitNick = parts[1] || '';
                technicalSlug = toNickBaseFromName(subjectForSlug(parts[0]));
                mailNick = explicitNick
                    ? maybeUpper(explicitNick.replace(/[^A-Za-z0-9-]/g, ''))
                    : buildMailNicknameFromSubject(parts[0]);
            } else {
                const raw = parts[0];
                displayName = displayNameFromSubjectLine(raw);
                technicalSlug = toNickBaseFromName(subjectForSlug(raw));
                mailNick = buildMailNicknameFromSubject(raw);
            }

            if (!displayName) return;
            if (!mailNick) {
                errors.push('Zeile ' + (idx + 1) + ': Mail-Nickname konnte nicht erzeugt werden.');
                return;
            }
            const key = displayName.toLowerCase();
            if (seen.has(key)) return;
            seen.add(key);
            parsed.push({
                displayName,
                mailNick,
                owner: '',
                description: 'ARGE-Gruppe: ' + displayName,
                technicalSlug
            });
        });
        return { parsed, errors };
    }

    let argePreviewDebounce;
    function scheduleArgePreviewRefresh() {
        clearTimeout(argePreviewDebounce);
        argePreviewDebounce = setTimeout(refreshArgePreview, 120);
    }

    function refreshArgePreview() {
        const tbody = document.getElementById('argePreviewBody');
        if (!tbody) return;
        try {
            const { parsed } = parseArgeInput();
            const rows = parsed.map(r => ({ ...r }));
            resolveDuplicateNicks(rows);

            if (!rows.length) {
                tbody.innerHTML =
                    '<tr><td colspan="4" style="color:#6c757d;">Noch keine Zeilen – oben Fächer einfügen.</td></tr>';
                return;
            }

            const domain = getDomain() || '…';
            tbody.replaceChildren();
            rows.forEach(r => {
                const tr = document.createElement('tr');
                const tech = r.technicalSlug || toNickBaseFromName(subjectForSlug(r.displayName));
                const td1 = document.createElement('td');
                td1.textContent = r.displayName;
                const td2 = document.createElement('td');
                td2.textContent = tech;
                td2.style.fontFamily = 'Consolas,monospace';
                td2.style.fontSize = '0.9em';
                const td3 = document.createElement('td');
                td3.textContent = r.mailNick;
                td3.style.fontFamily = 'Consolas,monospace';
                td3.style.fontSize = '0.9em';
                const td4 = document.createElement('td');
                td4.textContent = r.mailNick + '@' + domain;
                tr.append(td1, td2, td3, td4);
                tbody.appendChild(tr);
            });
        } catch (e) {
            console.error('ARGE-Vorschau:', e);
            tbody.innerHTML =
                '<tr><td colspan="4" style="color:#dc3545;">Vorschau konnte nicht berechnet werden. Seite neu laden oder Konsole prüfen.</td></tr>';
        }
    }

    function resolveDuplicateNicks(rows) {
        const seen = new Map();
        rows.forEach(r => {
            const base = r.mailNick;
            let candidate = base;
            let n = 2;
            while (seen.has(candidate)) {
                candidate = base + '-' + n;
                n++;
            }
            r.mailNick = candidate;
            seen.set(candidate, true);
        });
    }

    /**
     * Parst die ARGE-Liste neu und übernimmt Besitzer aus dem vorherigen Stand (gleicher Anzeigename).
     * @returns {{ ok: true } | { ok: false, errors: string[] }}
     */
    function syncArgeRowsFromInputPreservingOwners() {
        const { parsed, errors } = parseArgeInput();
        if (errors.length) {
            return { ok: false, errors };
        }
        if (!parsed.length) {
            return { ok: false, errors: ['Bitte mindestens eine ARGE-Zeile eintragen.'] };
        }
        const rows = parsed.map(r => ({ ...r }));
        resolveDuplicateNicks(rows);
        const ownerByKey = new Map(argeRows.map(r => [r.displayName.toLowerCase(), r.owner]));
        argeRows = rows.map(r => ({
            displayName: r.displayName,
            mailNick: r.mailNick,
            owner: ownerByKey.get(r.displayName.toLowerCase()) || '',
            description: r.description
        }));
        rebuildArgeOwnerTableFromRows();
        return { ok: true };
    }

    const ARGE_STORAGE_KEY = 'ms365-arge-state-v2';

    /** Alte Reihenfolge 1=Grundlagen, 2=Liste, 3=Besitzer → neue 1=Liste, 2=Besitzer, 3=Einstellungen */
    function migrateArgeStepFromV1(step) {
        const m = { 1: 3, 2: 1, 3: 2, 4: 4 };
        const n = m[step];
        return n !== undefined ? n : step;
    }

    function getArgeCreateTeams() {
        const el = document.getElementById('argeCreateTeams');
        return el ? !!el.checked : true;
    }

    function getArgeExchangeSmtp() {
        const el = document.getElementById('argeExchangeSmtp');
        return el ? !!el.checked : true;
    }

    function getArgeAdminAsOwner() {
        const el = document.getElementById('argeAdminAsOwner');
        return el ? !!el.checked : true;
    }

    function refreshArgeScriptIfStep4() {
        if (argeCurrentStep !== 4 || !argeRows.length) return;
        const missing = argeRows.filter(r => !r.owner);
        if (missing.length) return;
        const pre = document.getElementById('argePowerShellScript');
        if (pre)
            pre.textContent = buildStandaloneArgePs1(
                false,
                getArgeCreateTeams(),
                getArgeExchangeSmtp(),
                getArgeAdminAsOwner()
            );
    }

    function rebuildArgeOwnerTableFromRows() {
        const domain = getDomain();
        const tbody = document.getElementById('argeOwnerBody');
        tbody.replaceChildren();
        argeRows.forEach((row, index) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = row.displayName;
            const td2 = document.createElement('td');
            td2.textContent = row.mailNick + '@' + domain;
            const td3 = document.createElement('td');
            td3.textContent = row.mailNick;
            const td4 = document.createElement('td');
            const inp = document.createElement('input');
            inp.type = 'email';
            inp.placeholder = 'besitzer@' + domain;
            inp.style.width = '100%';
            inp.style.padding = '8px';
            inp.value = row.owner || '';
            inp.addEventListener('input', () => {
                argeRows[index].owner = inp.value.trim();
            });
            td4.appendChild(inp);
            tr.append(td1, td2, td3, td4);
            tbody.appendChild(tr);
        });
    }

    function saveArgeState() {
        try {
            const state = {
                argeStepOrder: 'v2',
                argeCurrentStep,
                argeRows,
                argeDefaultPrefix: document.getElementById('argeDefaultPrefix').value,
                argeUpperNick: document.getElementById('argeUpperNick').checked,
                argeCreateTeams: getArgeCreateTeams(),
                argeExchangeSmtp: getArgeExchangeSmtp(),
                argeAdminAsOwner: getArgeAdminAsOwner(),
                argeLines: document.getElementById('argeLines').value,
                argePowerShellScript: document.getElementById('argePowerShellScript').textContent
            };
            localStorage.setItem(ARGE_STORAGE_KEY, JSON.stringify(state));
            showToast('ARGEs: Zwischenstand gespeichert.');
        } catch (e) {
            showToast('Speichern fehlgeschlagen: ' + e.message);
        }
    }

    function loadArgeState() {
        try {
            let raw = localStorage.getItem(ARGE_STORAGE_KEY);
            if (!raw) {
                raw = localStorage.getItem('ms365-arge-state-v1');
            }
            if (!raw) {
                showToast('Kein gespeicherter Stand für ARGEs.');
                return;
            }
            const state = JSON.parse(raw);
            let step = typeof state.argeCurrentStep === 'number' ? state.argeCurrentStep : 1;
            if (state.argeStepOrder !== 'v2') {
                step = migrateArgeStepFromV1(step);
            }
            argeCurrentStep = step;
            argeRows = Array.isArray(state.argeRows) ? state.argeRows : [];
            if (
                typeof window.ms365SetSchoolDomainNoAt === 'function' &&
                state.argeDomain !== undefined &&
                String(state.argeDomain).trim() !== ''
            ) {
                window.ms365SetSchoolDomainNoAt(state.argeDomain);
            }
            document.getElementById('argeDefaultPrefix').value =
                state.argeDefaultPrefix !== undefined ? state.argeDefaultPrefix : '';
            document.getElementById('argeUpperNick').checked = !!state.argeUpperNick;
            const argeTeamsEl = document.getElementById('argeCreateTeams');
            if (argeTeamsEl) {
                argeTeamsEl.checked = state.argeCreateTeams !== undefined ? !!state.argeCreateTeams : true;
            }
            const argeExoEl = document.getElementById('argeExchangeSmtp');
            if (argeExoEl) {
                argeExoEl.checked = state.argeExchangeSmtp !== undefined ? !!state.argeExchangeSmtp : true;
            }
            const argeAdminEl = document.getElementById('argeAdminAsOwner');
            if (argeAdminEl) {
                argeAdminEl.checked = state.argeAdminAsOwner !== undefined ? !!state.argeAdminAsOwner : true;
            }
            document.getElementById('argeLines').value = state.argeLines || '';
            document.getElementById('argeParseError').style.display = 'none';
            const pre = document.getElementById('argePowerShellScript');
            if (pre && state.argePowerShellScript !== undefined) {
                pre.textContent = state.argePowerShellScript;
            }
            if (argeRows.length) {
                rebuildArgeOwnerTableFromRows();
            } else {
                document.getElementById('argeOwnerBody').replaceChildren();
            }
            goToArgeStep(Math.min(Math.max(1, argeCurrentStep), 4));
            scheduleArgePreviewRefresh();
            showToast('ARGEs: Stand geladen.');
        } catch (e) {
            showToast('Laden fehlgeschlagen: ' + e.message);
        }
    }

    function clearArgeState() {
        if (!confirm('Gespeicherten Zwischenstand für ARGEs wirklich löschen?')) {
            return;
        }
        try {
            localStorage.removeItem(ARGE_STORAGE_KEY);
            localStorage.removeItem('ms365-arge-state-v1');
            argeCurrentStep = 1;
            argeRows = [];
            document.getElementById('argeDefaultPrefix').value = '';
            document.getElementById('argeUpperNick').checked = false;
            const argeTeamsClear = document.getElementById('argeCreateTeams');
            if (argeTeamsClear) argeTeamsClear.checked = true;
            const argeExoClear = document.getElementById('argeExchangeSmtp');
            if (argeExoClear) argeExoClear.checked = true;
            const argeAdminClear = document.getElementById('argeAdminAsOwner');
            if (argeAdminClear) argeAdminClear.checked = true;
            document.getElementById('argeLines').value = '';
            document.getElementById('argeParseError').style.display = 'none';
            document.getElementById('argeOwnerBody').replaceChildren();
            document.getElementById('argePowerShellScript').textContent = '';
            goToArgeStep(1);
            scheduleArgePreviewRefresh();
            showToast('ARGEs: Speicher geleert.');
        } catch (e) {
            showToast('Fehler: ' + e.message);
        }
    }

    window.ms365SaveArge = saveArgeState;
    window.ms365LoadArge = loadArgeState;
    window.ms365ClearArge = clearArgeState;
    window.ms365ShowToast = showToast;

    /**
     * Snapshot für Online-Ausführung (Microsoft Graph im Browser), siehe arge-graph.js
     * @returns {{ rows: { displayName: string, mailNick: string, owner: string, description: string }[], createTeams: boolean, exchangeSmtp: boolean }}
     */
    window.ms365GetArgeSnapshotForGraph = function () {
        return {
            rows: argeRows.map(function (r) {
                return {
                    displayName: r.displayName,
                    mailNick: r.mailNick,
                    owner: r.owner,
                    description: r.description
                };
            }),
            createTeams: getArgeCreateTeams(),
            exchangeSmtp: getArgeExchangeSmtp(),
            adminAsOwner: getArgeAdminAsOwner()
        };
    };

    function psEscapeSingle(s) {
        return String(s ?? '').replace(/'/g, "''");
    }

    function downloadBlob(filename, text, mime) {
        const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function buildStandaloneArgePs1(standalone, createTeams, setExchangeSmtp, adminAsOwner) {
        if (createTeams === undefined) createTeams = true;
        if (setExchangeSmtp === undefined) setExchangeSmtp = true;
        if (adminAsOwner === undefined) adminAsOwner = true;
        const domain = getDomain();
        const domainTrim = (domain || '').trim();
        const setExoEffective = setExchangeSmtp && domainTrim.length > 0;
        const stamp = new Date().toISOString();
        const lines = [];
        // Team an Gruppe: Graph verlangt i.d.R. nur Group.ReadWrite.All (s. team-put-teams). Team.ReadWrite.All
        // loest bei Connect-MgGraph oft AADSTS70011 (ungueltiger Scope beim Graph-PowerShell-Client).
        const scopesLine =
            '$scopes = @("Group.ReadWrite.All","User.Read.All","User.Read")';

        if (standalone) {
            lines.push('#Requires -Version 5.1');
            lines.push('# ARGE-Gruppen (M365 Unified); optional Teams ($Ms365CreateTeams); optional Exchange-SMTP ($Ms365SetExchangeSmtp)');
            lines.push('# Erzeugt in der Browser-App am ' + stamp);
            lines.push('# Daten sind unten eingebettet.');
            lines.push('');
            lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
            lines.push('$ErrorActionPreference = "Continue"');
            lines.push('');
            lines.push('Write-Host ""');
            lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
            lines.push('Write-Host "  ARGE-Gruppen (Microsoft Graph)"       -ForegroundColor Cyan');
            lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
            lines.push('Write-Host ""');
            lines.push(
                '# Meta-Modul Microsoft.Graph (einheitliche DLL-Versionen; PS-5.1 „4096 Funktionen“ per MaximumFunctionCount)'
            );
            lines.push('$MaximumFunctionCount = 32768');
            lines.push('Write-Host "Lade Microsoft.Graph ..." -ForegroundColor Gray');
            lines.push('try {');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('} catch {');
            lines.push('    Write-Host "Microsoft.Graph nicht gefunden – Installation (einmalig, kann einige Minuten dauern) ..." -ForegroundColor Yellow');
            lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('}');
            lines.push('');
            lines.push(scopesLine);
            lines.push('Write-Host "Starte Microsoft Graph-Anmeldung (Browser/Dialog oder Geraetecode) ..." -ForegroundColor Yellow');
            lines.push('Write-Host "Hinweis: Fenster ggf. im Hintergrund – Taskleiste pruefen." -ForegroundColor Gray');
            lines.push('$script:Ms365OldEap = $ErrorActionPreference');
            lines.push('$ErrorActionPreference = "Stop"');
            lines.push('try {');
            lines.push('    Connect-MgGraph -Scopes $scopes -NoWelcome');
            lines.push('} catch {');
            lines.push('    Write-Host ("Hinweis (interaktive Anmeldung): {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow');
            lines.push('}');
            lines.push('$ErrorActionPreference = $script:Ms365OldEap');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    Write-Host ""');
            lines.push('    Write-Host "Kein Graph-Kontext – Geraetecode-Anmeldung (Code erscheint unten, Browser: https://microsoft.com/devicelogin ) ..." -ForegroundColor Yellow');
            lines.push('    $ErrorActionPreference = "Stop"');
            lines.push('    try {');
            lines.push('        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome');
            lines.push('    } catch {');
            lines.push('        Write-Error ("Microsoft Graph: Anmeldung fehlgeschlagen: {0}" -f $_.Exception.Message)');
            lines.push('        exit 1');
            lines.push('    }');
            lines.push('    $ErrorActionPreference = $script:Ms365OldEap');
            lines.push('}');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    Write-Error "Microsoft Graph: Keine Sitzung – Anmeldung nicht erfolgreich. Skript wird beendet."');
            lines.push('    exit 1');
            lines.push('}');
            lines.push('$mgCtx = Get-MgContext');
            lines.push('Write-Host ("Angemeldet (Tenant: {0})" -f $mgCtx.TenantId) -ForegroundColor Green');
            lines.push('');
        } else {
            lines.push('# Microsoft Graph: ARGE-Gruppen als Microsoft 365-Gruppen (Unified Group, kein Kursteam)');
            lines.push('# Voraussetzung: Install-Module Microsoft.Graph');
            lines.push('# https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup');
            lines.push('');
            lines.push('Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber -ErrorAction SilentlyContinue');
            lines.push('$MaximumFunctionCount = 32768');
            lines.push('try {');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('} catch {');
            lines.push('    Write-Host "Microsoft.Graph nicht gefunden – Installation (einmalig, kann einige Minuten dauern) ..." -ForegroundColor Yellow');
            lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber');
            lines.push('    Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('}');
            lines.push('');
            lines.push(scopesLine);
            lines.push('Write-Host "Starte Microsoft Graph-Anmeldung (Browser/Dialog oder Geraetecode) ..." -ForegroundColor Yellow');
            lines.push('Write-Host "Hinweis: Fenster ggf. im Hintergrund – Taskleiste pruefen." -ForegroundColor Gray');
            lines.push('$script:Ms365OldEap = $ErrorActionPreference');
            lines.push('$ErrorActionPreference = "Stop"');
            lines.push('try {');
            lines.push('    Connect-MgGraph -Scopes $scopes -NoWelcome');
            lines.push('} catch {');
            lines.push('    Write-Host ("Hinweis (interaktive Anmeldung): {0}" -f $_.Exception.Message) -ForegroundColor DarkYellow');
            lines.push('}');
            lines.push('$ErrorActionPreference = $script:Ms365OldEap');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    Write-Host ""');
            lines.push('    Write-Host "Kein Graph-Kontext – Geraetecode-Anmeldung (Code erscheint unten, Browser: https://microsoft.com/devicelogin ) ..." -ForegroundColor Yellow');
            lines.push('    $ErrorActionPreference = "Stop"');
            lines.push('    try {');
            lines.push('        Connect-MgGraph -Scopes $scopes -UseDeviceAuthentication -NoWelcome');
            lines.push('    } catch {');
            lines.push('        throw ("Microsoft Graph: Anmeldung fehlgeschlagen: {0}" -f $_.Exception.Message)');
            lines.push('    }');
            lines.push('    $ErrorActionPreference = $script:Ms365OldEap');
            lines.push('}');
            lines.push('if (-not (Get-MgContext)) {');
            lines.push('    throw "Microsoft Graph: Keine Sitzung – Anmeldung nicht erfolgreich."');
            lines.push('}');
            lines.push('$mgCtx = Get-MgContext');
            lines.push('Write-Host ("Angemeldet (Tenant: {0})" -f $mgCtx.TenantId) -ForegroundColor Green');
            lines.push('');
        }

        lines.push('$Ms365CreateTeams = $' + (createTeams ? 'true' : 'false'));
        lines.push('$Ms365SetExchangeSmtp = $' + (setExoEffective ? 'true' : 'false'));
        lines.push('$Ms365AdminAsOwner = $' + (adminAsOwner ? 'true' : 'false'));
        lines.push("$Ms365ExchangeDomain = '" + psEscapeSingle(domainTrim) + "'");
        lines.push('');
        if (setExoEffective) {
            lines.push('$script:Ms365ExoConnected = $false');
            lines.push('function Ensure-Ms365ExchangeOnline {');
            lines.push('    if ($script:Ms365ExoConnected) { return }');
            lines.push(
                '    Write-Host "Exchange Online: Modul laden und anmelden (zweiter Dialog) …" -ForegroundColor Yellow'
            );
            lines.push('    try {');
            lines.push('        Import-Module ExchangeOnlineManagement -ErrorAction Stop');
            lines.push('    } catch {');
            lines.push('        Write-Host "Installiere ExchangeOnlineManagement …" -ForegroundColor Yellow');
            lines.push('        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber');
            lines.push('        Import-Module ExchangeOnlineManagement -ErrorAction Stop');
            lines.push('    }');
            lines.push('    Connect-ExchangeOnline -ShowBanner:$false');
            lines.push('    $script:Ms365ExoConnected = $true');
            lines.push('    Write-Host "Exchange Online: angemeldet." -ForegroundColor Green');
            lines.push('}');
            lines.push(
                '# Exchange-Anmeldung erst bei Set-UnifiedGroup (nicht vor Graph), damit bei Graph-Fehlern kein zweiter Dialog nötig ist.'
            );
            lines.push('');
        }
        lines.push('$rows = @(');
        argeRows.forEach((r, i) => {
            const last = i === argeRows.length - 1;
            lines.push(
                "    [PSCustomObject]@{ DisplayName = '" +
                    psEscapeSingle(r.displayName) +
                    "'; MailNickname = '" +
                    psEscapeSingle(r.mailNick) +
                    "'; OwnerUpn = '" +
                    psEscapeSingle(r.owner) +
                    "'; Description = '" +
                    psEscapeSingle(r.description) +
                    "' }" + (last ? '' : ',')
            );
        });
        lines.push(')');
        lines.push('$meUser = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/me" -ErrorAction Stop');
        lines.push('$meId = $meUser.id');
        lines.push('');
        lines.push('$i = 0');
        lines.push('foreach ($r in $rows) {');
        lines.push('    $i++');
        lines.push('    try {');
        lines.push('        $owner = Get-MgUser -UserId $r.OwnerUpn -ErrorAction Stop');
        lines.push(
            '        # M365 Unified Group: New-MgGroup -BodyParameter (Bulk-Muster, vgl. https://m365corner.com/m365-powershell/using-new-mggroup-in-graph-powershell.html )'
        );
        lines.push('        $groupBody = @{');
        lines.push('            DisplayName     = $r.DisplayName');
        lines.push('            Description     = $r.Description');
        lines.push('            MailNickname    = $r.MailNickname');
        lines.push('            MailEnabled     = $true');
        lines.push('            SecurityEnabled = $false');
        lines.push('            GroupTypes      = @("Unified")');
        lines.push('            Visibility      = "Private"');
        lines.push('        }');
        lines.push('        $group = New-MgGroup -BodyParameter $groupBody -ErrorAction Stop');
        lines.push('        Start-Sleep -Seconds 2  # Replikation vor Owner-Zuweisung');
        lines.push('        try {');
        lines.push('            New-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id -ErrorAction Stop');
        lines.push('        } catch {');
        lines.push(
            '            if ($_.Exception.Message -notmatch "already exist") { throw }'
        );
        lines.push(
            '            Write-Host ("  Hinweis (Besitzer): {0}" -f $_.Exception.Message) -ForegroundColor DarkGray'
        );
        lines.push('        }');
        lines.push('        if (-not $Ms365AdminAsOwner -and $meId -ne $owner.Id) {');
        lines.push('            try {');
        lines.push(
            '                Invoke-MgGraphRequest -Method DELETE -Uri ("https://graph.microsoft.com/v1.0/groups/{0}/owners/{1}/`$ref" -f $group.Id, $meId) -ErrorAction Stop'
        );
        lines.push(
            '                Write-Host "  Angemeldeter Administrator als Besitzer entfernt (nur Besitzer aus Schritt 2)." -ForegroundColor DarkGray'
        );
        lines.push('            } catch {');
        lines.push(
            '                Write-Host ("  Hinweis (Admin-Besitzer entfernen): {0}" -f $_.Exception.Message) -ForegroundColor DarkGray'
        );
        lines.push('            }');
        lines.push('        }');
        lines.push('        try {');
        lines.push(
            '            $memberRef = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($owner.Id)" }'
        );
        lines.push(
            '            Invoke-MgGraphRequest -Method POST -Uri (' +
                "'" +
                'https://graph.microsoft.com/v1.0/groups/{0}/members/$ref' +
                "'" +
                ' -f $group.Id) -Body ($memberRef | ConvertTo-Json -Compress) -ErrorAction Stop'
        );
        lines.push('        } catch {');
        lines.push(
            '            Write-Host ("Hinweis (Besitzer als Mitglied): {0}" -f $_.Exception.Message) -ForegroundColor DarkGray'
        );
        lines.push('        }');
        lines.push('        if ($Ms365CreateTeams) {');
        lines.push('            $teamProps = @{');
        lines.push('                memberSettings = @{ allowCreatePrivateChannels = $true; allowCreateUpdateChannels = $true }');
        lines.push(
            '                messagingSettings = @{ allowUserEditMessages = $true; allowUserDeleteMessages = $true }'
        );
        lines.push('                funSettings = @{ allowGiphy = $true; giphyContentRating = "moderate" }');
        lines.push('                guestSettings = @{ allowCreateUpdateChannels = $false }');
        lines.push('            }');
        lines.push(
            '            # Verschachtelte Hashtables zu JSON (Depth wichtig für PS 5.1 / korrekten Graph-Body)'
        );
        lines.push('            $teamJson = $teamProps | ConvertTo-Json -Depth 10 -Compress');
        lines.push('            $teamUri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/team"');
        lines.push('            for ($ti = 0; $ti -lt 8; $ti++) {');
        lines.push('                try {');
        lines.push(
            '                    Invoke-MgGraphRequest -Method PUT -Uri $teamUri -Body $teamJson -ContentType "application/json" -ErrorAction Stop'
        );
        lines.push(
            '                    Write-Host ("Teams: {0} – Team bereitgestellt." -f $r.DisplayName) -ForegroundColor Cyan'
        );
        lines.push('                    break');
        lines.push('                } catch {');
        lines.push('                    if ($ti -lt 7) {');
        lines.push(
            '                        Write-Host ("Teams: Warte auf Replikation ({0}/8) …" -f ($ti + 1)) -ForegroundColor DarkYellow'
        );
        lines.push('                        Start-Sleep -Seconds 10');
        lines.push('                    } else {');
        lines.push(
            '                        Write-Warning ("Teams: {0} – Team konnte nicht angelegt werden: {1}" -f $r.DisplayName, $_.Exception.Message)'
        );
        lines.push('                    }');
        lines.push('                }');
        lines.push('            }');
        lines.push('        }');
        lines.push('        if ($Ms365SetExchangeSmtp -and $Ms365ExchangeDomain) {');
        lines.push('            Ensure-Ms365ExchangeOnline');
        lines.push('            $wantedSmtp = "$($r.MailNickname)@$Ms365ExchangeDomain"');
        lines.push('            for ($ei = 0; $ei -lt 6; $ei++) {');
        lines.push('                try {');
        lines.push(
            '                    Set-UnifiedGroup -Identity $group.Id -PrimarySmtpAddress $wantedSmtp -ErrorAction Stop'
        );
        lines.push(
            '                    Write-Host ("Exchange: {0} – PrimarySmtpAddress = {1}" -f $r.DisplayName, $wantedSmtp) -ForegroundColor Green'
        );
        lines.push('                    break');
        lines.push('                } catch {');
        lines.push('                    if ($ei -lt 5) {');
        lines.push(
            '                        Write-Host ("Exchange: Warte auf Postfach ({0}/6) …" -f ($ei + 1)) -ForegroundColor DarkYellow'
        );
        lines.push('                        Start-Sleep -Seconds 15');
        lines.push('                    } else {');
        lines.push(
            '                        Write-Warning ("Exchange: {0} – PrimarySmtpAddress nicht gesetzt: {1}" -f $r.DisplayName, $_.Exception.Message)'
        );
        lines.push('                    }');
        lines.push('                }');
        lines.push('            }');
        lines.push('        }');
        lines.push('        Write-Host ("OK [{0}/{1}] {2} -> {3}" -f $i, $rows.Count, $r.DisplayName, $r.MailNickname) -ForegroundColor Green');
        lines.push('    }');
        lines.push('    catch {');
        lines.push('        $ex = $_.Exception');
        lines.push('        $detail = $ex.Message');
        lines.push('        if ($ex.InnerException) { $detail += " | " + $ex.InnerException.Message }');
        lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $r.DisplayName, $detail)');
        lines.push('    }');
        lines.push('    Start-Sleep -Seconds 2');
        lines.push('}');
        lines.push('');
        lines.push(
            '# SMTP: Graph legt nur mailNickname an. Mit $Ms365SetExchangeSmtp wird die primäre Adresse per Exchange gesetzt.'
        );
        lines.push('# Zieldomain (App): ' + psEscapeSingle(domainTrim || domain));
        lines.push('# Set-UnifiedGroup: https://learn.microsoft.com/powershell/module/exchange/set-unifiedgroup');
        if (setExoEffective) {
            lines.push('if ($script:Ms365ExoConnected) {');
            lines.push(
                '    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}'
            );
            lines.push('}');
            lines.push('');
        }
        if (standalone) {
            lines.push('');
            lines.push('Write-Host ""');
            lines.push('Write-Host "Fertig." -ForegroundColor Cyan');
            lines.push('Read-Host "Enter druecken zum Beenden"');
        }
        return lines.join('\r\n');
    }

    function downloadArgeStandalonePackage() {
        if (!argeRows.length) {
            showToast('Keine ARGE-Daten – zuerst ARGE-Liste, Besitzer und Einstellungen durchgehen.');
            return;
        }
        const missing = argeRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle ARGEs einen Besitzer eintragen.');
            return;
        }
        if (typeof window.ms365BuildPolyglotCmd !== 'function') {
            showToast('polyglot-cmd.js fehlt – Seite neu laden.');
            return;
        }
        if (getArgeExchangeSmtp() && !getDomain().trim()) {
            showToast('Für die Exchange-Option bitte oben die E-Mail-Domain der Schule eintragen.');
            return;
        }
        const ps1 = buildStandaloneArgePs1(
            true,
            getArgeCreateTeams(),
            getArgeExchangeSmtp(),
            getArgeAdminAsOwner()
        );
        const cmd = window.ms365BuildPolyglotCmd({
            title: 'ARGE-Gruppen-Anlage',
            echoLine: 'Starte ARGE-Gruppen-Anlage Microsoft Graph ...',
            psBody: ps1
        });
        downloadBlob('ARGE-Gruppen-Anlage.cmd', cmd);
        showToast('ARGE-Gruppen-Anlage.cmd heruntergeladen – Doppelklick zum Start.');
    }

    window.downloadArgeStandalonePackage = downloadArgeStandalonePackage;

    // UI Wiring — Vorschau zuerst, damit Eingabe auch bei späteren Fehlern funktioniert
    const argeLinesEl = document.getElementById('argeLines');
    if (argeLinesEl) {
        argeLinesEl.addEventListener('input', scheduleArgePreviewRefresh);
        argeLinesEl.addEventListener('paste', () => setTimeout(scheduleArgePreviewRefresh, 0));
    }
    ['schoolEmailDomain', 'argeDefaultPrefix'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', scheduleArgePreviewRefresh);
            el.addEventListener('input', refreshArgeScriptIfStep4);
        }
    });
    const argeUpperEl = document.getElementById('argeUpperNick');
    if (argeUpperEl) argeUpperEl.addEventListener('change', scheduleArgePreviewRefresh);
    const argeTeamsEl = document.getElementById('argeCreateTeams');
    if (argeTeamsEl) argeTeamsEl.addEventListener('change', refreshArgeScriptIfStep4);
    const argeExoEl = document.getElementById('argeExchangeSmtp');
    if (argeExoEl) argeExoEl.addEventListener('change', refreshArgeScriptIfStep4);
    const argeAdminAsOwnerEl = document.getElementById('argeAdminAsOwner');
    if (argeAdminAsOwnerEl) argeAdminAsOwnerEl.addEventListener('change', refreshArgeScriptIfStep4);

    document.getElementById('argeBack1').addEventListener('click', () => goToArgeStep(1));
    document.getElementById('argeGoTo3').addEventListener('click', () => goToArgeStep(3));
    document.getElementById('argeBack2').addEventListener('click', () => goToArgeStep(2));
    document.getElementById('argeBack3').addEventListener('click', () => goToArgeStep(3));

    document.getElementById('argeParseAndGo3').addEventListener('click', () => {
        const errEl = document.getElementById('argeParseError');
        errEl.style.display = 'none';
        const domain = getDomain();
        if (!domain) {
            errEl.textContent = 'Bitte oben die E-Mail-Domain angeben.';
            errEl.style.display = 'block';
            return;
        }

        const { parsed, errors } = parseArgeInput();

        if (errors.length) {
            errEl.textContent = errors.join('\n');
            errEl.style.display = 'block';
            return;
        }
        if (!parsed.length) {
            errEl.textContent = 'Bitte mindestens eine ARGE-Zeile eintragen.';
            errEl.style.display = 'block';
            return;
        }

        const rows = parsed.map(r => ({ ...r }));
        resolveDuplicateNicks(rows);
        argeRows = rows.map(r => ({
            displayName: r.displayName,
            mailNick: r.mailNick,
            owner: '',
            description: r.description
        }));

        rebuildArgeOwnerTableFromRows();

        goToArgeStep(2);
    });

    document.getElementById('argeGoTo4').addEventListener('click', () => {
        const sync = syncArgeRowsFromInputPreservingOwners();
        if (!sync.ok) {
            const errEl = document.getElementById('argeParseError');
            if (errEl) {
                errEl.textContent = sync.errors.join('\n');
                errEl.style.display = 'block';
            }
            showToast(sync.errors[0] || 'ARGE-Liste konnte nicht verarbeitet werden.');
            goToArgeStep(1);
            scheduleArgePreviewRefresh();
            return;
        }
        const missing = argeRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle ARGEs einen Besitzer (UPN) eintragen (Schritt 2).');
            goToArgeStep(2);
            return;
        }
        if (getArgeExchangeSmtp() && !getDomain().trim()) {
            showToast('Für die Exchange-Option bitte oben die E-Mail-Domain der Schule eintragen.');
            return;
        }
        document.getElementById('argeParseError').style.display = 'none';
        document.getElementById('argePowerShellScript').textContent = buildStandaloneArgePs1(
            false,
            getArgeCreateTeams(),
            getArgeExchangeSmtp(),
            getArgeAdminAsOwner()
        );
        goToArgeStep(4);
    });

    document.getElementById('argeCopyScript').addEventListener('click', () => {
        const t = document.getElementById('argePowerShellScript').textContent;
        navigator.clipboard.writeText(t).then(() => showToast('Script kopiert.'));
    });

    // step header keyboard support
    document.querySelectorAll('.arge-steps .step').forEach(el => {
        el.setAttribute('tabindex', '0');
        el.addEventListener('click', () => {
            const s = argeStepNum(el);
            if (s <= argeCurrentStep || el.classList.contains('completed')) {
                goToArgeStep(s);
            }
        });
        el.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                el.click();
            }
        });
    });

    applyInitialModeFromUrl();
    if (panelA && panelA.style.display !== 'none') {
        scheduleArgePreviewRefresh();
    }
})();

