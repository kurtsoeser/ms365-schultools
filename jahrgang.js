(function () {
    'use strict';

    let jgCurrentStep = 1;
    /** @type {{ klasse: string, jahr: string, suffix: string, mailNick: string, owner: string }[]} */
    let jgRows = [];

    const panelW = document.getElementById('panelWebuntis');
    const panelJ = document.getElementById('panelJahrgang');
    const btnModeW = document.getElementById('modeWebuntis');
    const btnModeJ = document.getElementById('modeJahrgang');
    const panelA = document.getElementById('panelArge');
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
        panelW.style.display = w ? '' : 'none';
        panelJ.style.display = j ? '' : 'none';
        if (panelA) panelA.style.display = a ? '' : 'none';
        btnModeW.classList.toggle('btn-success', w);
        btnModeJ.classList.toggle('btn-success', j);
        if (btnModeA) btnModeA.classList.toggle('btn-success', a);
    }

    btnModeW.addEventListener('click', () => setMode('webuntis'));
    btnModeJ.addEventListener('click', () => setMode('jahrgang'));
    if (btnModeA) btnModeA.addEventListener('click', () => setMode('arge'));

    function applyInitialModeFromUrl() {
        try {
            const mode = new URLSearchParams(window.location.search).get('mode');
            if (!mode) return;
            if (mode.toLowerCase() === 'jahrgang') setMode('jahrgang');
            else if (mode.toLowerCase() === 'arge') setMode('arge');
            else if (mode.toLowerCase() === 'kursteams' || mode.toLowerCase() === 'kursteam' || mode.toLowerCase() === 'webuntis') setMode('webuntis');
        } catch {
            // ignore
        }
    }

    function goToJgStep(step) {
        jgCurrentStep = step;
        document.querySelectorAll('.jg-step-content').forEach(el => {
            el.classList.toggle('active', parseFloat(el.dataset.jgStep) === step);
        });
        document.querySelectorAll('.jg-steps .step').forEach(el => {
            const s = parseFloat(el.dataset.jgStep);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
    }

    function parseClassLine(line) {
        const trimmed = line.trim();
        if (!trimmed || trimmed.startsWith('#')) return { skip: true };
        const parts = trimmed.split(/[;,\t]/).map(s => s.trim()).filter(Boolean);
        if (parts.length < 2) {
            return { error: 'Zeile benötigt Klasse und Jahr: ' + trimmed };
        }
        const klasse = parts[0];
        const jahr = parts[1];
        if (!/^\d{4}$/.test(jahr)) {
            return { error: 'Abschlussjahr muss vierstellig sein: ' + trimmed };
        }
        const m = klasse.match(/^(\d+)([A-Za-z]+)$/);
        if (!m) {
            return { error: 'Klasse erwartet z.B. 1AK (Ziffern + Buchstaben): ' + trimmed };
        }
        return { klasse, jahr, suffix: m[2] };
    }

    function getDomain() {
        if (typeof window.ms365GetSchoolDomainNoAt === 'function') {
            return window.ms365GetSchoolDomainNoAt();
        }
        return '';
    }

    function getPrefix() {
        return (document.getElementById('jgPrefix').value || 'jg').trim().toLowerCase().replace(/[^a-z0-9]/g, '') || 'jg';
    }

    function suffixForNick(suffix) {
        const upper = document.getElementById('jgSuffixUpper').checked;
        const s = suffix.replace(/[^A-Za-z0-9]/g, '');
        return upper ? s.toUpperCase() : s.toLowerCase();
    }

    function buildMailNickname(prefix, year, suffix) {
        const suf = suffixForNick(suffix);
        return (prefix + year + '-' + suf).replace(/[^a-zA-Z0-9-]/g, '');
    }

    function resolveDuplicateNicks(rows) {
        const seen = new Map();
        rows.forEach(row => {
            let base = row.mailNick;
            let candidate = base;
            let n = 2;
            while (seen.has(candidate)) {
                candidate = base + n;
                n++;
            }
            row.mailNick = candidate;
            seen.set(candidate, true);
        });
    }

    const JG_STORAGE_KEY = 'ms365-jahrgang-state-v1';

    function getJgCreateTeams() {
        const el = document.getElementById('jgCreateTeams');
        return el ? !!el.checked : true;
    }

    function getJgExchangeSmtp() {
        const el = document.getElementById('jgExchangeSmtp');
        return el ? !!el.checked : true;
    }

    function refreshJgScriptIfStep4() {
        if (jgCurrentStep !== 4 || !jgRows.length) return;
        const missing = jgRows.filter(r => !r.owner);
        if (missing.length) return;
        const pre = document.getElementById('jgPowerShellScript');
        if (pre) pre.textContent = buildStandaloneJahrgangPs1(false, getJgCreateTeams(), getJgExchangeSmtp());
    }

    function rebuildJgOwnerTableFromRows() {
        const domain = getDomain();
        const tbody = document.getElementById('jgOwnerBody');
        tbody.replaceChildren();
        jgRows.forEach((row, index) => {
            const tr = document.createElement('tr');
            const td1 = document.createElement('td');
            td1.textContent = row.klasse;
            const td2 = document.createElement('td');
            td2.textContent = row.mailNick + '@' + domain;
            const td3 = document.createElement('td');
            td3.textContent = row.mailNick;
            const td4 = document.createElement('td');
            const inp = document.createElement('input');
            inp.type = 'email';
            inp.placeholder = 'lehrer@' + domain;
            inp.style.width = '100%';
            inp.style.padding = '8px';
            inp.style.boxSizing = 'border-box';
            inp.value = row.owner || '';
            inp.addEventListener('input', () => {
                jgRows[index].owner = inp.value.trim();
            });
            td4.appendChild(inp);
            tr.append(td1, td2, td3, td4);
            tbody.appendChild(tr);
        });
    }

    function saveJahrgangState() {
        try {
            const state = {
                jgCurrentStep,
                jgRows,
                jgPrefix: document.getElementById('jgPrefix').value,
                jgSuffixUpper: document.getElementById('jgSuffixUpper').checked,
                jgCreateTeams: getJgCreateTeams(),
                jgExchangeSmtp: getJgExchangeSmtp(),
                jgClassLines: document.getElementById('jgClassLines').value,
                jgPowerShellScript: document.getElementById('jgPowerShellScript').textContent
            };
            localStorage.setItem(JG_STORAGE_KEY, JSON.stringify(state));
            showToast('Jahrgangsgruppen: Zwischenstand gespeichert.');
        } catch (e) {
            showToast('Speichern fehlgeschlagen: ' + e.message);
        }
    }

    function loadJahrgangState() {
        try {
            const raw = localStorage.getItem(JG_STORAGE_KEY);
            if (!raw) {
                showToast('Kein gespeicherter Stand für Jahrgangsgruppen.');
                return;
            }
            const state = JSON.parse(raw);
            jgCurrentStep = typeof state.jgCurrentStep === 'number' ? state.jgCurrentStep : 1;
            jgRows = Array.isArray(state.jgRows) ? state.jgRows : [];
            if (
                typeof window.ms365SetSchoolDomainNoAt === 'function' &&
                state.jgDomain !== undefined &&
                String(state.jgDomain).trim() !== ''
            ) {
                window.ms365SetSchoolDomainNoAt(state.jgDomain);
            }
            document.getElementById('jgPrefix').value = state.jgPrefix !== undefined ? state.jgPrefix : 'jg';
            document.getElementById('jgSuffixUpper').checked = state.jgSuffixUpper !== false;
            const jgTeamsEl = document.getElementById('jgCreateTeams');
            if (jgTeamsEl) {
                jgTeamsEl.checked = state.jgCreateTeams !== undefined ? !!state.jgCreateTeams : true;
            }
            const jgExoEl = document.getElementById('jgExchangeSmtp');
            if (jgExoEl) {
                jgExoEl.checked = state.jgExchangeSmtp !== undefined ? !!state.jgExchangeSmtp : true;
            }
            document.getElementById('jgClassLines').value = state.jgClassLines || '';
            document.getElementById('jgParseError').style.display = 'none';
            const pre = document.getElementById('jgPowerShellScript');
            if (pre && state.jgPowerShellScript !== undefined) {
                pre.textContent = state.jgPowerShellScript;
            }
            updatePrefixExample();
            if (jgRows.length) {
                rebuildJgOwnerTableFromRows();
            } else {
                document.getElementById('jgOwnerBody').replaceChildren();
            }
            const step = Math.min(Math.max(1, jgCurrentStep), 4);
            goToJgStep(step);
            showToast('Jahrgangsgruppen: Stand geladen.');
        } catch (e) {
            showToast('Laden fehlgeschlagen: ' + e.message);
        }
    }

    function clearJahrgangState() {
        if (!confirm('Gespeicherten Zwischenstand für Jahrgangsgruppen wirklich löschen?')) {
            return;
        }
        try {
            localStorage.removeItem(JG_STORAGE_KEY);
            jgCurrentStep = 1;
            jgRows = [];
            document.getElementById('jgPrefix').value = 'jg';
            document.getElementById('jgSuffixUpper').checked = true;
            const jgTeamsClear = document.getElementById('jgCreateTeams');
            if (jgTeamsClear) jgTeamsClear.checked = true;
            const jgExoClear = document.getElementById('jgExchangeSmtp');
            if (jgExoClear) jgExoClear.checked = true;
            document.getElementById('jgClassLines').value = '';
            document.getElementById('jgParseError').style.display = 'none';
            document.getElementById('jgOwnerBody').replaceChildren();
            document.getElementById('jgPowerShellScript').textContent = '';
            updatePrefixExample();
            goToJgStep(1);
            showToast('Jahrgangsgruppen: Speicher geleert.');
        } catch (e) {
            showToast('Fehler: ' + e.message);
        }
    }

    window.ms365SaveJahrgang = saveJahrgangState;
    window.ms365LoadJahrgang = loadJahrgangState;
    window.ms365ClearJahrgang = clearJahrgangState;

    function updatePrefixExample() {
        const dom = getDomain();
        const pre = getPrefix();
        const ex = buildMailNickname(pre, '2030', 'AK');
        const el = document.getElementById('jgPrefixExample');
        if (el) {
            const fallback =
                typeof window.ms365DefaultSchoolDomainNoAt === 'function'
                    ? window.ms365DefaultSchoolDomainNoAt()
                    : 'ms365.schule';
            el.textContent = ex + '@' + (dom || fallback);
        }
    }

    ['schoolEmailDomain', 'jgPrefix', 'jgSuffixUpper'].forEach(id => {
        const el = document.getElementById(id);
        if (el) {
            el.addEventListener('input', updatePrefixExample);
            el.addEventListener('change', updatePrefixExample);
            el.addEventListener('input', refreshJgScriptIfStep4);
            el.addEventListener('change', refreshJgScriptIfStep4);
        }
    });
    const jgTeamsEl = document.getElementById('jgCreateTeams');
    if (jgTeamsEl) jgTeamsEl.addEventListener('change', refreshJgScriptIfStep4);
    const jgExoEl = document.getElementById('jgExchangeSmtp');
    if (jgExoEl) jgExoEl.addEventListener('change', refreshJgScriptIfStep4);

    updatePrefixExample();

    document.getElementById('jgGoTo2').addEventListener('click', () => goToJgStep(2));

    document.getElementById('jgBack1').addEventListener('click', () => goToJgStep(1));

    document.getElementById('jgParseAndGo3').addEventListener('click', () => {
        const errEl = document.getElementById('jgParseError');
        errEl.style.display = 'none';
        const lines = document.getElementById('jgClassLines').value.split(/\r?\n/);
        const parsed = [];
        const errors = [];
        lines.forEach((line, idx) => {
            const r = parseClassLine(line);
            if (r.skip) return;
            if (r.error) {
                errors.push('Zeile ' + (idx + 1) + ': ' + r.error);
                return;
            }
            parsed.push(r);
        });
        if (errors.length) {
            errEl.textContent = errors.join('\n');
            errEl.style.display = 'block';
            return;
        }
        if (!parsed.length) {
            errEl.textContent = 'Bitte mindestens eine Klassenzeile eintragen.';
            errEl.style.display = 'block';
            return;
        }
        const seenKlasse = new Set();
        const deduped = [];
        for (const p of parsed) {
            if (seenKlasse.has(p.klasse)) {
                continue;
            }
            seenKlasse.add(p.klasse);
            deduped.push(p);
        }
        const parsedFinal = deduped;

        const domain = getDomain();
        if (!domain) {
            errEl.textContent = 'Bitte oben die E-Mail-Domain der Schule eintragen.';
            errEl.style.display = 'block';
            return;
        }

        const prefix = getPrefix();
        jgRows = parsedFinal.map(p => ({
            klasse: p.klasse,
            jahr: p.jahr,
            suffix: p.suffix,
            mailNick: buildMailNickname(prefix, p.jahr, p.suffix),
            owner: ''
        }));
        resolveDuplicateNicks(jgRows);

        rebuildJgOwnerTableFromRows();

        goToJgStep(3);
    });

    document.getElementById('jgBack2').addEventListener('click', () => goToJgStep(2));

    document.getElementById('jgGoTo4').addEventListener('click', () => {
        if (!jgRows.length) {
            showToast('Bitte zuerst die Klassenliste parsen (Schritt 2).');
            return;
        }
        const missing = jgRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle Klassen eine Besitzer-E-Mail (UPN) eintragen.');
            return;
        }
        if (getJgExchangeSmtp() && !getDomain().trim()) {
            showToast('Für die Exchange-Option bitte oben die E-Mail-Domain der Schule eintragen.');
            return;
        }
        document.getElementById('jgPowerShellScript').textContent = buildPowerShellScript();
        goToJgStep(4);
    });

    document.getElementById('jgBack3').addEventListener('click', () => goToJgStep(3));

    document.getElementById('jgCopyScript').addEventListener('click', () => {
        const t = document.getElementById('jgPowerShellScript').textContent;
        navigator.clipboard.writeText(t).then(() => showToast('Script kopiert.'));
    });

    function psEscapeSingle(s) {
        return String(s).replace(/'/g, "''");
    }

    function buildPowerShellScript() {
        return buildStandaloneJahrgangPs1(false, getJgCreateTeams(), getJgExchangeSmtp());
    }

    function downloadBlob(filename, text, mime) {
        const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function buildStandaloneJahrgangPs1(standalone, createTeams, setExchangeSmtp) {
        if (createTeams === undefined) createTeams = true;
        if (setExchangeSmtp === undefined) setExchangeSmtp = true;
        const domain = getDomain();
        const domainTrim = (domain || '').trim();
        const setExoEffective = setExchangeSmtp && domainTrim.length > 0;
        const stamp = new Date().toISOString();
        const lines = [];
        const scopesLine = '$scopes = @("Group.ReadWrite.All","User.Read.All")';

        if (standalone) {
            lines.push('#Requires -Version 5.1');
            lines.push('# Jahrgangsgruppen (M365 Unified); optional Teams ($Ms365CreateTeams); optional Exchange-SMTP ($Ms365SetExchangeSmtp)');
            lines.push('# Erzeugt in der Browser-App am ' + stamp);
            lines.push('# Daten sind unten eingebettet.');
            lines.push('');
            lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
            lines.push('$ErrorActionPreference = "Continue"');
            lines.push('');
            lines.push('Write-Host ""');
            lines.push('Write-Host "========================================"  -ForegroundColor Cyan');
            lines.push('Write-Host "  Jahrgangsgruppen (Microsoft Graph)"   -ForegroundColor Cyan');
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
            lines.push('# Microsoft Graph: Jahrgangsgruppen als Microsoft 365-Gruppen (Unified Group)');
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
            lines.push('Ensure-Ms365ExchangeOnline');
            lines.push('');
        }
        lines.push('$rows = @(');
        jgRows.forEach((r, i) => {
            const last = i === jgRows.length - 1;
            lines.push(
                "    [PSCustomObject]@{ Klasse = '" +
                    psEscapeSingle(r.klasse) +
                    "'; MailNickname = '" +
                    psEscapeSingle(r.mailNick) +
                    "'; OwnerUpn = '" +
                    psEscapeSingle(r.owner) +
                    "'; Description = 'Jahrgangsgruppe " +
                    psEscapeSingle(r.klasse) +
                    " (Abschluss " +
                    psEscapeSingle(r.jahr) +
                    ")' }" +
                    (last ? '' : ',')
            );
        });
        lines.push(')');
        lines.push('');
        lines.push('$i = 0');
        lines.push('foreach ($r in $rows) {');
        lines.push('    $i++');
        lines.push('    try {');
        lines.push("        $owner = Get-MgUser -UserId $r.OwnerUpn -ErrorAction Stop");
        lines.push(
            '        # M365 Unified Group: New-MgGroup -BodyParameter (Bulk-Muster, vgl. https://m365corner.com/m365-powershell/using-new-mggroup-in-graph-powershell.html )'
        );
        lines.push('        $groupBody = @{');
        lines.push('            DisplayName     = $r.Klasse');
        lines.push('            Description     = $r.Description');
        lines.push('            MailNickname    = $r.MailNickname');
        lines.push('            MailEnabled     = $true');
        lines.push('            SecurityEnabled = $false');
        lines.push('            GroupTypes      = @("Unified")');
        lines.push('            Visibility      = "Private"');
        lines.push('        }');
        lines.push('        $group = New-MgGroup -BodyParameter $groupBody -ErrorAction Stop');
        lines.push('        Start-Sleep -Seconds 2  # Replikation vor Owner-Zuweisung');
        lines.push('        New-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id');
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
        lines.push('            $teamUri = "https://graph.microsoft.com/v1.0/groups/$($group.Id)/team"');
        lines.push('            for ($ti = 0; $ti -lt 8; $ti++) {');
        lines.push('                try {');
        lines.push(
            '                    Invoke-MgGraphRequest -Method PUT -Uri $teamUri -Body $teamProps -ErrorAction Stop'
        );
        lines.push(
            '                    Write-Host ("Teams: {0} – Team bereitgestellt." -f $r.Klasse) -ForegroundColor Cyan'
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
            '                        Write-Warning ("Teams: {0} – Team konnte nicht angelegt werden: {1}" -f $r.Klasse, $_.Exception.Message)'
        );
        lines.push('                    }');
        lines.push('                }');
        lines.push('            }');
        lines.push('        }');
        lines.push('        if ($Ms365SetExchangeSmtp -and $Ms365ExchangeDomain) {');
        lines.push('            $wantedSmtp = "$($r.MailNickname)@$Ms365ExchangeDomain"');
        lines.push('            for ($ei = 0; $ei -lt 6; $ei++) {');
        lines.push('                try {');
        lines.push(
            '                    Set-UnifiedGroup -Identity $group.Id -PrimarySmtpAddress $wantedSmtp -ErrorAction Stop'
        );
        lines.push(
            '                    Write-Host ("Exchange: {0} – PrimarySmtpAddress = {1}" -f $r.Klasse, $wantedSmtp) -ForegroundColor Green'
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
            '                        Write-Warning ("Exchange: {0} – PrimarySmtpAddress nicht gesetzt: {1}" -f $r.Klasse, $_.Exception.Message)'
        );
        lines.push('                    }');
        lines.push('                }');
        lines.push('            }');
        lines.push('        }');
        lines.push(
            '        Write-Host ("OK [{0}/{1}] {2} -> {3}" -f $i, $rows.Count, $r.Klasse, $r.MailNickname) -ForegroundColor Green'
        );
        lines.push('    }');
        lines.push('    catch {');
        lines.push('        $ex = $_.Exception');
        lines.push('        $detail = $ex.Message');
        lines.push('        if ($ex.InnerException) { $detail += " | " + $ex.InnerException.Message }');
        lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $r.Klasse, $detail)');
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

    function downloadJahrgangStandalonePackage() {
        if (!jgRows.length) {
            showToast('Keine Klassen – zuerst Schritt 2 und 3 abschließen.');
            return;
        }
        const missing = jgRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle Klassen einen Besitzer eintragen.');
            return;
        }
        if (typeof window.ms365BuildPolyglotCmd !== 'function') {
            showToast('polyglot-cmd.js fehlt – Seite neu laden.');
            return;
        }
        if (getJgExchangeSmtp() && !getDomain().trim()) {
            showToast('Für die Exchange-Option bitte oben die E-Mail-Domain der Schule eintragen.');
            return;
        }
        const ps1 = buildStandaloneJahrgangPs1(true, getJgCreateTeams(), getJgExchangeSmtp());
        const cmd = window.ms365BuildPolyglotCmd({
            title: 'Jahrgangsgruppen-Anlage',
            echoLine: 'Starte Jahrgangsgruppen-Anlage Microsoft Graph ...',
            psBody: ps1
        });
        downloadBlob('Jahrgangsgruppen-Anlage.cmd', cmd);
        showToast('Jahrgangsgruppen-Anlage.cmd heruntergeladen – Doppelklick zum Start.');
    }

    window.downloadJahrgangStandalonePackage = downloadJahrgangStandalonePackage;

    updatePrefixExample();
    applyInitialModeFromUrl();

    document.querySelectorAll('.jg-steps .step').forEach(el => {
        el.setAttribute('tabindex', '0');
        el.addEventListener('click', () => {
            const s = parseFloat(el.dataset.jgStep);
            if (s <= jgCurrentStep || el.classList.contains('completed')) {
                goToJgStep(s);
            }
        });
        el.addEventListener('keydown', (e) => {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                el.click();
            }
        });
    });
})();
