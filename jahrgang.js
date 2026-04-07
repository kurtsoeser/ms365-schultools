(function () {
    'use strict';

    let jgCurrentStep = 1;
    /** @type {{ klasse: string, jahr: string, suffix: string, mailNick: string, owner: string }[]} */
    let jgRows = [];

    const panelW = document.getElementById('panelWebuntis');
    const panelJ = document.getElementById('panelJahrgang');
    const btnModeW = document.getElementById('modeWebuntis');
    const btnModeJ = document.getElementById('modeJahrgang');

    function showToast(msg) {
        const el = document.getElementById('toast');
        if (!el) return;
        el.textContent = msg;
        el.classList.add('show');
        clearTimeout(showToast._t);
        showToast._t = setTimeout(() => el.classList.remove('show'), 3500);
    }

    function setMode(webuntis) {
        if (webuntis) {
            panelW.style.display = '';
            panelJ.style.display = 'none';
            btnModeW.classList.add('btn-success');
            btnModeJ.classList.remove('btn-success');
        } else {
            panelW.style.display = 'none';
            panelJ.style.display = '';
            btnModeJ.classList.add('btn-success');
            btnModeW.classList.remove('btn-success');
        }
    }

    btnModeW.addEventListener('click', () => setMode(true));
    btnModeJ.addEventListener('click', () => setMode(false));

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
        let d = (document.getElementById('jgDomain').value || '').trim().replace(/^@/, '');
        return d;
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

    function updatePrefixExample() {
        const dom = getDomain();
        const pre = getPrefix();
        const ex = buildMailNickname(pre, '2030', 'AK');
        const el = document.getElementById('jgPrefixExample');
        if (el) el.textContent = ex + '@' + (dom || 'ihre-schule.at');
    }

    ['jgDomain', 'jgPrefix', 'jgSuffixUpper'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', updatePrefixExample);
        if (el) el.addEventListener('change', updatePrefixExample);
    });

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
            errEl.textContent = 'Bitte in Schritt 1 die E-Mail-Domain angeben.';
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
            inp.dataset.index = String(index);
            inp.addEventListener('input', () => {
                jgRows[index].owner = inp.value.trim();
            });
            td4.appendChild(inp);
            tr.append(td1, td2, td3, td4);
            tbody.appendChild(tr);
        });

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
        return buildStandaloneJahrgangPs1(false);
    }

    function downloadBlob(filename, text, mime) {
        const blob = new Blob([text], { type: mime || 'text/plain;charset=utf-8' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        a.click();
        URL.revokeObjectURL(a.href);
    }

    function buildStandaloneJahrgangPs1(standalone) {
        const domain = getDomain();
        const stamp = new Date().toISOString();
        const lines = [];
        if (standalone) {
            lines.push('#Requires -Version 5.1');
            lines.push('# Jahrgangsgruppen (Microsoft 365 Unified Groups, kein Kursteam)');
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
            lines.push('');
            lines.push('if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {');
            lines.push('    Write-Host "Installiere Microsoft.Graph (einmalig)..." -ForegroundColor Yellow');
            lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force');
            lines.push('}');
            lines.push('Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('');
            lines.push('Write-Host "Anmeldung bei Microsoft Graph (interaktiv, MFA moeglich)..." -ForegroundColor Yellow');
            lines.push('Write-Host "Es oeffnet sich ein Browser- oder Anmeldedialog." -ForegroundColor Gray');
            lines.push('Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All"');
            lines.push('');
        } else {
            lines.push('# Microsoft Graph: Jahrgangsgruppen als Microsoft 365-Gruppen (Unified Group, kein Kursteam)');
            lines.push('# Voraussetzung: Install-Module Microsoft.Graph');
            lines.push('# https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup');
            lines.push('');
            lines.push("Install-Module Microsoft.Graph -Scope CurrentUser -ErrorAction SilentlyContinue");
            lines.push("Import-Module Microsoft.Graph -ErrorAction Stop");
            lines.push('Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All"');
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
        lines.push('        $group = New-MgGroup `');
        lines.push('            -DisplayName $r.Klasse `');
        lines.push('            -Description $r.Description `');
        lines.push('            -MailNickname $r.MailNickname `');
        lines.push('            -MailEnabled:$true `');
        lines.push('            -SecurityEnabled:$false `');
        lines.push('            -GroupTypes @("Unified") `');
        lines.push('            -Visibility "Private" `');
        lines.push('            -ErrorAction Stop');
        lines.push('        New-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id');
        lines.push('        Write-Host ("OK [{0}/{1}] {2} -> {3}" -f $i, $rows.Count, $r.Klasse, $r.MailNickname)');
        lines.push('    }');
        lines.push('    catch {');
        lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $r.Klasse, $_.Exception.Message)');
        lines.push('    }');
        lines.push('    Start-Sleep -Seconds 2');
        lines.push('}');
        lines.push('');
        lines.push('# Gruppen-E-Mail: <MailNickname>@' + psEscapeSingle(domain));
        if (standalone) {
            lines.push('');
            lines.push('Write-Host ""');
            lines.push('Write-Host "Fertig." -ForegroundColor Cyan');
            lines.push('Read-Host "Enter druecken zum Beenden"');
        }
        return lines.join('\r\n');
    }

    function buildJahrgangCmdContent() {
        return [
            '@echo off',
            'chcp 65001 >nul',
            'title Jahrgangsgruppen-Anlage',
            'cd /d "%~dp0"',
            'echo.',
            'echo Starte Jahrgangsgruppen-Anlage (Microsoft Graph)...',
            'echo.',
            'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Jahrgangsgruppen-Anlage.ps1"',
            'set ERR=%ERRORLEVEL%',
            'if not "%ERR%"=="0" (',
            '  echo.',
            '  echo Fehlercode: %ERR%',
            ')',
            'echo.',
            'pause',
            ''
        ].join('\r\n');
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
        const ps1 = buildStandaloneJahrgangPs1(true);
        downloadBlob('Jahrgangsgruppen-Anlage.ps1', ps1);
        setTimeout(() => {
            downloadBlob('Jahrgangsgruppen-Anlage.cmd', buildJahrgangCmdContent());
            showToast('Dateien: Jahrgangsgruppen-Anlage.ps1 + .cmd heruntergeladen.');
        }, 500);
    }

    window.downloadJahrgangStandalonePackage = downloadJahrgangStandalonePackage;

    updatePrefixExample();

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
