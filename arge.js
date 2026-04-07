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

    function goToArgeStep(step) {
        argeCurrentStep = step;
        document.querySelectorAll('.arge-step-content').forEach(el => {
            el.classList.toggle('active', parseFloat(el.dataset.argeStep) === step);
        });
        document.querySelectorAll('.arge-steps .step').forEach(el => {
            const s = parseFloat(el.dataset.argeStep);
            el.classList.toggle('active', s === step);
            el.classList.toggle('completed', s < step);
        });
        if (step === 2) {
            scheduleArgePreviewRefresh();
        }
    }

    function getDomain() {
        return (document.getElementById('argeDomain').value || '').trim().replace(/^@/, '');
    }

    function getPrefix() {
        const raw = (document.getElementById('argeDefaultPrefix').value || '').trim();
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
        const upper = document.getElementById('argeUpperNick').checked;
        return upper ? s.toUpperCase() : s.toLowerCase();
    }

    /** Mail-Nickname nur aus dem Fach (Präfix aus Schritt 1), nicht aus „ARGE …“ doppelt */
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
        const lines = document.getElementById('argeLines').value.split(/\r?\n/);
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

    function buildStandaloneArgePs1(standalone) {
        const domain = getDomain();
        const stamp = new Date().toISOString();
        const lines = [];

        if (standalone) {
            lines.push('#Requires -Version 5.1');
            lines.push('# ARGE-Gruppen (Microsoft 365 Unified Groups, kein Kursteam)');
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
            lines.push('');
            lines.push('if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {');
            lines.push('    Write-Host "Installiere Microsoft.Graph (einmalig)..." -ForegroundColor Yellow');
            lines.push('    Install-Module Microsoft.Graph -Scope CurrentUser -Force');
            lines.push('}');
            lines.push('Import-Module Microsoft.Graph -ErrorAction Stop');
            lines.push('');
            lines.push('Write-Host "Anmeldung bei Microsoft Graph (interaktiv, MFA moeglich)..." -ForegroundColor Yellow');
            lines.push('Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All"');
            lines.push('');
        } else {
            lines.push('# Microsoft Graph: ARGE-Gruppen als Microsoft 365-Gruppen (Unified Group, kein Kursteam)');
            lines.push('# Voraussetzung: Install-Module Microsoft.Graph');
            lines.push('# https://learn.microsoft.com/powershell/module/microsoft.graph.groups/new-mggroup');
            lines.push('');
            lines.push("Install-Module Microsoft.Graph -Scope CurrentUser -ErrorAction SilentlyContinue");
            lines.push("Import-Module Microsoft.Graph -ErrorAction Stop");
            lines.push('Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All"');
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
        lines.push('');
        lines.push('$i = 0');
        lines.push('foreach ($r in $rows) {');
        lines.push('    $i++');
        lines.push('    try {');
        lines.push('        $owner = Get-MgUser -UserId $r.OwnerUpn -ErrorAction Stop');
        lines.push('        $group = New-MgGroup `');
        lines.push('            -DisplayName $r.DisplayName `');
        lines.push('            -Description $r.Description `');
        lines.push('            -MailNickname $r.MailNickname `');
        lines.push('            -MailEnabled:$true `');
        lines.push('            -SecurityEnabled:$false `');
        lines.push('            -GroupTypes @("Unified") `');
        lines.push('            -Visibility "Private" `');
        lines.push('            -ErrorAction Stop');
        lines.push('        New-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id');
        lines.push('        Write-Host ("OK [{0}/{1}] {2} -> {3}" -f $i, $rows.Count, $r.DisplayName, $r.MailNickname) -ForegroundColor Green');
        lines.push('    }');
        lines.push('    catch {');
        lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $r.DisplayName, $_.Exception.Message)');
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

    function buildArgeCmdContent() {
        return [
            '@echo off',
            'chcp 65001 >nul',
            'title ARGE-Gruppen-Anlage',
            'cd /d \"%~dp0\"',
            'echo.',
            'echo Starte ARGE-Gruppen-Anlage (Microsoft Graph)...',
            'echo.',
            'powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"%~dp0ARGE-Gruppen-Anlage.ps1\"',
            'set ERR=%ERRORLEVEL%',
            'if not \"%ERR%\"==\"0\" (',
            '  echo.',
            '  echo Fehlercode: %ERR%',
            ')',
            'echo.',
            'pause',
            ''
        ].join('\r\n');
    }

    function downloadArgeStandalonePackage() {
        if (!argeRows.length) {
            showToast('Keine ARGE-Daten – zuerst Schritt 2 und 3 abschließen.');
            return;
        }
        const missing = argeRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle ARGEs einen Besitzer eintragen.');
            return;
        }
        downloadBlob('ARGE-Gruppen-Anlage.ps1', buildStandaloneArgePs1(true));
        setTimeout(() => downloadBlob('ARGE-Gruppen-Anlage.cmd', buildArgeCmdContent()), 500);
        showToast('Dateien: ARGE-Gruppen-Anlage.ps1 + .cmd heruntergeladen.');
    }

    window.downloadArgeStandalonePackage = downloadArgeStandalonePackage;

    // UI Wiring
    document.getElementById('argeGoTo2').addEventListener('click', () => goToArgeStep(2));
    document.getElementById('argeBack1').addEventListener('click', () => goToArgeStep(1));
    document.getElementById('argeBack2').addEventListener('click', () => goToArgeStep(2));
    document.getElementById('argeBack3').addEventListener('click', () => goToArgeStep(3));

    document.getElementById('argeParseAndGo3').addEventListener('click', () => {
        const errEl = document.getElementById('argeParseError');
        errEl.style.display = 'none';
        const domain = getDomain();
        if (!domain) {
            errEl.textContent = 'Bitte Domain in Schritt 1 angeben.';
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
            inp.addEventListener('input', () => {
                argeRows[index].owner = inp.value.trim();
            });
            td4.appendChild(inp);
            tr.append(td1, td2, td3, td4);
            tbody.appendChild(tr);
        });

        goToArgeStep(3);
    });

    document.getElementById('argeGoTo4').addEventListener('click', () => {
        if (!argeRows.length) {
            showToast('Bitte zuerst die ARGE-Liste parsen (Schritt 2).');
            return;
        }
        const missing = argeRows.filter(r => !r.owner);
        if (missing.length) {
            showToast('Bitte für alle ARGEs einen Besitzer (UPN) eintragen.');
            return;
        }
        document.getElementById('argePowerShellScript').textContent = buildStandaloneArgePs1(false);
        goToArgeStep(4);
    });

    document.getElementById('argeCopyScript').addEventListener('click', () => {
        const t = document.getElementById('argePowerShellScript').textContent;
        navigator.clipboard.writeText(t).then(() => showToast('Script kopiert.'));
    });

    const argeLinesEl = document.getElementById('argeLines');
    if (argeLinesEl) {
        argeLinesEl.addEventListener('input', scheduleArgePreviewRefresh);
        argeLinesEl.addEventListener('paste', () => setTimeout(scheduleArgePreviewRefresh, 0));
    }
    ['argeDomain', 'argeDefaultPrefix'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.addEventListener('input', scheduleArgePreviewRefresh);
    });
    const argeUpperEl = document.getElementById('argeUpperNick');
    if (argeUpperEl) argeUpperEl.addEventListener('change', scheduleArgePreviewRefresh);

    // step header keyboard support
    document.querySelectorAll('.arge-steps .step').forEach(el => {
        el.setAttribute('tabindex', '0');
        el.addEventListener('click', () => {
            const s = parseFloat(el.dataset.argeStep);
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
})();

