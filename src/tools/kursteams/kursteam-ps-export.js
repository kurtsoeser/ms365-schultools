/**
 * PowerShell-Generatoren für Kursteam-Anlage (CMD/Polyglot).
 * v1 = bestehender Weg; v2 = Test-Phase mit Checkpoint, Retry und ETA.
 */

export function psEscapeForExport(s) {
    return String(s ?? '').replace(/'/g, "''");
}

function buildKursteamLoginBlock() {
    return [
        'Write-Host ""',
        'Write-Host "=== Anmeldung bei Microsoft Teams / Microsoft 365 ===" -ForegroundColor Cyan',
        'Write-Host "Konten mit MFA: bitte Option A waehlen (Browser-Anmeldung)." -ForegroundColor Yellow',
        'Write-Host ""',
        'Write-Host " [A] Interaktive Anmeldung (empfohlen, MFA moeglich)"',
        'Write-Host " [B] Benutzername + Passwort (Get-Credential) – oft nur ohne MFA zuverlaessig"',
        'Write-Host ""',
        '$loginChoice = Read-Host "Auswahl eingeben (A oder B, Standard A)"',
        'if ($loginChoice -eq "B" -or $loginChoice -eq "b") {',
        '    $script:Ms365Cred = Get-Credential -Message "Microsoft 365 / Teams Administrator"',
        '    if ($null -eq $script:Ms365Cred) { Write-Error "Anmeldung abgebrochen."; exit 1 }',
        '    Connect-MicrosoftTeams -Credential $script:Ms365Cred',
        '} else {',
        '    Connect-MicrosoftTeams',
        '}',
        ''
    ].join('\r\n');
}

function buildKursteamExchangeBlock(domainTrim) {
    if (!domainTrim) return { header: [], afterTeamOk: [] };
    const header = [
        '$Ms365SetExchangeSmtp = $true',
        "$Ms365ExchangeDomain = '" + domainTrim.replace(/'/g, "''") + "'",
        '$script:Ms365ExoConnected = $false',
        'function Ensure-KtExchangeOnline {',
        '    if ($script:Ms365ExoConnected) { return }',
        '    Write-Host "Exchange Online: Anmeldung für Schul-Domain …" -ForegroundColor Yellow',
        '    try { Import-Module ExchangeOnlineManagement -ErrorAction Stop } catch {',
        '        Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber',
        '        Import-Module ExchangeOnlineManagement -ErrorAction Stop',
        '    }',
        '    Connect-ExchangeOnline -ShowBanner:$false',
        '    $script:Ms365ExoConnected = $true',
        '}',
        ''
    ];
    const afterTeamOk = [
        '        if ($Ms365SetExchangeSmtp -and $Ms365ExchangeDomain) {',
        '            Ensure-KtExchangeOnline',
        '            $wantedSmtp = "$($Team.Gruppenmail)@$Ms365ExchangeDomain"',
        '            for ($ei = 0; $ei -lt 6; $ei++) {',
        '                try {',
        '                    $ktTeam = Get-Team -MailNickName $Team.Gruppenmail -ErrorAction Stop',
        '                    $prevWarn = $WarningPreference',
        '                    $WarningPreference = "SilentlyContinue"',
        '                    try {',
        '                        Set-UnifiedGroup -Identity $ktTeam.GroupId -PrimarySmtpAddress $wantedSmtp -ErrorAction Stop',
        '                    } finally { $WarningPreference = $prevWarn }',
        '                    Write-KtDetail ("Exchange OK: {0}" -f $wantedSmtp)',
        '                    break',
        '                } catch {',
        '                    if ($ei -lt 5) { Start-Sleep -Seconds 15 } else {',
        '                        Write-KtDetail ("Exchange FEHLER: {0}" -f $_.Exception.Message)',
        '                        if (Get-Command Clear-KtProgressLine -ErrorAction SilentlyContinue) { Clear-KtProgressLine }',
        '                        Write-Host ("         Exchange: PrimarySmtpAddress nicht gesetzt") -ForegroundColor DarkYellow',
        '                    }',
        '                }',
        '            }',
        '        }'
    ];
    return { header, afterTeamOk };
}

function buildKursteamTeamRows(validTeams, escapeFn) {
    return validTeams.map(t =>
        "    [PSCustomObject]@{ TeamName = '" +
            escapeFn(t.teamName) +
            "'; Gruppenmail = '" +
            escapeFn(t.gruppenmail) +
            "'; Besitzer = '" +
            escapeFn(t.besitzer) +
            "' }"
    );
}

function kursteamPsHeader(stamp, variantLabel) {
    const lines = [];
    lines.push('#Requires -Version 5.1');
    lines.push('# Kursteam-Anlage (Microsoft Teams, Vorlage EDU_Class)' + (variantLabel ? ' – ' + variantLabel : ''));
    lines.push('# Entspricht Microsoft Learn: New-Team -Template "EDU_Class" (gueltige Werte: EDU_Class, EDU_PLC).');
    lines.push('# Microsoft empfiehlt fuer Klassen-Teams das Modul MicrosoftTeams in Version 7.3.1 oder neuer.');
    lines.push('# Erzeugt in der Browser-App am ' + stamp);
    lines.push('# Daten sind unten eingebettet – keine separate CSV noetig.');
    lines.push('');
    lines.push('[Console]::OutputEncoding = [System.Text.Encoding]::UTF8');
    lines.push('$ErrorActionPreference = "Continue"');
    lines.push('# Unterdrueckt Write-Progress der MicrosoftTeams-Cmdlets (sonst "Fetching teams"-Spam).');
    lines.push('$ProgressPreference = "SilentlyContinue"');
    lines.push('function Write-KtDetail([string]$Message) { Write-Host ("  {0}" -f $Message) -ForegroundColor DarkGray }');
    lines.push('');
    lines.push('if (-not (Get-Module -ListAvailable -Name MicrosoftTeams)) {');
    lines.push('    Write-Host "Installiere Modul MicrosoftTeams (einmalig)..." -ForegroundColor Yellow');
    lines.push('    Install-Module MicrosoftTeams -Scope CurrentUser -Force');
    lines.push('}');
    lines.push('Import-Module MicrosoftTeams -ErrorAction Stop');
    lines.push('');
    return lines;
}

/** Einfache Variante ohne Checkpoint/Retry (Alternative für kleine Mengen). */
export function buildStandaloneKursteamPs1(validTeams, escapeFn = psEscapeForExport, domain = '') {
    const stamp = new Date().toISOString();
    const rows = buildKursteamTeamRows(validTeams, escapeFn);
    const exo = buildKursteamExchangeBlock(String(domain || '').trim());
    const lines = kursteamPsHeader(stamp, 'einfach');
    lines.push(buildKursteamLoginBlock());
    exo.header.forEach((l) => lines.push(l));
    lines.push('$TeamsList = @(');
    lines.push(rows.join(',\r\n'));
    lines.push(')');
    lines.push('');
    lines.push('$i = 0');
    lines.push('$skipped = 0');
    lines.push('$failed = 0');
    lines.push('foreach ($Team in $TeamsList) {');
    lines.push('    $i++');
    lines.push('    try {');
    lines.push('        # Idempotenz-Prüfung: Team bereits vorhanden?');
    lines.push('        $existing = Get-Team -MailNickName $Team.Gruppenmail -ErrorAction SilentlyContinue');
    lines.push('        if ($existing) {');
    lines.push('            Write-Host ("ÜBERSPRUNGEN [{0}/{1}] {2} (existiert bereits)" -f $i, $TeamsList.Count, $Team.Gruppenmail) -ForegroundColor Yellow');
    lines.push('            $skipped++');
    lines.push('            continue');
    lines.push('        }');
    lines.push('        $null = New-Team -Template "EDU_Class" -DisplayName $Team.TeamName -MailNickName $Team.Gruppenmail -Owner $Team.Besitzer -ErrorAction Stop');
    exo.afterTeamOk.forEach((l) => lines.push(l));
    lines.push('        Write-Host ("OK [{0}/{1}] {2}" -f $i, $TeamsList.Count, $Team.Gruppenmail) -ForegroundColor Green');
    lines.push('    }');
    lines.push('    catch {');
    lines.push('        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $Team.Gruppenmail, $_.Exception.Message)');
    lines.push('        $failed++');
    lines.push('    }');
    lines.push('    Start-Sleep -Seconds 2');
    lines.push('}');
    lines.push('');
    lines.push('Write-Host ""');
    lines.push('Write-Host ("Zusammenfassung: {0} neu angelegt, {1} übersprungen (existierten), {2} Fehler" -f ($i - $skipped - $failed), $skipped, $failed) -ForegroundColor Cyan');
    lines.push('');
    lines.push('Write-Host ""');
    lines.push('Write-Host "Fertig. Fenster schliesst nicht automatisch." -ForegroundColor Cyan');
    lines.push('Read-Host "Enter druecken zum Beenden"');
    return lines.join('\r\n');
}

/** CSV-basiertes Vorschau-Script (Schritt 7, ohne eingebettete Daten). */
export function buildKursteamCsvPreviewPs1(domain = '') {
    const exo = buildKursteamExchangeBlock(String(domain || '').trim());
    const lines = [
        '$TeamsList = Import-Csv -Path .\\neueteams.csv -Encoding UTF8',
        'Connect-MicrosoftTeams',
        ''
    ];
    exo.header.forEach((l) => lines.push(l));
    lines.push('$i = 0; $skipped = 0; $failed = 0');
    lines.push('foreach ($Team in $TeamsList) {');
    lines.push('    $i++');
    lines.push('    try {');
    lines.push('        # Idempotenz: Team bereits vorhanden?');
    lines.push('        $existing = Get-Team -MailNickName $Team.Gruppenmail -ErrorAction SilentlyContinue');
    lines.push('        if ($existing) {');
    lines.push('            Write-Host "ÜBERSPRUNGEN [$i/$($TeamsList.Count)] $($Team.Gruppenmail) (existiert)" -ForegroundColor Yellow');
    lines.push('            $skipped++; continue');
    lines.push('        }');
    lines.push('        $null = New-Team -Template "EDU_Class" -DisplayName $Team.TeamName -MailNickName $Team.Gruppenmail -Owner $Team.Besitzer -ErrorAction Stop');
    exo.afterTeamOk.forEach((l) => lines.push(l));
    lines.push('        Write-Host "OK [$i/$($TeamsList.Count)] $($Team.Gruppenmail)" -ForegroundColor Green');
    lines.push('    }');
    lines.push('    catch {');
    lines.push('        Write-Warning "Fehler [$i] $($Team.Gruppenmail): $($_.Exception.Message)"');
    lines.push('        $failed++');
    lines.push('    }');
    lines.push('    Start-Sleep -Seconds 2');
    lines.push('}');
    lines.push('Write-Host ""');
    lines.push('Write-Host "Fertig: $($i-$skipped-$failed) neu, $skipped übersprungen, $failed Fehler" -ForegroundColor Cyan');
    if (exo.header.length) {
        lines.push('if ($script:Ms365ExoConnected) {');
        lines.push('    try { Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue } catch {}');
        lines.push('}');
    }
    return lines.join('\r\n');
}

/** Empfohlener Generator: Checkpoint, Retry, ETA, Log. */
export function buildStandaloneKursteamPs1V2(validTeams, escapeFn = psEscapeForExport, domain = '') {
    const stamp = new Date().toISOString();
    const rows = buildKursteamTeamRows(validTeams, escapeFn);
    const exo = buildKursteamExchangeBlock(String(domain || '').trim());
    const lines = kursteamPsHeader(stamp, 'empfohlen');
    lines.push('# Checkpoint/Resume, Retry bei Drosselung, Fortschritt/ETA, Log neben der CMD-Datei.');
    lines.push('');
    lines.push('$ScriptDir = if ($env:MS365_SELF) { Split-Path -Parent $env:MS365_SELF } else { (Get-Location).Path }');
    lines.push('$CheckpointPath = Join-Path $ScriptDir "Kursteam-Anlage-checkpoint.json"');
    lines.push('$LogPath = Join-Path $ScriptDir "Kursteam-Anlage.log"');
    lines.push('');
    lines.push('$script:KtProgressActive = $false');
    lines.push('function Write-KtFile([string]$Message) {');
    lines.push('    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"');
    lines.push('    try { Add-Content -LiteralPath $LogPath -Value ("[{0}] {1}" -f $ts, $Message) -Encoding UTF8 } catch { }');
    lines.push('}');
    lines.push('function Clear-KtProgressLine {');
    lines.push('    if ($script:KtProgressActive) {');
    lines.push('        $w = 100');
    lines.push('        try { if ([Console]::WindowWidth -gt 2) { $w = [Console]::WindowWidth - 1 } } catch { }');
    lines.push('        Write-Host ("`r" + (" " * $w) + "`r") -NoNewline');
    lines.push('        $script:KtProgressActive = $false');
    lines.push('    }');
    lines.push('}');
    lines.push('function Format-KtEta([int]$EtaSec) {');
    lines.push('    if ($EtaSec -lt 0) { return "…" }');
    lines.push('    $h = [int][Math]::Floor($EtaSec / 3600)');
    lines.push('    $m = [int][Math]::Floor(($EtaSec % 3600) / 60)');
    lines.push('    $s = [int]($EtaSec % 60)');
    lines.push('    if ($h -gt 0) { return ("{0}h {1:D2}m" -f $h, $m) }');
    lines.push('    return ("{0:D2}:{1:D2}" -f $m, $s)');
    lines.push('}');
    lines.push('function Write-KtProgress {');
    lines.push('    param([int]$Index, [int]$Total, [int]$Ok, [int]$Skip, [int]$Fail, [int]$EtaSec, [int]$PauseSec)');
    lines.push('    $pct = if ($Total -gt 0) { [int](100 * $Index / $Total) } else { 0 }');
    lines.push('    $barLen = 20');
    lines.push('    $filled = [Math]::Max(0, [Math]::Min($barLen, [int]($barLen * $Index / [Math]::Max(1, $Total))))');
    lines.push('    $bar = ("#" * $filled) + ("-" * ($barLen - $filled))');
    lines.push('    $msg = ("[{0}] {1,3}% | {2}/{3} | OK:{4} Skip:{5} Err:{6} | ETA {7} | Pause {8}s" -f $bar, $pct, $Index, $Total, $Ok, $Skip, $Fail, (Format-KtEta $EtaSec), $PauseSec)');
    lines.push('    $w = 120');
    lines.push('    try { if ([Console]::WindowWidth -gt 2) { $w = [Console]::WindowWidth - 1 } } catch { }');
    lines.push('    if ($msg.Length -gt $w) { $msg = $msg.Substring(0, $w) }');
    lines.push('    Write-Host ("`r" + $msg) -NoNewline -ForegroundColor Cyan');
    lines.push('    $script:KtProgressActive = $true');
    lines.push('}');
    lines.push('function Write-KtEvent {');
    lines.push('    param([string]$Kind, [string]$Message, [ConsoleColor]$Color = [ConsoleColor]::White)');
    lines.push('    Clear-KtProgressLine');
    lines.push('    $line = ("{0,-8} {1}" -f $Kind, $Message)');
    lines.push('    Write-KtFile $line');
    lines.push('    Write-Host $line -ForegroundColor $Color');
    lines.push('}');
    lines.push('function Write-KtDetail([string]$Message) { Write-KtFile $Message }');
    lines.push('function Write-KtLog {');
    lines.push('    param([string]$Message, [ConsoleColor]$Color = [ConsoleColor]::White)');
    lines.push('    Write-KtEvent "INFO" $Message $Color');
    lines.push('}');
    lines.push('');
    lines.push('function Get-KtCheckpoint {');
    lines.push('    if (-not (Test-Path -LiteralPath $CheckpointPath)) { return @{ completed = @() } }');
    lines.push('    try {');
    lines.push('        $raw = Get-Content -LiteralPath $CheckpointPath -Raw -Encoding UTF8');
    lines.push('        return ($raw | ConvertFrom-Json)');
    lines.push('    } catch {');
    lines.push('        Write-Warning "Checkpoint unlesbar – starte ohne Fortsetzung."');
    lines.push('        return @{ completed = @() }');
    lines.push('    }');
    lines.push('}');
    lines.push('');
    lines.push('function Save-KtCheckpoint {');
    lines.push('    param([string[]]$Completed, [hashtable]$Stats)');
    lines.push('    $payload = @{');
    lines.push('        version = 2');
    lines.push('        completed = @($Completed | Sort-Object -Unique)');
    lines.push('        lastRun = (Get-Date).ToUniversalTime().ToString("o")');
    lines.push('        stats = $Stats');
    lines.push('    }');
    lines.push('    $payload | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $CheckpointPath -Encoding UTF8');
    lines.push('}');
    lines.push('');
    lines.push('function Invoke-KtTeamCreateWithRetry {');
    lines.push('    param(');
    lines.push('        [string]$DisplayName,');
    lines.push('        [string]$MailNickName,');
    lines.push('        [string]$Owner,');
    lines.push('        [int]$MaxAttempts = 6');
    lines.push('    )');
    lines.push('    $baseWait = 3');
    lines.push('    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {');
    lines.push('        try {');
    lines.push('            $null = New-Team -Template "EDU_Class" -DisplayName $DisplayName -MailNickName $MailNickName -Owner $Owner -ErrorAction Stop');
    lines.push('            return @{ ok = $true }');
    lines.push('        } catch {');
    lines.push('            $msg = $_.Exception.Message');
    lines.push('            if ($msg -match "AadGroupCreationLimitExceeded|250.*group|creation limit|Directory_ObjectQuotaExceeded") {');
    lines.push('                Write-KtLog "HINWEIS: Entra-ID Limit (ca. 250 Gruppen pro Nicht-Admin) erreicht. Bitte globales Admin-Konto nutzen." Red');
    lines.push('                return @{ ok = $false; fatal = $true; message = $msg }');
    lines.push('            }');
    lines.push('            $isThrottle = $msg -match "throttl|429|Too Many Requests|rate limit|Request_ThrottledTemporarily|service is busy"');
    lines.push('            if ($isThrottle -and $attempt -lt $MaxAttempts) {');
    lines.push('                $wait = [Math]::Min(120, [int]($baseWait * [Math]::Pow(2, $attempt - 1)))');
    lines.push('                Write-KtEvent "WARTE" ("Drosselung – {0}s (Versuch {1}/{2})" -f $wait, $attempt, $MaxAttempts) Yellow');
    lines.push('                Start-Sleep -Seconds $wait');
    lines.push('                continue');
    lines.push('            }');
    lines.push('            if ($attempt -lt $MaxAttempts -and $msg -match "timeout|temporarily|Unavailable|503") {');
    lines.push('                Start-Sleep -Seconds (5 * $attempt)');
    lines.push('                continue');
    lines.push('            }');
    lines.push('            return @{ ok = $false; fatal = $false; message = $msg }');
    lines.push('        }');
    lines.push('    }');
    lines.push('    return @{ ok = $false; fatal = $false; message = "Max. Versuche erreicht" }');
    lines.push('}');
    lines.push('');
    lines.push(buildKursteamLoginBlock());
    exo.header.forEach((l) => lines.push(l));
    lines.push('$TeamsList = @(');
    lines.push(rows.join(',\r\n'));
    lines.push(')');
    lines.push('');
    lines.push('Write-Host ("=== Kursteam-Anlage | {0} Teams ===" -f $TeamsList.Count) -ForegroundColor Cyan');
    lines.push('$estMin = [Math]::Max(1, [Math]::Ceiling($TeamsList.Count * 2.5 / 60))');
    lines.push('Write-Host ("Geschaetzte Skript-Laufzeit ca. {0} Min. (ohne Drosselung)." -f $estMin) -ForegroundColor Yellow');
    lines.push('Write-Host ("Checkpoint: {0}" -f $CheckpointPath) -ForegroundColor DarkGray');
    lines.push('Write-Host ("Log: {0}  (Exchange-Details nur hier)" -f $LogPath) -ForegroundColor DarkGray');
    lines.push('Write-Host "Konsole: eine Fortschrittszeile + dauerhafte OK / SKIP / FEHLER." -ForegroundColor DarkGray');
    lines.push('Write-KtFile ("=== Start | {0} Teams ===" -f $TeamsList.Count)');
    lines.push('');
    lines.push('$cp = Get-KtCheckpoint');
    lines.push('$completedSet = @{}');
    lines.push('foreach ($m in @($cp.completed)) { if ($m) { $completedSet[$m] = $true } }');
    lines.push('if ($completedSet.Count -gt 0) {');
    lines.push('    Write-Host ("Checkpoint gefunden: {0} Teams bereits erledigt." -f $completedSet.Count) -ForegroundColor Cyan');
    lines.push('    $resume = Read-Host "Fortsetzen? [J/n] – n = Checkpoint loeschen und von vorn"');
    lines.push('    if ($resume -eq "n" -or $resume -eq "N") {');
    lines.push('        Remove-Item -LiteralPath $CheckpointPath -Force -ErrorAction SilentlyContinue');
    lines.push('        $completedSet = @{}');
    lines.push('        Write-KtEvent "INFO" "Checkpoint geloescht – Neustart." Yellow');
    lines.push('    }');
    lines.push('}');
    lines.push('');
    lines.push('$i = 0; $created = 0; $skipped = 0; $failed = 0; $fromCheckpoint = 0');
    lines.push('$workDone = 0');
    lines.push('$pauseSec = 2');
    lines.push('$startTime = Get-Date');
    lines.push('$workStartTime = $null');
    lines.push('function Update-KtEta {');
    lines.push('    if ($workDone -le 0 -or $null -eq $workStartTime) { return -1 }');
    lines.push('    $elapsed = ((Get-Date) - $workStartTime).TotalSeconds');
    lines.push('    $remaining = $TeamsList.Count - $i');
    lines.push('    if ($elapsed -le 0 -or $remaining -le 0) { if ($remaining -le 0) { return 0 } else { return -1 } }');
    lines.push('    return [int](($elapsed / $workDone) * $remaining)');
    lines.push('}');
    lines.push('foreach ($Team in $TeamsList) {');
    lines.push('    $i++');
    lines.push('    if ($completedSet.ContainsKey($Team.Gruppenmail)) {');
    lines.push('        Write-KtFile ("CHECKPOINT [{0}/{1}] {2}" -f $i, $TeamsList.Count, $Team.Gruppenmail)');
    lines.push('        $fromCheckpoint++');
    lines.push('        Write-KtProgress -Index $i -Total $TeamsList.Count -Ok $created -Skip ($skipped + $fromCheckpoint) -Fail $failed -EtaSec (Update-KtEta) -PauseSec $pauseSec');
    lines.push('        continue');
    lines.push('    }');
    lines.push('    if ($null -eq $workStartTime) { $workStartTime = Get-Date }');
    lines.push('    try {');
    lines.push('        $existing = Get-Team -MailNickName $Team.Gruppenmail -ErrorAction SilentlyContinue');
    lines.push('        if ($existing) {');
    lines.push('            Write-KtEvent "SKIP" ("[{0}/{1}] {2} (existiert)" -f $i, $TeamsList.Count, $Team.TeamName) DarkYellow');
    lines.push('            $skipped++');
    lines.push('            $workDone++');
    lines.push('            $completedSet[$Team.Gruppenmail] = $true');
    lines.push('            Save-KtCheckpoint -Completed @($completedSet.Keys) -Stats @{ created = $created; skipped = $skipped; failed = $failed }');
    lines.push('            Write-KtProgress -Index $i -Total $TeamsList.Count -Ok $created -Skip ($skipped + $fromCheckpoint) -Fail $failed -EtaSec (Update-KtEta) -PauseSec $pauseSec');
    lines.push('            continue');
    lines.push('        }');
    lines.push('        $result = Invoke-KtTeamCreateWithRetry -DisplayName $Team.TeamName -MailNickName $Team.Gruppenmail -Owner $Team.Besitzer');
    lines.push('        if ($result.ok) {');
    exo.afterTeamOk.forEach((l) => lines.push(l));
    lines.push('            Write-KtEvent "OK" ("[{0}/{1}] {2}" -f $i, $TeamsList.Count, $Team.TeamName) Green');
    lines.push('            $created++');
    lines.push('            $workDone++');
    lines.push('            $completedSet[$Team.Gruppenmail] = $true');
    lines.push('            if ($pauseSec -gt 2) { $pauseSec = [Math]::Max(2, $pauseSec - 1) }');
    lines.push('        } else {');
    lines.push('            Write-KtEvent "FEHLER" ("[{0}/{1}] {2}: {3}" -f $i, $TeamsList.Count, $Team.TeamName, $result.message) Red');
    lines.push('            $failed++');
    lines.push('            $workDone++');
    lines.push('            $pauseSec = [Math]::Min(30, $pauseSec + 2)');
    lines.push('            if ($result.fatal) { break }');
    lines.push('        }');
    lines.push('    } catch {');
    lines.push('        Write-KtEvent "FEHLER" ("[{0}/{1}] {2}: {3}" -f $i, $TeamsList.Count, $Team.TeamName, $_.Exception.Message) Red');
    lines.push('        $failed++');
    lines.push('        $workDone++');
    lines.push('        $pauseSec = [Math]::Min(30, $pauseSec + 2)');
    lines.push('    }');
    lines.push('    Save-KtCheckpoint -Completed @($completedSet.Keys) -Stats @{ created = $created; skipped = $skipped; failed = $failed }');
    lines.push('    Write-KtProgress -Index $i -Total $TeamsList.Count -Ok $created -Skip ($skipped + $fromCheckpoint) -Fail $failed -EtaSec (Update-KtEta) -PauseSec $pauseSec');
    lines.push('    Start-Sleep -Seconds $pauseSec');
    lines.push('}');
    lines.push('');
    lines.push('Clear-KtProgressLine');
    lines.push('Write-Host ""');
    lines.push('Write-KtEvent "FERTIG" ("{0} neu | {1} uebersprungen | {2} Checkpoint | {3} Fehler | Dauer {4}" -f $created, $skipped, $fromCheckpoint, $failed, ((Get-Date) - $startTime).ToString("hh\\:mm\\:ss")) Cyan');
    lines.push('Write-Host "Bei Abbruch erneut starten – Checkpoint setzt fort." -ForegroundColor DarkGray');
    lines.push('Read-Host "Enter druecken zum Beenden"');
    return lines.join('\r\n');
}
