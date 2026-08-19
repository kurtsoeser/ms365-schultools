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
export function buildStandaloneKursteamPs1(validTeams, escapeFn = psEscapeForExport) {
    const stamp = new Date().toISOString();
    const rows = buildKursteamTeamRows(validTeams, escapeFn);
    const lines = kursteamPsHeader(stamp, 'einfach');
    lines.push(buildKursteamLoginBlock());
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
export function buildKursteamCsvPreviewPs1() {
    return [
        '$TeamsList = Import-Csv -Path .\\neueteams.csv -Encoding UTF8',
        'Connect-MicrosoftTeams',
        '',
        '$i = 0; $skipped = 0; $failed = 0',
        'foreach ($Team in $TeamsList) {',
        '    $i++',
        '    try {',
        '        # Idempotenz: Team bereits vorhanden?',
        '        $existing = Get-Team -MailNickName $Team.Gruppenmail -ErrorAction SilentlyContinue',
        '        if ($existing) {',
        '            Write-Host "ÜBERSPRUNGEN [$i/$($TeamsList.Count)] $($Team.Gruppenmail) (existiert)" -ForegroundColor Yellow',
        '            $skipped++; continue',
        '        }',
        '        $null = New-Team -Template "EDU_Class" -DisplayName $Team.TeamName -MailNickName $Team.Gruppenmail -Owner $Team.Besitzer -ErrorAction Stop',
        '        Write-Host "OK [$i/$($TeamsList.Count)] $($Team.Gruppenmail)" -ForegroundColor Green',
        '    }',
        '    catch {',
        '        Write-Warning "Fehler [$i] $($Team.Gruppenmail): $($_.Exception.Message)"',
        '        $failed++',
        '    }',
        '    Start-Sleep -Seconds 2',
        '}',
        'Write-Host ""',
        'Write-Host "Fertig: $($i-$skipped-$failed) neu, $skipped übersprungen, $failed Fehler" -ForegroundColor Cyan'
    ].join('\r\n');
}

/** Empfohlener Generator: Checkpoint, Retry, ETA, Log. */
export function buildStandaloneKursteamPs1V2(validTeams, escapeFn = psEscapeForExport) {
    const stamp = new Date().toISOString();
    const rows = buildKursteamTeamRows(validTeams, escapeFn);
    const lines = kursteamPsHeader(stamp, 'empfohlen');
    lines.push('# Checkpoint/Resume, Retry bei Drosselung, Fortschritt/ETA, Log neben der CMD-Datei.');
    lines.push('');
    lines.push('$ScriptDir = if ($env:MS365_SELF) { Split-Path -Parent $env:MS365_SELF } else { (Get-Location).Path }');
    lines.push('$CheckpointPath = Join-Path $ScriptDir "Kursteam-Anlage-checkpoint.json"');
    lines.push('$LogPath = Join-Path $ScriptDir "Kursteam-Anlage.log"');
    lines.push('');
    lines.push('function Write-KtLog {');
    lines.push('    param([string]$Message, [ConsoleColor]$Color = [ConsoleColor]::White)');
    lines.push('    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"');
    lines.push('    try { Add-Content -LiteralPath $LogPath -Value ("[{0}] {1}" -f $ts, $Message) -Encoding UTF8 } catch { }');
    lines.push('    Write-Host $Message -ForegroundColor $Color');
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
    lines.push('                Write-KtLog ("  Drosselung – warte {0}s (Versuch {1}/{2})..." -f $wait, $attempt, $MaxAttempts) Yellow');
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
    lines.push('$TeamsList = @(');
    lines.push(rows.join(',\r\n'));
    lines.push(')');
    lines.push('');
    lines.push('Write-KtLog ("=== Kursteam-Anlage | {0} Teams ===" -f $TeamsList.Count) Cyan');
    lines.push('$estMin = [Math]::Max(1, [Math]::Ceiling($TeamsList.Count * 2.5 / 60))');
    lines.push('Write-KtLog ("Geschaetzte Skript-Laufzeit ca. {0} Min. (ohne Drosselung). Teams-Bereitstellung im Hintergrund kann laenger dauern." -f $estMin) DarkYellow');
    lines.push('Write-KtLog ("Checkpoint: {0}" -f $CheckpointPath) DarkGray');
    lines.push('Write-KtLog ("Log: {0}" -f $LogPath) DarkGray');
    lines.push('');
    lines.push('$cp = Get-KtCheckpoint');
    lines.push('$completedSet = @{}');
    lines.push('foreach ($m in @($cp.completed)) { if ($m) { $completedSet[$m] = $true } }');
    lines.push('if ($completedSet.Count -gt 0) {');
    lines.push('    Write-KtLog ("Checkpoint gefunden: {0} Teams bereits erledigt." -f $completedSet.Count) Cyan');
    lines.push('    $resume = Read-Host "Fortsetzen? [J/n] – n = Checkpoint loeschen und von vorn"');
    lines.push('    if ($resume -eq "n" -or $resume -eq "N") {');
    lines.push('        Remove-Item -LiteralPath $CheckpointPath -Force -ErrorAction SilentlyContinue');
    lines.push('        $completedSet = @{}');
    lines.push('        Write-KtLog "Checkpoint geloescht – Neustart." Yellow');
    lines.push('    }');
    lines.push('}');
    lines.push('');
    lines.push('$i = 0; $created = 0; $skipped = 0; $failed = 0; $fromCheckpoint = 0');
    lines.push('$pauseSec = 2');
    lines.push('$startTime = Get-Date');
    lines.push('foreach ($Team in $TeamsList) {');
    lines.push('    $i++');
    lines.push('    if ($completedSet.ContainsKey($Team.Gruppenmail)) {');
    lines.push('        Write-Host ("CHECKPOINT [{0}/{1}] {2}" -f $i, $TeamsList.Count, $Team.Gruppenmail) -ForegroundColor DarkGray');
    lines.push('        $fromCheckpoint++');
    lines.push('        continue');
    lines.push('    }');
    lines.push('    try {');
    lines.push('        $existing = Get-Team -MailNickName $Team.Gruppenmail -ErrorAction SilentlyContinue');
    lines.push('        if ($existing) {');
    lines.push('            Write-KtLog ("UEBERSPRUNGEN [{0}/{1}] {2} (existiert)" -f $i, $TeamsList.Count, $Team.Gruppenmail) Yellow');
    lines.push('            $skipped++');
    lines.push('            $completedSet[$Team.Gruppenmail] = $true');
    lines.push('            Save-KtCheckpoint -Completed @($completedSet.Keys) -Stats @{ created = $created; skipped = $skipped; failed = $failed }');
    lines.push('            continue');
    lines.push('        }');
    lines.push('        $result = Invoke-KtTeamCreateWithRetry -DisplayName $Team.TeamName -MailNickName $Team.Gruppenmail -Owner $Team.Besitzer');
    lines.push('        if ($result.ok) {');
    lines.push('            Write-KtLog ("OK [{0}/{1}] {2}" -f $i, $TeamsList.Count, $Team.Gruppenmail) Green');
    lines.push('            $created++');
    lines.push('            $completedSet[$Team.Gruppenmail] = $true');
    lines.push('            if ($pauseSec -gt 2) { $pauseSec = [Math]::Max(2, $pauseSec - 1) }');
    lines.push('        } else {');
    lines.push('            Write-KtLog ("FEHLER [{0}] {1}: {2}" -f $i, $Team.Gruppenmail, $result.message) Red');
    lines.push('            $failed++');
    lines.push('            $pauseSec = [Math]::Min(30, $pauseSec + 2)');
    lines.push('            if ($result.fatal) { break }');
    lines.push('        }');
    lines.push('    } catch {');
    lines.push('        Write-KtLog ("FEHLER [{0}] {1}: {2}" -f $i, $Team.Gruppenmail, $_.Exception.Message) Red');
    lines.push('        $failed++');
    lines.push('        $pauseSec = [Math]::Min(30, $pauseSec + 2)');
    lines.push('    }');
    lines.push('    Save-KtCheckpoint -Completed @($completedSet.Keys) -Stats @{ created = $created; skipped = $skipped; failed = $failed }');
    lines.push('    $doneSoFar = $created + $skipped + $failed + $fromCheckpoint');
    lines.push('    if ($doneSoFar -gt 0) {');
    lines.push('        $elapsed = ((Get-Date) - $startTime).TotalSeconds');
    lines.push('        $remaining = $TeamsList.Count - $i');
    lines.push('        $etaSec = [int](($elapsed / $doneSoFar) * $remaining)');
    lines.push('        Write-Host ("  -> Fortschritt {0}/{1} | ETA ca. {2:mm\\:ss} | Pause {3}s" -f $i, $TeamsList.Count, [TimeSpan]::FromSeconds($etaSec), $pauseSec) -ForegroundColor DarkGray');
    lines.push('    }');
    lines.push('    Start-Sleep -Seconds $pauseSec');
    lines.push('}');
    lines.push('');
    lines.push('Write-Host ""');
    lines.push('Write-KtLog ("Zusammenfassung: {0} neu, {1} uebersprungen (existierten), {2} aus Checkpoint, {3} Fehler" -f $created, $skipped, $fromCheckpoint, $failed) Cyan');
    lines.push('Write-Host ""');
    lines.push('Write-Host "Fertig. Bei Abbruch erneut starten – Checkpoint setzt fort." -ForegroundColor Cyan');
    lines.push('Read-Host "Enter druecken zum Beenden"');
    return lines.join('\r\n');
}
