/**
 * Exchange-Online-PowerShell: Kalender-Postfächer (Shared Mailbox)
 * anlegen, Kalender-Ordner-Rechte setzen, optional iCal-Publish.
 */

import { normalizeSharedMailboxDelegates } from '../../shared/exchange-shared-mailbox-script.js';

/** Marker: alles danach ist PowerShell (für doppelklickbare .cmd). */
export const CALENDAR_POSTFAECHER_CMD_PS_MARKER = '___MS365_CALENDAR_POSTFAECHER_PS___';

function psSingleQuote(value) {
    return "'" + String(value == null ? '' : value).replace(/'/g, "''") + "'";
}

function psHereString(value) {
    const safe = String(value == null ? '' : value).replace(/^\s*'@\s*$/gm, "' @");
    return "@'\n" + safe + "\n'@";
}

/**
 * @param {unknown} raw
 * @returns {string[]}
 */
export function normalizeCalendarPrincipals(raw) {
    return normalizeSharedMailboxDelegates(raw);
}

/**
 * @param {unknown} item
 */
export function normalizeCalendarPlanItem(item) {
    const o = item && typeof item === 'object' ? item : {};
    const mode = String(o.mode || 'create').toLowerCase() === 'permissionsonly' ? 'permissionsOnly' : 'create';
    const editors = normalizeCalendarPrincipals(o.editors);
    const readersRaw = normalizeCalendarPrincipals(o.readers);
    const editorSet = new Set(editors.map((e) => e.toLowerCase()));
    const readers = readersRaw.filter((r) => !editorSet.has(r.toLowerCase()));
    return {
        displayName: String(o.displayName || '').trim(),
        alias: String(o.alias || '').trim(),
        primarySmtp: String(o.primarySmtp || '').trim(),
        description: String(o.description || '').trim(),
        mode: /** @type {'create'|'permissionsOnly'} */ (mode),
        editors: editors,
        readers: readers,
        sendAs: o.sendAs !== false,
        publish: o.publish === true,
        autoMapping: o.autoMapping !== false
    };
}

/**
 * @param {unknown[]} items
 * @param {{ suggestedAdminUpn?: string }} [options]
 */
export function normalizeCalendarPlan(items, options) {
    const list = Array.isArray(items) ? items : [];
    const normalized = list.map(normalizeCalendarPlanItem).filter(function (it) {
        return it.displayName && it.alias;
    });
    const o = options && typeof options === 'object' ? options : {};
    return {
        items: normalized,
        suggestedAdminUpn: String(o.suggestedAdminUpn || '').trim()
    };
}

/**
 * Geführter Batch-Assistent für alle geplanten Kalender-Postfächer.
 * @param {unknown[]} items
 * @param {{ suggestedAdminUpn?: string }} [options]
 */
export function buildCalendarPostfaecherScript(items, options) {
    const plan = normalizeCalendarPlan(items, options);
    const lines = [];

    lines.push('#Requires -Version 5.1');
    lines.push('# Kalender-Postfächer – geführter Assistent (MS365-Schul-Tools)');
    lines.push('# Shared Mailbox + Kalender-Ordner-Rechte + optional iCal-Publish');
    lines.push('# Doppelklick: .cmd-Datei. Oder dieses Skript in PowerShell speichern/einfügen.');
    lines.push('');
    lines.push('$ErrorActionPreference = "Stop"');
    lines.push('try { [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false) } catch {}');
    lines.push('$Host.UI.RawUI.WindowTitle = "Kalender-Postfaecher anlegen"');
    lines.push('');
    lines.push('function Write-Step([string]$Text) {');
    lines.push('  Write-Host ""');
    lines.push('  Write-Host (">>> " + $Text) -ForegroundColor Cyan');
    lines.push('}');
    lines.push('function Write-Ok([string]$Text) { Write-Host ("  OK  " + $Text) -ForegroundColor Green }');
    lines.push('function Write-WarnLine([string]$Text) { Write-Host ("  !!  " + $Text) -ForegroundColor Yellow }');
    lines.push('function Pause-End([string]$Msg) {');
    lines.push('  Write-Host ""');
    lines.push('  if ($Msg) { Write-Host $Msg -ForegroundColor DarkGray }');
    lines.push('  Read-Host "Eingabetaste zum Beenden" | Out-Null');
    lines.push('}');
    lines.push('');
    lines.push('function Resolve-CalendarFolderIdentity([string]$MailboxIdentity) {');
    lines.push('  foreach ($leaf in @("Kalender", "Calendar")) {');
    lines.push('    $id = $MailboxIdentity + ":\\" + $leaf');
    lines.push('    try {');
    lines.push('      Get-MailboxFolderPermission -Identity $id -ErrorAction Stop | Out-Null');
    lines.push('      return $id');
    lines.push('    } catch {}');
    lines.push('  }');
    lines.push('  return ($MailboxIdentity + ":\\Calendar")');
    lines.push('}');
    lines.push('');
    lines.push('function Set-CalendarFolderRight([string]$FolderId, [string]$User, [string]$AccessRights) {');
    lines.push('  try {');
    lines.push(
        '    Add-MailboxFolderPermission -Identity $FolderId -User $User -AccessRights $AccessRights -ErrorAction Stop | Out-Null'
    );
    lines.push('  } catch {');
    lines.push(
        '    Set-MailboxFolderPermission -Identity $FolderId -User $User -AccessRights $AccessRights -ErrorAction Continue | Out-Null'
    );
    lines.push('  }');
    lines.push('}');
    lines.push('');
    lines.push('$SuggestedAdminUpn = ' + psSingleQuote(plan.suggestedAdminUpn));
    lines.push('$Calendars = @(');
    plan.items.forEach(function (it, idx) {
        const comma = idx < plan.items.length - 1 ? ',' : '';
        lines.push('  [pscustomobject]@{');
        lines.push('    DisplayName = ' + psSingleQuote(it.displayName));
        lines.push('    Alias = ' + psSingleQuote(it.alias));
        lines.push('    PrimarySmtpAddress = ' + psSingleQuote(it.primarySmtp));
        lines.push('    Description = ' + psHereString(it.description));
        lines.push('    Mode = ' + psSingleQuote(it.mode));
        lines.push(
            '    Editors = @(' +
                (it.editors.length ? it.editors.map(psSingleQuote).join(', ') : '') +
                ')'
        );
        lines.push(
            '    Readers = @(' +
                (it.readers.length ? it.readers.map(psSingleQuote).join(', ') : '') +
                ')'
        );
        lines.push('    SendAs = $' + (it.sendAs ? 'true' : 'false'));
        lines.push('    Publish = $' + (it.publish ? 'true' : 'false'));
        lines.push('    AutoMapping = $' + (it.autoMapping ? 'true' : 'false'));
        lines.push('  }' + comma);
    });
    lines.push(')');
    lines.push('');
    lines.push('Write-Host ""');
    lines.push('Write-Host "============================================" -ForegroundColor Cyan');
    lines.push('Write-Host "  Kalender-Postfaecher" -ForegroundColor Cyan');
    lines.push('Write-Host "  MS365-Schul-Tools" -ForegroundColor DarkCyan');
    lines.push('Write-Host "============================================" -ForegroundColor Cyan');
    lines.push('Write-Host ""');
    lines.push('Write-Host ("Geplant: " + $Calendars.Count + " Kalender") -ForegroundColor White');
    lines.push('foreach ($c in $Calendars) {');
    lines.push(
        '  Write-Host ("  - " + $c.DisplayName + " (" + $c.Alias + ") [" + $c.Mode + "]") -ForegroundColor DarkGray'
    );
    lines.push('}');
    lines.push('Write-Host ""');
    lines.push(
        'Write-Host "Wichtig: Mit einem Exchange-Admin der RICHTIGEN Schule / des richtigen Tenants anmelden." -ForegroundColor Yellow'
    );
    lines.push('');
    lines.push('Write-Step "1/4  Admin-Konto wählen"');
    lines.push('if (-not [string]::IsNullOrWhiteSpace($SuggestedAdminUpn)) {');
    lines.push('  Write-Host ("  Vorschlag: " + $SuggestedAdminUpn) -ForegroundColor DarkGray');
    lines.push('}');
    lines.push(
        '$AdminUpn = Read-Host "  UPN / E-Mail für die Anmeldung (leer = interaktive Microsoft-Anmeldung)"'
    );
    lines.push(
        'if ([string]::IsNullOrWhiteSpace($AdminUpn) -and -not [string]::IsNullOrWhiteSpace($SuggestedAdminUpn)) {'
    );
    lines.push('  $useSug = Read-Host ("  Vorschlag verwenden? [" + $SuggestedAdminUpn + "]  J/n")');
    lines.push('  if ($useSug -notmatch "^[nN]") { $AdminUpn = $SuggestedAdminUpn }');
    lines.push('}');
    lines.push('if ($AdminUpn) { $AdminUpn = $AdminUpn.Trim() }');
    lines.push('');
    lines.push('Write-Step "2/4  Bestätigung"');
    lines.push('$confirm = Read-Host "  Kalender jetzt in diesem Tenant anlegen/aktualisieren? (J/N)"');
    lines.push('if ($confirm -notmatch "^[jJyY]") {');
    lines.push('  Write-WarnLine "Abgebrochen – es wurde nichts geändert."');
    lines.push('  Pause-End');
    lines.push('  return');
    lines.push('}');
    lines.push('');
    lines.push('Write-Step "3/4  Exchange-Online-Modul"');
    lines.push('if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {');
    lines.push('  Write-WarnLine "Modul ExchangeOnlineManagement fehlt."');
    lines.push('  $inst = Read-Host "  Jetzt für diesen Benutzer installieren? (J/N)"');
    lines.push('  if ($inst -match "^[jJyY]") {');
    lines.push(
        '    Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop'
    );
    lines.push('    Write-Ok "Modul installiert."');
    lines.push('  } else {');
    lines.push(
        '    throw "Bitte zuerst ausführen: Install-Module ExchangeOnlineManagement -Scope CurrentUser"'
    );
    lines.push('  }');
    lines.push('} else {');
    lines.push('  Write-Ok "Modul vorhanden."');
    lines.push('}');
    lines.push('Import-Module ExchangeOnlineManagement -ErrorAction Stop');
    lines.push('');
    lines.push('Write-Step "4/4  Anmelden bei Exchange Online"');
    lines.push('try {');
    lines.push('  if (-not [string]::IsNullOrWhiteSpace($AdminUpn)) {');
    lines.push('    Write-Host ("  Anmelden als: " + $AdminUpn) -ForegroundColor DarkCyan');
    lines.push('    Connect-ExchangeOnline -UserPrincipalName $AdminUpn -ShowBanner:$false | Out-Null');
    lines.push('  } else {');
    lines.push('    Write-Host "  Microsoft-Anmeldefenster öffnet sich …" -ForegroundColor DarkCyan');
    lines.push('    Connect-ExchangeOnline -ShowBanner:$false | Out-Null');
    lines.push('  }');
    lines.push('  Write-Ok "Verbunden."');
    lines.push('} catch {');
    lines.push('  Write-Host ("  Anmeldung fehlgeschlagen: " + $_.Exception.Message) -ForegroundColor Red');
    lines.push('  Pause-End "Prüfe Konto, MFA und ob Exchange-Admin-Rechte vorliegen."');
    lines.push('  throw');
    lines.push('}');
    lines.push('');
    lines.push('$PublishResults = New-Object System.Collections.Generic.List[object]');
    lines.push('try {');
    lines.push('  if ($Calendars.Count -eq 0) { throw "Keine Kalender in der Planung." }');
    lines.push('  $n = 0');
    lines.push('  foreach ($cal in $Calendars) {');
    lines.push('    $n++');
    lines.push(
        '    Write-Host ""'
    );
    lines.push(
        '    Write-Host ("== Kalender " + $n + "/" + $Calendars.Count + ": " + $cal.DisplayName + " ==") -ForegroundColor Cyan'
    );
    lines.push('    if ([string]::IsNullOrWhiteSpace($cal.DisplayName)) { throw "Anzeigename fehlt." }');
    lines.push('    if ([string]::IsNullOrWhiteSpace($cal.Alias)) { throw "Alias fehlt." }');
    lines.push(
        '    $MailboxIdentity = if ([string]::IsNullOrWhiteSpace($cal.PrimarySmtpAddress)) { $cal.Alias } else { $cal.PrimarySmtpAddress }'
    );
    lines.push('');
    lines.push('    if ($cal.Mode -ne "permissionsOnly") {');
    lines.push('      $existing = $null');
    lines.push('      try { $existing = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop } catch {}');
    lines.push('      if ($existing) {');
    lines.push('        Write-WarnLine ("Postfach existiert bereits: " + $existing.PrimarySmtpAddress)');
    lines.push('        $MailboxIdentity = $existing.PrimarySmtpAddress');
    lines.push('      } else {');
    lines.push('        $newMailboxParams = @{');
    lines.push('          Shared = $true');
    lines.push('          Name = $cal.DisplayName');
    lines.push('          DisplayName = $cal.DisplayName');
    lines.push('          Alias = $cal.Alias');
    lines.push('        }');
    lines.push('        if (-not [string]::IsNullOrWhiteSpace($cal.PrimarySmtpAddress)) {');
    lines.push('          $newMailboxParams.PrimarySmtpAddress = $cal.PrimarySmtpAddress');
    lines.push('        }');
    lines.push('        Write-Host ("  Erstelle Shared Mailbox: " + $cal.DisplayName) -ForegroundColor DarkCyan');
    lines.push('        New-Mailbox @newMailboxParams | Out-Null');
    lines.push('        Write-Host "  Warte auf Verfügbarkeit …" -ForegroundColor DarkCyan');
    lines.push('        $mailbox = $null');
    lines.push('        for ($i = 1; $i -le 24; $i++) {');
    lines.push('          try {');
    lines.push('            $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop');
    lines.push('            break');
    lines.push('          } catch { Start-Sleep -Seconds 5 }');
    lines.push('        }');
    lines.push(
        '        if (-not $mailbox) { throw ("Postfach angelegt, aber noch nicht abrufbar: " + $MailboxIdentity) }'
    );
    lines.push('        $MailboxIdentity = $mailbox.PrimarySmtpAddress');
    lines.push('        Write-Ok ("Postfach bereit: " + $MailboxIdentity)');
    lines.push('      }');
    lines.push('    } else {');
    lines.push('      $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop');
    lines.push('      $MailboxIdentity = $mailbox.PrimarySmtpAddress');
    lines.push('      Write-Ok ("Bestehendes Postfach: " + $MailboxIdentity)');
    lines.push('    }');
    lines.push('');
    lines.push('    if (-not [string]::IsNullOrWhiteSpace($cal.Description)) {');
    lines.push('      Set-Mailbox -Identity $MailboxIdentity -Notes $cal.Description -ErrorAction Continue');
    lines.push('    }');
    lines.push('');
    lines.push('    $folderId = Resolve-CalendarFolderIdentity -MailboxIdentity $MailboxIdentity');
    lines.push('    Write-Host ("  Kalenderordner: " + $folderId) -ForegroundColor DarkGray');
    lines.push('');
    lines.push('    foreach ($d in $cal.Editors) {');
    lines.push(
        '      Add-MailboxPermission -Identity $MailboxIdentity -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$($cal.AutoMapping) -ErrorAction Continue | Out-Null'
    );
    lines.push('      Set-CalendarFolderRight -FolderId $folderId -User $d -AccessRights Editor');
    lines.push('      if ($cal.SendAs) {');
    lines.push(
        '        Add-RecipientPermission -Identity $MailboxIdentity -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Continue | Out-Null'
    );
    lines.push('      }');
    lines.push('      Write-Ok ("Eintragen (Editor + FullAccess) → " + $d)');
    lines.push('    }');
    lines.push('');
    lines.push('    foreach ($d in $cal.Readers) {');
    lines.push('      Set-CalendarFolderRight -FolderId $folderId -User $d -AccessRights Reviewer');
    lines.push('      Write-Ok ("Nur lesen (Reviewer) → " + $d)');
    lines.push('    }');
    lines.push('');
    lines.push('    if ($cal.Publish) {');
    lines.push('      try {');
    lines.push(
        '        Set-MailboxCalendarFolder -Identity $folderId -PublishEnabled $true -DetailLevel LimitedDetails -PublishDateRange ThreeMonths -ErrorAction Stop'
    );
    lines.push('        $pub = Get-MailboxCalendarFolder -Identity $folderId -ErrorAction Stop');
    lines.push('        $url = [string]$pub.PublishedCalendarUrl');
    lines.push('        if ([string]::IsNullOrWhiteSpace($url)) { $url = "(URL noch leer – später erneut Get-MailboxCalendarFolder)" }');
    lines.push('        Write-Ok ("Publish aktiv: " + $url)');
    lines.push(
        '        $PublishResults.Add([pscustomobject]@{ DisplayName = $cal.DisplayName; Mailbox = $MailboxIdentity; PublishedCalendarUrl = $url })'
    );
    lines.push('      } catch {');
    lines.push(
        '        Write-WarnLine ("Publish fehlgeschlagen: " + $_.Exception.Message + " (Tenant-Policy / Outlook-Web prüfen)")'
    );
    lines.push('      }');
    lines.push('    }');
    lines.push('  }');
    lines.push('');
    lines.push('  Write-Host ""');
    lines.push('  Write-Host "== Ergebnis / iCal-Links ==" -ForegroundColor Cyan');
    lines.push('  if ($PublishResults.Count -eq 0) {');
    lines.push('    Write-Host "  Keine Publish-URLs (Publish war aus oder fehlgeschlagen)." -ForegroundColor DarkGray');
    lines.push('  } else {');
    lines.push('    $PublishResults | Format-Table -AutoSize');
    lines.push('    Write-Host "Diese URLs können auf der Schulwebsite / für Abonnements genutzt werden." -ForegroundColor Green');
    lines.push('  }');
    lines.push('  Write-Host ""');
    lines.push('  Write-Host "Fertig. Outlook: Kalender hinzufügen → aus Verzeichnis." -ForegroundColor Green');
    lines.push('} catch {');
    lines.push('  Write-Host ""');
    lines.push('  Write-Host ("Fehler: " + $_.Exception.Message) -ForegroundColor Red');
    lines.push('  throw');
    lines.push('} finally {');
    lines.push('  try { Disconnect-ExchangeOnline -Confirm:$false | Out-Null } catch {}');
    lines.push('  Pause-End');
    lines.push('}');

    return lines.join('\n');
}

/**
 * Doppelklickbare .cmd mit eingebettetem PowerShell-Assistenten.
 * @param {unknown[]} items
 * @param {{ suggestedAdminUpn?: string }} [options]
 */
export function buildCalendarPostfaecherCmd(items, options) {
    const ps = buildCalendarPostfaecherScript(items, options);
    const withEnv =
        '@echo off\r\n' +
        'setlocal EnableExtensions\r\n' +
        'chcp 65001 >nul\r\n' +
        'title Kalender-Postfaecher anlegen\r\n' +
        'cd /d "%~dp0"\r\n' +
        'set "MS365_CMD_PATH=%~f0"\r\n' +
        'echo.\r\n' +
        'echo  Starte gefuehrten Assistenten (Exchange Online) ...\r\n' +
        'echo  Bitte danach das Admin-Konto der richtigen Schule / des Tenants waehlen.\r\n' +
        'echo.\r\n' +
        'powershell.exe -NoProfile -ExecutionPolicy Bypass -NoLogo -Command ^\r\n' +
        '  "$ErrorActionPreference=\'Stop\';' +
        ' $path=$env:MS365_CMD_PATH;' +
        ' if([string]::IsNullOrWhiteSpace($path)){ throw \'.cmd-Pfad fehlt.\' };' +
        ' $lines=Get-Content -LiteralPath $path -Encoding UTF8;' +
        " $marker='" +
        CALENDAR_POSTFAECHER_CMD_PS_MARKER +
        "';" +
        ' $i=0; while($i -lt $lines.Count -and $lines[$i].Trim() -ne $marker){ $i++ };' +
        ' if($i -ge $lines.Count){ throw \'PowerShell-Teil in .cmd nicht gefunden.\' };' +
        ' $script=($lines[($i+1)..($lines.Count-1)] -join [Environment]::NewLine);' +
        " $ps1=Join-Path $env:TEMP ('ms365-kalender-postfaecher-' + [guid]::NewGuid().ToString('N') + '.ps1');" +
        ' Set-Content -LiteralPath $ps1 -Value $script -Encoding UTF8;' +
        ' try { & powershell.exe -NoProfile -ExecutionPolicy Bypass -NoLogo -File $ps1; exit $LASTEXITCODE }' +
        ' finally { Remove-Item -LiteralPath $ps1 -Force -ErrorAction SilentlyContinue }"\r\n' +
        'set "EC=%ERRORLEVEL%"\r\n' +
        'endlocal & exit /b %EC%\r\n' +
        '\r\n' +
        CALENDAR_POSTFAECHER_CMD_PS_MARKER +
        '\r\n';

    return withEnv + ps.replace(/\r?\n/g, '\r\n') + '\r\n';
}
