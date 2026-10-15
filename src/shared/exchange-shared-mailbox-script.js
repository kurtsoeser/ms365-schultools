/**
 * Exchange-Online-PowerShell zum Anlegen freigegebener Postfächer
 * (Graph kann Shared Mailboxes nicht zuverlässig anlegen).
 *
 * Liefert einen geführten Assistenten und eine doppelklickbare .cmd-Datei.
 */

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
export function normalizeSharedMailboxDelegates(raw) {
    const list = Array.isArray(raw) ? raw : [];
    /** @type {Set<string>} */
    const seen = new Set();
    /** @type {string[]} */
    const out = [];
    list.forEach(function (entry) {
        let mail = '';
        if (entry && typeof entry === 'object') {
            mail = String(entry.email || entry.upn || entry.userPrincipalName || entry.mail || '').trim();
        } else {
            mail = String(entry == null ? '' : entry).trim();
        }
        if (!mail || mail.indexOf('@') === -1) return;
        const key = mail.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        out.push(mail);
    });
    return out;
}

/**
 * @param {unknown} raw
 * @returns {Array<'FullAccess'|'SendAs'|'SendOnBehalf'>}
 */
export function normalizeSharedMailboxRights(raw) {
    const list = Array.isArray(raw) ? raw : [];
    /** @type {Array<'FullAccess'|'SendAs'|'SendOnBehalf'>} */
    const out = [];
    /** @type {Set<string>} */
    const seen = new Set();
    list.forEach(function (r) {
        const key = String(r || '')
            .trim()
            .toLowerCase();
        if (key === 'fullaccess' && !seen.has('FullAccess')) {
            seen.add('FullAccess');
            out.push('FullAccess');
        } else if (key === 'sendas' && !seen.has('SendAs')) {
            seen.add('SendAs');
            out.push('SendAs');
        } else if (
            (key === 'sendonbehalf' || key === 'send_on_behalf') &&
            !seen.has('SendOnBehalf')
        ) {
            seen.add('SendOnBehalf');
            out.push('SendOnBehalf');
        }
    });
    return out;
}

/**
 * @param {{
 *   displayName?: string,
 *   alias?: string,
 *   primarySmtp?: string,
 *   description?: string,
 *   hideFromAddressLists?: boolean,
 *   delegates?: Array<string|{ email?: string, upn?: string, userPrincipalName?: string, mail?: string }>,
 *   rights?: Array<string>,
 *   autoMapping?: boolean,
 *   suggestedAdminUpn?: string
 * }} options
 */
function resolveMailboxOptions(options) {
    const o = options && typeof options === 'object' ? options : {};
    const displayName = String(o.displayName || '').trim();
    const alias = String(o.alias || '').trim();
    const primarySmtp = String(o.primarySmtp || '').trim();
    const description = String(o.description || '').trim();
    const hideFromAddressLists = o.hideFromAddressLists === true;
    const delegates = normalizeSharedMailboxDelegates(o.delegates);
    const rights =
        o.rights == null
            ? delegates.length
                ? /** @type {Array<'FullAccess'|'SendAs'>} */ (['FullAccess', 'SendAs'])
                : []
            : normalizeSharedMailboxRights(o.rights);
    const autoMapping = o.autoMapping !== false;
    const suggestedAdminUpn = String(o.suggestedAdminUpn || '').trim();
    return {
        displayName: displayName,
        alias: alias,
        primarySmtp: primarySmtp,
        description: description,
        hideFromAddressLists: hideFromAddressLists,
        delegates: delegates,
        rights: rights,
        autoMapping: autoMapping,
        suggestedAdminUpn: suggestedAdminUpn
    };
}

/** Marker: alles danach ist PowerShell (für doppelklickbare .cmd). */
export const SHARED_MAILBOX_CMD_PS_MARKER = '___MS365_SHARED_MAILBOX_PS___';

/**
 * Geführter PowerShell-Assistent: Konto fragen, bestätigen, Modul prüfen, anlegen, Rechte setzen.
 * @param {Parameters<typeof resolveMailboxOptions>[0]} options
 */
export function buildSharedMailboxCreateScript(options) {
    const o = resolveMailboxOptions(options);
    const lines = [];

    lines.push('#Requires -Version 5.1');
    lines.push('# Freigegebenes Postfach – geführter Assistent (MS365-Schul-Tools / Organigramm)');
    lines.push('# Doppelklick: .cmd-Datei. Oder dieses Skript in PowerShell speichern/einfügen.');
    lines.push('');
    lines.push('$ErrorActionPreference = "Stop"');
    lines.push('try { [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false) } catch {}');
    lines.push('$Host.UI.RawUI.WindowTitle = "Freigegebenes Postfach anlegen"');
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
    lines.push('$DisplayName = ' + psSingleQuote(o.displayName));
    lines.push('$Alias = ' + psSingleQuote(o.alias));
    lines.push('$PrimarySmtpAddress = ' + psSingleQuote(o.primarySmtp));
    lines.push('$Description = ' + psHereString(o.description));
    lines.push('$HideFromAddressLists = $' + (o.hideFromAddressLists ? 'true' : 'false'));
    lines.push(
        '$Delegates = @(' +
            (o.delegates.length ? o.delegates.map(psSingleQuote).join(', ') : '') +
            ')'
    );
    lines.push('$SuggestedAdminUpn = ' + psSingleQuote(o.suggestedAdminUpn));
    lines.push('$AutoMapping = $' + (o.autoMapping ? 'true' : 'false'));
    lines.push('$WantFullAccess = $' + (o.rights.includes('FullAccess') ? 'true' : 'false'));
    lines.push('$WantSendAs = $' + (o.rights.includes('SendAs') ? 'true' : 'false'));
    lines.push('$WantSendOnBehalf = $' + (o.rights.includes('SendOnBehalf') ? 'true' : 'false'));
    lines.push('');
    lines.push('Write-Host ""');
    lines.push('Write-Host "============================================" -ForegroundColor Cyan');
    lines.push('Write-Host "  Freigegebenes Postfach anlegen" -ForegroundColor Cyan');
    lines.push('Write-Host "  MS365-Schul-Tools · Organigramm" -ForegroundColor DarkCyan');
    lines.push('Write-Host "============================================" -ForegroundColor Cyan');
    lines.push('Write-Host ""');
    lines.push('Write-Host "Geplant:" -ForegroundColor White');
    lines.push('Write-Host ("  Anzeigename : " + $DisplayName)');
    lines.push('Write-Host ("  Alias       : " + $Alias)');
    lines.push(
        'Write-Host ("  Primäre Mail: " + $(if ([string]::IsNullOrWhiteSpace($PrimarySmtpAddress)) { "(wird von Exchange vergeben)" } else { $PrimarySmtpAddress }))'
    );
    lines.push('if ($Delegates.Count -gt 0) {');
    lines.push('  Write-Host ("  Berechtigte : " + ($Delegates -join ", "))');
    lines.push(
        '  Write-Host ("  Rechte      : " + (@($(if ($WantFullAccess) { "FullAccess" }), $(if ($WantSendAs) { "SendAs" }), $(if ($WantSendOnBehalf) { "SendOnBehalf" })) | Where-Object { $_ }) -join ", "))'
    );
    lines.push('} else {');
    lines.push('  Write-Host "  Berechtigte : (keine – nur Postfach anlegen)"');
    lines.push('}');
    lines.push('Write-Host ""');
    lines.push(
        'Write-Host "Wichtig: Mit einem Exchange-Admin der RICHTIGEN Schule / des richtigen Tenants anmelden." -ForegroundColor Yellow'
    );
    lines.push('');
    lines.push('Write-Step "1/5  Admin-Konto wählen"');
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
    lines.push('Write-Step "2/5  Bestätigung"');
    lines.push('$confirm = Read-Host "  Postfach jetzt in diesem Tenant anlegen? (J/N)"');
    lines.push('if ($confirm -notmatch "^[jJyY]") {');
    lines.push('  Write-WarnLine "Abgebrochen – es wurde nichts angelegt."');
    lines.push('  Pause-End');
    lines.push('  return');
    lines.push('}');
    lines.push('');
    lines.push('Write-Step "3/5  Exchange-Online-Modul"');
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
    lines.push('Write-Step "4/5  Anmelden bei Exchange Online"');
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
    lines.push('try {');
    lines.push('  if ([string]::IsNullOrWhiteSpace($DisplayName)) { throw "Anzeigename fehlt." }');
    lines.push('  if ([string]::IsNullOrWhiteSpace($Alias)) { throw "Alias fehlt." }');
    lines.push('');
    lines.push('  Write-Step "5/5  Postfach anlegen"');
    lines.push('  $newMailboxParams = @{');
    lines.push('    Shared = $true');
    lines.push('    Name = $DisplayName');
    lines.push('    DisplayName = $DisplayName');
    lines.push('    Alias = $Alias');
    lines.push('  }');
    lines.push('  if (-not [string]::IsNullOrWhiteSpace($PrimarySmtpAddress)) {');
    lines.push('    $newMailboxParams.PrimarySmtpAddress = $PrimarySmtpAddress');
    lines.push('  }');
    lines.push('  Write-Host ("  Erstelle: " + $DisplayName) -ForegroundColor DarkCyan');
    lines.push('  New-Mailbox @newMailboxParams | Out-Null');
    lines.push(
        '  $MailboxIdentity = if ([string]::IsNullOrWhiteSpace($PrimarySmtpAddress)) { $Alias } else { $PrimarySmtpAddress }'
    );
    lines.push('');
    lines.push('  Write-Host "  Warte auf Verfügbarkeit in Exchange Online …" -ForegroundColor DarkCyan');
    lines.push('  $mailbox = $null');
    lines.push('  for ($i = 1; $i -le 24; $i++) {');
    lines.push('    try {');
    lines.push('      $mailbox = Get-Mailbox -Identity $MailboxIdentity -ErrorAction Stop');
    lines.push('      break');
    lines.push('    } catch {');
    lines.push('      Start-Sleep -Seconds 5');
    lines.push('    }');
    lines.push('  }');
    lines.push(
        '  if (-not $mailbox) { throw "Postfach wurde erstellt, ist aber noch nicht abrufbar. Später mit Get-Mailbox prüfen und Berechtigungen nachziehen." }'
    );
    lines.push('  Write-Ok ("Postfach bereit: " + $mailbox.PrimarySmtpAddress)');
    lines.push('');
    lines.push('  if (-not [string]::IsNullOrWhiteSpace($Description)) {');
    lines.push('    Set-Mailbox -Identity $MailboxIdentity -Notes $Description');
    lines.push('  }');
    lines.push('  if ($HideFromAddressLists) {');
    lines.push('    Set-Mailbox -Identity $MailboxIdentity -HiddenFromAddressListsEnabled $true');
    lines.push('  }');
    lines.push('');

    if (!o.rights.length) {
        lines.push('  Write-WarnLine "Keine Rechte ausgewählt – es werden keine Personen berechtigt."');
    } else {
        lines.push('  if ($Delegates.Count -eq 0) {');
        lines.push('    Write-WarnLine "Keine Personen aus dem Organigramm – Rechte übersprungen."');
        lines.push('  } else {');
        lines.push(
            '    Write-Host ("  Setze Berechtigungen für " + $Delegates.Count + " Person(en) …") -ForegroundColor DarkCyan'
        );
        if (o.rights.includes('FullAccess')) {
            lines.push('    if ($WantFullAccess) {');
            lines.push('      foreach ($d in $Delegates) {');
            lines.push(
                '        Add-MailboxPermission -Identity $MailboxIdentity -User $d -AccessRights FullAccess -InheritanceType All -AutoMapping:$AutoMapping -ErrorAction Continue | Out-Null'
            );
            lines.push('        Write-Ok ("FullAccess → " + $d)');
            lines.push('      }');
            lines.push('    }');
        }
        if (o.rights.includes('SendAs')) {
            lines.push('    if ($WantSendAs) {');
            lines.push('      foreach ($d in $Delegates) {');
            lines.push(
                '        Add-RecipientPermission -Identity $MailboxIdentity -Trustee $d -AccessRights SendAs -Confirm:$false -ErrorAction Continue | Out-Null'
            );
            lines.push('        Write-Ok ("SendAs → " + $d)');
            lines.push('      }');
            lines.push('    }');
        }
        if (o.rights.includes('SendOnBehalf')) {
            lines.push('    if ($WantSendOnBehalf) {');
            lines.push(
                '      Set-Mailbox -Identity $MailboxIdentity -GrantSendOnBehalfTo @{Add=$Delegates} -ErrorAction Continue'
            );
            lines.push('      Write-Ok "SendOnBehalf gesetzt."');
            lines.push('    }');
        }
        lines.push('  }');
    }

    lines.push('');
    lines.push('  Write-Host ""');
    lines.push('  Write-Host "== Ergebnis ==" -ForegroundColor Cyan');
    lines.push(
        '  Get-Mailbox -Identity $MailboxIdentity | Select-Object DisplayName, PrimarySmtpAddress, Alias, RecipientTypeDetails, HiddenFromAddressListsEnabled | Format-List'
    );
    lines.push('  Write-Host "FullAccess:" -ForegroundColor DarkCyan');
    lines.push(
        '  Get-MailboxPermission -Identity $MailboxIdentity | Where-Object { $_.IsInherited -eq $false -and $_.User -notlike "NT AUTHORITY*" -and $_.User -notlike "S-1-5-*" } | Select-Object User,AccessRights,Deny,IsInherited | Format-Table -AutoSize'
    );
    lines.push('  Write-Host "SendAs:" -ForegroundColor DarkCyan');
    lines.push(
        '  Get-RecipientPermission -Identity $MailboxIdentity | Where-Object { $_.Trustee -and $_.AccessRights -contains "SendAs" } | Select-Object Trustee,AccessRights,IsInherited | Format-Table -AutoSize'
    );
    lines.push('  Write-Host ""');
    lines.push(
        '  Write-Host "Fertig. Das Postfach ist im Organigramm bereits vorgemerkt – Dialog dort ggf. speichern." -ForegroundColor Green'
    );
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
 * Doppelklickbare .cmd-Datei: startet PowerShell, liest den Assistenten aus derselben Datei.
 * @param {Parameters<typeof resolveMailboxOptions>[0]} options
 */
export function buildSharedMailboxCreateCmd(options) {
    const ps = buildSharedMailboxCreateScript(options);
    const withEnv =
        '@echo off\r\n' +
        'setlocal EnableExtensions\r\n' +
        'chcp 65001 >nul\r\n' +
        'title Freigegebenes Postfach anlegen\r\n' +
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
        SHARED_MAILBOX_CMD_PS_MARKER +
        "';" +
        ' $i=0; while($i -lt $lines.Count -and $lines[$i].Trim() -ne $marker){ $i++ };' +
        ' if($i -ge $lines.Count){ throw \'PowerShell-Teil in .cmd nicht gefunden.\' };' +
        ' $script=($lines[($i+1)..($lines.Count-1)] -join [Environment]::NewLine);' +
        " $ps1=Join-Path $env:TEMP ('ms365-shared-mailbox-' + [guid]::NewGuid().ToString('N') + '.ps1');" +
        ' Set-Content -LiteralPath $ps1 -Value $script -Encoding UTF8;' +
        ' try { & powershell.exe -NoProfile -ExecutionPolicy Bypass -NoLogo -File $ps1; exit $LASTEXITCODE }' +
        ' finally { Remove-Item -LiteralPath $ps1 -Force -ErrorAction SilentlyContinue }"\r\n' +
        'set "EC=%ERRORLEVEL%"\r\n' +
        'endlocal & exit /b %EC%\r\n' +
        '\r\n' +
        SHARED_MAILBOX_CMD_PS_MARKER +
        '\r\n';

    return withEnv + ps.replace(/\r?\n/g, '\r\n') + '\r\n';
}

/**
 * Vorschlag für Alias/SMTP aus Anzeigename + Schuildomain.
 * @param {string} displayName
 * @param {string} [domain]
 */
export function suggestSharedMailboxIdentity(displayName, domain) {
    const nick = String(displayName || '')
        .toLowerCase()
        .replace(/ä/g, 'ae')
        .replace(/ö/g, 'oe')
        .replace(/ü/g, 'ue')
        .replace(/ß/g, 'ss')
        .replace(/[^a-z0-9._-]+/g, '-')
        .replace(/-+/g, '-')
        .replace(/^[-._]+|[-._]+$/g, '')
        .slice(0, 60);
    const alias = nick || 'postfach';
    const dom = String(domain || '')
        .trim()
        .replace(/^@/, '')
        .toLowerCase();
    return {
        alias: alias,
        primarySmtp: dom ? alias + '@' + dom : ''
    };
}
