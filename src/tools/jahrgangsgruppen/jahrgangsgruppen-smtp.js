/**
 * Exchange-SMTP PowerShell-Export für Klassengruppen (Analyse 02 Phase B).
 */

export function psEscapeSingle(s) {
    return String(s || '').replace(/'/g, "''");
}

/**
 * @param {Array<{ id: string, name: string, smtp: string }>} items
 * @param {string} domain
 */
export function buildClassSmtpPs1(items, domain) {
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
    (items || []).forEach(function (it, idx) {
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
