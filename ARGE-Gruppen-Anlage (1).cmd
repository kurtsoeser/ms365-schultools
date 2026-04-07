@echo off
chcp 65001 >nul
title ARGE-Gruppen-Anlage
cd /d "%~dp0"
echo.
echo Starte ARGE-Gruppen-Anlage Microsoft Graph ...
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -EncodedCommand cABhAHIAYQBtACgAWwBzAHQAcgBpAG4AZwBdACQAUABhAHQAaAApAA0ACgAkAHQAZQB4AHQAIAA9ACAAWwBTAHkAcwB0AGUAbQAuAEkATwAuAEYAaQBsAGUAXQA6ADoAUgBlAGEAZABBAGwAbABUAGUAeAB0ACgAJABQAGEAdABoACkADQAKACQAbQBhAHIAawBlAHIAIAA9ACAAJwA6ADoATQBTADMANgA1AF8AUABTAF8AQgBFAEcASQBOADoAOgAnAA0ACgAkAGkAZAB4ACAAPQAgACQAdABlAHgAdAAuAEkAbgBkAGUAeABPAGYAKAAkAG0AYQByAGsAZQByACkADQAKAGkAZgAgACgAJABpAGQAeAAgAC0AbAB0ACAAMAApACAAewAgAFcAcgBpAHQAZQAtAEUAcgByAG8AcgAgACcATQBTADMANgA1ADoAIABNAGEAcgBrAGUAcgAgAG4AaQBjAGgAdAAgAGcAZQBmAHUAbgBkAGUAbgAuACcAOwAgAGUAeABpAHQAIAAxACAAfQANAAoAJABiAG8AZAB5ACAAPQAgACQAdABlAHgAdAAuAFMAdQBiAHMAdAByAGkAbgBnACgAJABpAGQAeAAgACsAIAAkAG0AYQByAGsAZQByAC4ATABlAG4AZwB0AGgAKQAuAFQAcgBpAG0AUwB0AGEAcgB0ACgAKQANAAoAJAB0AG0AcAAgAD0AIABKAG8AaQBuAC0AUABhAHQAaAAgACQAZQBuAHYAOgBUAEUATQBQACAAKAAnAG0AcwAzADYANQAtACcAIAArACAAWwBnAHUAaQBkAF0AOgA6AE4AZQB3AEcAdQBpAGQAKAApAC4AVABvAFMAdAByAGkAbgBnACgAKQAgACsAIAAnAC4AcABzADEAJwApAA0ACgAkAHUAdABmADgAIAA9ACAATgBlAHcALQBPAGIAagBlAGMAdAAgAFMAeQBzAHQAZQBtAC4AVABlAHgAdAAuAFUAVABGADgARQBuAGMAbwBkAGkAbgBnACAAJAB0AHIAdQBlAA0ACgBbAFMAeQBzAHQAZQBtAC4ASQBPAC4ARgBpAGwAZQBdADoAOgBXAHIAaQB0AGUAQQBsAGwAVABlAHgAdAAoACQAdABtAHAALAAgACQAYgBvAGQAeQAsACAAJAB1AHQAZgA4ACkADQAKAHQAcgB5ACAAewANAAoAIAAgACYAIABwAG8AdwBlAHIAcwBoAGUAbABsAC4AZQB4AGUAIAAtAE4AbwBQAHIAbwBmAGkAbABlACAALQBFAHgAZQBjAHUAdABpAG8AbgBQAG8AbABpAGMAeQAgAEIAeQBwAGEAcwBzACAALQBGAGkAbABlACAAJAB0AG0AcAANAAoAIAAgAGUAeABpAHQAIAAkAEwAQQBTAFQARQBYAEkAVABDAE8ARABFAA0ACgB9ACAAZgBpAG4AYQBsAGwAeQAgAHsADQAKACAAIABSAGUAbQBvAHYAZQAtAEkAdABlAG0AIAAtAEwAaQB0AGUAcgBhAGwAUABhAHQAaAAgACQAdABtAHAAIAAtAEYAbwByAGMAZQAgAC0ARQByAHIAbwByAEEAYwB0AGkAbwBuACAAUwBpAGwAZQBuAHQAbAB5AEMAbwBuAHQAaQBuAHUAZQANAAoAfQA= "%~f0"
set ERR=%ERRORLEVEL%
if not "%ERR%"=="0" (
  echo.
  echo Fehlercode: %ERR%
)
echo.
pause
exit /b
::MS365_PS_BEGIN::
#Requires -Version 5.1
# ARGE-Gruppen (Microsoft 365 Unified Groups, kein Kursteam)
# Erzeugt in der Browser-App am 2026-04-07T14:29:27.942Z
# Daten sind unten eingebettet.

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$ErrorActionPreference = "Continue"

Write-Host ""
Write-Host "========================================"  -ForegroundColor Cyan
Write-Host "  ARGE-Gruppen (Microsoft Graph)"       -ForegroundColor Cyan
Write-Host "========================================"  -ForegroundColor Cyan
Write-Host ""

if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Host "Installiere Microsoft.Graph (einmalig)..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Scope CurrentUser -Force
}
Import-Module Microsoft.Graph -ErrorAction Stop

Write-Host "Anmeldung bei Microsoft Graph (interaktiv, MFA moeglich)..." -ForegroundColor Yellow
Connect-MgGraph -Scopes "Group.ReadWrite.All","User.Read.All"

$rows = @(
    [PSCustomObject]@{ DisplayName = 'ARGE Deutch'; MailNickname = 'ag-deutch'; OwnerUpn = 'dave.grohl@kurtrocks.com'; Description = 'ARGE-Gruppe: ARGE Deutch' },
    [PSCustomObject]@{ DisplayName = 'ARGE Mathematik'; MailNickname = 'ag-mathematik'; OwnerUpn = 'dave.grohl@kurtrocks.com'; Description = 'ARGE-Gruppe: ARGE Mathematik' },
    [PSCustomObject]@{ DisplayName = 'ARGE Englisch'; MailNickname = 'ag-englisch'; OwnerUpn = 'dave.grohl@kurtrocks.com'; Description = 'ARGE-Gruppe: ARGE Englisch' }
)

$i = 0
foreach ($r in $rows) {
    $i++
    try {
        $owner = Get-MgUser -UserId $r.OwnerUpn -ErrorAction Stop
        $group = New-MgGroup `
            -DisplayName $r.DisplayName `
            -Description $r.Description `
            -MailNickname $r.MailNickname `
            -MailEnabled:$true `
            -SecurityEnabled:$false `
            -GroupTypes @("Unified") `
            -Visibility "Private" `
            -ErrorAction Stop
        New-MgGroupOwner -GroupId $group.Id -DirectoryObjectId $owner.Id
        Write-Host ("OK [{0}/{1}] {2} -> {3}" -f $i, $rows.Count, $r.DisplayName, $r.MailNickname) -ForegroundColor Green
    }
    catch {
        Write-Warning ("Fehler [{0}] {1}: {2}" -f $i, $r.DisplayName, $_.Exception.Message)
    }
    Start-Sleep -Seconds 2
}

# Gruppen-E-Mail: <MailNickname>@ms365.schule

Write-Host ""
Write-Host "Fertig." -ForegroundColor Cyan
Read-Host "Enter druecken zum Beenden"