# Fügt Scope Election.Access zur License-Backend-App hinzu und pre-authorisiert die SPA.
# Idempotent – zwei Graph-PATCHes (Scope zuerst, dann PreAuth).
#
#   cd backend\election-api
#   .\scripts\ensure-entra-election-scope.ps1

param(
    [string] $ApiAppId = "12e0cfe2-8337-4b35-93e8-542faf658eb3",
    [string] $SpaAppId = "e1d877c3-004c-4040-8c3b-81a59e0c7050",
    [string] $ScopeValue = "Election.Access",
    [string] $ScopeId = "c4e8b1a2-7d3f-4a9e-b6c1-8f2e0d5a9147"
)

$ErrorActionPreference = "Stop"

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw "Azure CLI (az) nicht gefunden."
}

function Write-JsonTemp($obj) {
    $tmp = Join-Path $env:TEMP ("election-entra-" + [guid]::NewGuid().ToString() + ".json")
    $json = $obj | ConvertTo-Json -Depth 12
    [System.IO.File]::WriteAllText($tmp, $json, [System.Text.UTF8Encoding]::new($false))
    return $tmp
}

function Invoke-GraphPatch([string]$objectId, $bodyObj) {
    $tmp = Write-JsonTemp $bodyObj
    try {
        az rest --method PATCH `
            --uri "https://graph.microsoft.com/v1.0/applications/$objectId" `
            --headers "Content-Type=application/json" `
            --body "@$tmp" | Out-Null
    } finally {
        Remove-Item $tmp -ErrorAction SilentlyContinue
    }
}

Write-Host "API-App laden: $ApiAppId"
$app = az ad app show --id $ApiAppId -o json | ConvertFrom-Json
$objectId = [string]$app.id

$scopes = @()
if ($app.api -and $app.api.oauth2PermissionScopes) {
    $scopes = @($app.api.oauth2PermissionScopes)
}

$existing = $scopes | Where-Object { $_.value -eq $ScopeValue }
if ($existing) {
    $ScopeId = [string]$existing.id
    Write-Host "Scope $ScopeValue existiert bereits ($ScopeId)."
} else {
    Write-Host "Schritt 1/2: Scope $ScopeValue anlegen ($ScopeId) …"
    $scopeHashes = @()
    foreach ($s in $scopes) {
        $scopeHashes += @{
            adminConsentDescription = [string]$s.adminConsentDescription
            adminConsentDisplayName = [string]$s.adminConsentDisplayName
            id                      = [string]$s.id
            isEnabled               = [bool]$s.isEnabled
            type                    = [string]$s.type
            userConsentDescription  = [string]$s.userConsentDescription
            userConsentDisplayName  = [string]$s.userConsentDisplayName
            value                   = [string]$s.value
        }
    }
    $scopeHashes += @{
        adminConsentDescription = "Anonyme Schulsprecherwahl: Meta, Stimmenabgabe und Auszaehlung."
        adminConsentDisplayName = "Schulsprecherwahl-API"
        id                      = $ScopeId
        isEnabled               = $true
        type                    = "User"
        userConsentDescription  = "Ermoeglicht der App, deine Stimme anonym abzugeben und Wahlergebnisse zu laden."
        userConsentDisplayName  = "Schulsprecherwahl nutzen"
        value                   = $ScopeValue
    }
    Invoke-GraphPatch $objectId @{
        api = @{
            oauth2PermissionScopes      = $scopeHashes
            requestedAccessTokenVersion = 2
        }
    }
    Start-Sleep -Seconds 3
    $app = az ad app show --id $ApiAppId -o json | ConvertFrom-Json
    $scopes = @($app.api.oauth2PermissionScopes)
}

$licenseScopeId = ($scopes | Where-Object { $_.value -eq "License.Access" } | Select-Object -First 1).id
$electionScopeId = ($scopes | Where-Object { $_.value -eq $ScopeValue } | Select-Object -First 1).id
if (-not $electionScopeId) { throw "Scope $ScopeValue nach Anlegen nicht sichtbar." }
$ScopeId = [string]$electionScopeId

$wantedIds = @($ScopeId)
if ($licenseScopeId) { $wantedIds += [string]$licenseScopeId }

Write-Host "Schritt 2/2: SPA $SpaAppId pre-authorisieren …"
$preAuth = @()
if ($app.api -and $app.api.preAuthorizedApplications) {
    $preAuth = @($app.api.preAuthorizedApplications)
}
$spaEntry = $preAuth | Where-Object { $_.appId -eq $SpaAppId }
if (-not $spaEntry) {
    $ids = $wantedIds
} else {
    $ids = @($spaEntry.delegatedPermissionIds | ForEach-Object { [string]$_ })
    foreach ($wid in $wantedIds) {
        if ($ids -notcontains $wid) { $ids += $wid }
    }
}

$preHashes = @()
foreach ($p in $preAuth) {
    if ($p.appId -eq $SpaAppId) { continue }
    $preHashes += @{
        appId                  = [string]$p.appId
        delegatedPermissionIds = @($p.delegatedPermissionIds | ForEach-Object { [string]$_ })
    }
}
$preHashes += @{
    appId                  = $SpaAppId
    delegatedPermissionIds = @($ids)
}

# Scopes unverändert mitschicken (Graph ersetzt sonst oft die Liste)
$scopeHashes = @()
foreach ($s in $scopes) {
    $scopeHashes += @{
        adminConsentDescription = [string]$s.adminConsentDescription
        adminConsentDisplayName = [string]$s.adminConsentDisplayName
        id                      = [string]$s.id
        isEnabled               = [bool]$s.isEnabled
        type                    = [string]$s.type
        userConsentDescription  = [string]$s.userConsentDescription
        userConsentDisplayName  = [string]$s.userConsentDisplayName
        value                   = [string]$s.value
    }
}

Invoke-GraphPatch $objectId @{
    api = @{
        oauth2PermissionScopes      = $scopeHashes
        preAuthorizedApplications   = $preHashes
        requestedAccessTokenVersion = 2
    }
}

Write-Host "SPA requiredResourceAccess aktualisieren …"
$spa = az ad app show --id $SpaAppId -o json | ConvertFrom-Json
$spaObjectId = [string]$spa.id
$rra = @()
if ($spa.requiredResourceAccess) { $rra = @($spa.requiredResourceAccess) }

$other = @($rra | Where-Object { $_.resourceAppId -ne $ApiAppId })
$apiBlock = $rra | Where-Object { $_.resourceAppId -eq $ApiAppId }
$raIds = @()
if ($apiBlock) {
    $raIds = @($apiBlock.resourceAccess | ForEach-Object { [string]$_.id })
}
foreach ($wid in $wantedIds) {
    if ($raIds -notcontains $wid) { $raIds += $wid }
}

$rraHashes = @()
foreach ($block in $other) {
    $access = @()
    foreach ($a in @($block.resourceAccess)) {
        $access += @{ id = [string]$a.id; type = [string]$a.type }
    }
    $rraHashes += @{
        resourceAppId  = [string]$block.resourceAppId
        resourceAccess = $access
    }
}
$rraHashes += @{
    resourceAppId  = $ApiAppId
    resourceAccess = @($raIds | ForEach-Object { @{ id = $_; type = "Scope" } })
}

Invoke-GraphPatch $spaObjectId @{ requiredResourceAccess = $rraHashes }

Write-Host ""
Write-Host "Fertig."
Write-Host "  Scope: api://$ApiAppId/$ScopeValue"
Write-Host "  Scope-Id: $ScopeId"
Write-Host "  SPA pre-authorized: $SpaAppId"
Write-Host ""
Write-Host "Hinweis: Nutzer:innen ggf. einmal neu anmelden / Consent fuer Election.Access."
