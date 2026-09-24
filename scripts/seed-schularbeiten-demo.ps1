<#
.SYNOPSIS
  Schreibt Demo-Daten Schularbeiten 2026/27 auf eine SharePoint-Site (Microsoft Graph).

.DESCRIPTION
  Voraussetzung: Listen-Paket (Regelwerk, Terminfenster, Schularbeiten, SA-FachMeta) existiert.
  Anmeldung: Connect-MgGraph -Scopes "Sites.ReadWrite.All","User.Read"

.EXAMPLE
  pwsh -File scripts/seed-schularbeiten-demo.ps1
  pwsh -File scripts/seed-schularbeiten-demo.ps1 -SiteUrl "https://kurtrocks.sharepoint.com/sites/MS365-Schultools"
#>
param(
    [string]$SiteUrl = 'https://kurtrocks.sharepoint.com/sites/MS365-Schultools',
    [string]$JsonPath = ''
)

$ErrorActionPreference = 'Stop'

function Get-RepoRoot {
    Split-Path -Parent $PSScriptRoot
}

if (-not $JsonPath) {
    $JsonPath = Join-Path (Get-RepoRoot) 'docs/demo-data/schularbeiten-2026-27.json'
}
if (-not (Test-Path -LiteralPath $JsonPath)) {
    throw "JSON fehlt: $JsonPath – zuerst: node scripts/generate-schularbeiten-demo-json.mjs"
}

$pack = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Host "Site: $SiteUrl"
Write-Host "Schularbeiten: $($pack.counts.schularbeiten) · Fenster: $($pack.counts.terminfenster) · FachMeta: $($pack.counts.fachMeta)"

$ctx = Get-MgContext
if (-not $ctx) {
    Write-Host 'Melde an (Browser) …'
    Connect-MgGraph -Scopes 'Sites.ReadWrite.All','User.Read' -NoWelcome
}

$uri = [Uri]$SiteUrl
$hostName = $uri.Host
$path = $uri.AbsolutePath.Trim('/')
$sitePath = "$hostName`:/$path"
Write-Host "Resolve site: $sitePath"
$site = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$sitePath"
$siteId = $site.id
Write-Host "SiteId: $siteId"

function Get-ListId([string]$displayName) {
    $filter = [uri]::EscapeDataString("displayName eq '$displayName'")
    $res = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists?`$filter=$filter&`$select=id,displayName"
    $list = @($res.value) | Select-Object -First 1
    if (-not $list) { throw "Liste fehlt: $displayName – zuerst Paket anlegen." }
    return $list.id
}

function Get-AllListItems([string]$listId) {
    $items = @()
    $url = "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$listId/items?`$expand=fields&`$top=100"
    while ($url) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $url
        $items += @($page.value)
        $url = $page.'@odata.nextLink'
    }
    return $items
}

function Upsert-ByField {
    param(
        [string]$ListId,
        [string]$KeyField,
        [string]$KeyValue,
        [hashtable]$Fields,
        [object[]]$ExistingItems
    )
    $match = $ExistingItems | Where-Object { $_.fields.$KeyField -eq $KeyValue } | Select-Object -First 1
    $body = @{ fields = $Fields }
    if ($match) {
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$ListId/items/$($match.id)/fields" -Body ($Fields | ConvertTo-Json -Depth 6) -ContentType 'application/json'
        return 'updated'
    }
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$ListId/items" -Body ($body | ConvertTo-Json -Depth 6) -ContentType 'application/json'
    return 'created'
}

function Convert-ToHashtable($obj) {
    $h = @{}
    foreach ($p in $obj.PSObject.Properties) {
        if ($null -ne $p.Value -and $p.Value -ne '') { $h[$p.Name] = $p.Value }
    }
    return $h
}

# Regelwerk
$rwList = Get-ListId 'Regelwerk'
$rwItems = Get-AllListItems $rwList
$rwFields = Convert-ToHashtable $pack.regelwerk
$rwResult = Upsert-ByField -ListId $rwList -KeyField 'RegelwerkId' -KeyValue $pack.regelwerk.RegelwerkId -Fields $rwFields -ExistingItems $rwItems
Write-Host "Regelwerk: $rwResult"

# FachMeta
$fmList = Get-ListId 'SA-FachMeta'
$fmItems = Get-AllListItems $fmList
$fmC = 0; $fmU = 0
foreach ($row in $pack.fachMeta) {
    $f = Convert-ToHashtable $row
    $r = Upsert-ByField -ListId $fmList -KeyField 'FachCode' -KeyValue $row.FachCode -Fields $f -ExistingItems $fmItems
    if ($r -eq 'created') { $fmC++ } else { $fmU++; $fmItems = Get-AllListItems $fmList }
    Start-Sleep -Milliseconds 60
}
Write-Host "SA-FachMeta: +$fmC ~$fmU"

# Terminfenster
$tfList = Get-ListId 'Terminfenster'
$tfItems = Get-AllListItems $tfList
$tfC = 0; $tfU = 0
foreach ($row in $pack.terminfenster) {
    $f = Convert-ToHashtable $row
    $r = Upsert-ByField -ListId $tfList -KeyField 'TerminfensterId' -KeyValue $row.TerminfensterId -Fields $f -ExistingItems $tfItems
    if ($r -eq 'created') { $tfC++; $tfItems = @(Get-AllListItems $tfList) } else { $tfU++ }
    Start-Sleep -Milliseconds 60
}
Write-Host "Terminfenster: +$tfC ~$tfU"

# Schularbeiten
$saList = Get-ListId 'Schularbeiten'
$saItems = Get-AllListItems $saList
$saC = 0; $saU = 0
$n = 0
foreach ($row in $pack.schularbeiten) {
    $n++
    $f = Convert-ToHashtable $row
    $r = Upsert-ByField -ListId $saList -KeyField 'SchularbeitId' -KeyValue $row.SchularbeitId -Fields $f -ExistingItems $saItems
    if ($r -eq 'created') {
        $saC++
        if ($saC % 5 -eq 0) { $saItems = @(Get-AllListItems $saList) }
    } else { $saU++ }
    if ($n % 10 -eq 0) { Write-Host "  … $n/$($pack.counts.schularbeiten)" }
    Start-Sleep -Milliseconds 60
}
Write-Host "Schularbeiten: +$saC ~$saU"
Write-Host 'Fertig.'
Write-Host 'Stammdaten-Codes: siehe JSON .stammdaten (tenant.html).'
