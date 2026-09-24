<#
.SYNOPSIS
  Schreibt Demo-Daten Projektwochen auf eine SharePoint-Site (Microsoft Graph).

.DESCRIPTION
  Voraussetzung: Listen PW-Aktionen und PW-Angebote existieren.
  Anmeldung: Connect-MgGraph -Scopes "Sites.ReadWrite.All","User.Read"

.EXAMPLE
  pwsh -File scripts/seed-projektwochen-demo.ps1
  pwsh -File scripts/seed-projektwochen-demo.ps1 -SiteUrl "https://kurtrocks.sharepoint.com/sites/MS365-Schultools"
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
    $JsonPath = Join-Path (Get-RepoRoot) 'docs/demo-data/projektwochen-demo.json'
}
if (-not (Test-Path -LiteralPath $JsonPath)) {
    throw "JSON fehlt: $JsonPath – zuerst: node scripts/generate-projektwochen-demo-json.mjs"
}

$pack = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
Write-Host "Site: $SiteUrl"
Write-Host "Angebote: $($pack.counts.angebote) · freigegeben: $($pack.counts.freigegeben)"

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
    if (-not $list) { throw "Liste fehlt: $displayName – zuerst Listen-Paket anlegen." }
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
    $clean = @{}
    foreach ($k in $Fields.Keys) {
        if ($null -ne $Fields[$k] -and "$($Fields[$k])" -ne '') { $clean[$k] = $Fields[$k] }
    }
    if ($match) {
        Invoke-MgGraphRequest -Method PATCH -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$ListId/items/$($match.id)/fields" -Body ($clean | ConvertTo-Json -Depth 6) -ContentType 'application/json'
        return 'updated'
    }
    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/sites/$siteId/lists/$ListId/items" -Body (@{ fields = $clean } | ConvertTo-Json -Depth 6) -ContentType 'application/json'
    return 'created'
}

function ConvertTo-Hashtable($obj) {
    $h = @{}
    if (-not $obj) { return $h }
    $obj.PSObject.Properties | ForEach-Object {
        if ($null -ne $_.Value -and "$($_.Value)" -ne '') { $h[$_.Name] = $_.Value }
    }
    return $h
}

$listAktionen = Get-ListId 'PW-Aktionen'
$listAngebote = Get-ListId 'PW-Angebote'
$existingAktionen = @(Get-AllListItems $listAktionen)
$existingAngebote = @(Get-AllListItems $listAngebote)

$aktionFields = ConvertTo-Hashtable $pack.aktion
Write-Host "Aktion $($aktionFields.AktionId) …"
Upsert-ByField -ListId $listAktionen -KeyField 'AktionId' -KeyValue $aktionFields.AktionId -Fields $aktionFields -ExistingItems $existingAktionen | Out-Null

$i = 0
foreach ($row in @($pack.angebote)) {
    $i++
    $fields = ConvertTo-Hashtable $row
    $r = Upsert-ByField -ListId $listAngebote -KeyField 'AngebotId' -KeyValue $fields.AngebotId -Fields $fields -ExistingItems $existingAngebote
    if ($i % 5 -eq 0) { Write-Host "  … $i / $($pack.angebote.Count) ($r)" }
}

Write-Host "Fertig. Seed $($pack.seedTag) auf $SiteUrl"
