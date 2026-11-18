# Deploy Schüler-Stammdaten-Portal-API nach Azure Functions (Flex Consumption).
#
# Beispiel:
#   cd backend\stammdaten-portal-api
#   .\scripts\deploy-azure.ps1 -ResourceGroup "rg-ms365schule-kursteams-dev" `
#       -FunctionAppName "func-ms365-ssp-dev" -CreateResources
#
# Live (Stand): https://func-ms365-ssp-dev.azurewebsites.net/api/health
#
# SharePoint danach per App Settings:
#   STAMMDATEN_STORE=sharepoint
#   STAMMDATEN_GRAPH_TENANT_ID / CLIENT_ID / CLIENT_SECRET
#   STAMMDATEN_SP_SITE_URL=https://….sharepoint.com/sites/…
# Dann: POST /api/admin/ensure-list (Function Key)

param(
    [Parameter(Mandatory = $true)]
    [string] $ResourceGroup,

    [Parameter(Mandatory = $true)]
    [string] $FunctionAppName,

    [string] $Location = "westeurope",

    [string] $StorageAccountName = "",

    [string] $CorsExtra = "https://app.ms365.schule",

    [switch] $CreateResources,

    [switch] $SkipPublish
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw "Azure CLI (az) nicht gefunden. Bitte installieren und 'az login' ausführen."
}

function New-RandomAlnum([int]$len) {
    -join ((48..57 + 97..122) | Get-Random -Count $len | ForEach-Object { [char]$_ })
}

function Get-OrCreateSecret([string]$name, [int]$len = 48) {
    $existing = az functionapp config appsettings list `
        --name $FunctionAppName `
        --resource-group $ResourceGroup `
        --query "[?name=='$name'].value | [0]" -o tsv 2>$null
    if ($existing) {
        Write-Host "$name bleibt unverändert."
        return $existing
    }
    Write-Host "$name neu erzeugen …"
    return (-join ((48..57 + 65..90 + 97..122) | Get-Random -Count $len | ForEach-Object { [char]$_ }))
}

if ($CreateResources) {
    if (-not $StorageAccountName) {
        $StorageAccountName = "stms365stam" + (New-RandomAlnum 6)
    }
    Write-Host "Resource Group: $ResourceGroup ($Location)"
    az group create --name $ResourceGroup --location $Location | Out-Null

    $prevEap = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    $exists = az storage account show --name $StorageAccountName --resource-group $ResourceGroup -o tsv --query name 2>$null
    if ($LASTEXITCODE -ne 0) { $exists = $null }
    $ErrorActionPreference = $prevEap

    if (-not $exists) {
        Write-Host "Storage Account: $StorageAccountName"
        az storage account create `
            --name $StorageAccountName `
            --resource-group $ResourceGroup `
            --location $Location `
            --sku Standard_LRS `
            --kind StorageV2 | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Storage Account konnte nicht erstellt werden." }
    } else {
        Write-Host "Storage Account existiert: $StorageAccountName"
    }

    $ErrorActionPreference = "Continue"
    $fa = az functionapp show --name $FunctionAppName --resource-group $ResourceGroup -o tsv --query name 2>$null
    if ($LASTEXITCODE -ne 0) { $fa = $null }
    $ErrorActionPreference = $prevEap

    if (-not $fa) {
        Write-Host "Function App (Flex Consumption): $FunctionAppName"
        az functionapp create `
            --resource-group $ResourceGroup `
            --name $FunctionAppName `
            --storage-account $StorageAccountName `
            --flexconsumption-location $Location `
            --runtime node `
            --runtime-version 22 `
            --os-type Linux | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Function App konnte nicht erstellt werden." }
    } else {
        Write-Host "Function App existiert: $FunctionAppName"
    }
}

$pepper = Get-OrCreateSecret "STAMMDATEN_PIN_PEPPER"
$session = Get-OrCreateSecret "STAMMDATEN_SESSION_SECRET"
$adminKey = Get-OrCreateSecret "STAMMDATEN_ADMIN_KEY" 40

$existingStore = az functionapp config appsettings list `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --query "[?name=='STAMMDATEN_STORE'].value | [0]" -o tsv 2>$null
if (-not $existingStore) { $existingStore = "memory" }

Write-Host "App Settings setzen …"
# Flex Consumption: FUNCTIONS_WORKER_RUNTIME nicht setzen (wird vom Runtime-Stack verwaltet)
# CORS=* im App-Code; Azure-Host-CORS muss die Browser-Origins ebenfalls erlauben.
az functionapp config appsettings set `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --settings `
        "STAMMDATEN_PIN_PEPPER=$pepper" `
        "STAMMDATEN_SESSION_SECRET=$session" `
        "STAMMDATEN_ADMIN_KEY=$adminKey" `
        "STAMMDATEN_CORS_ORIGINS=*" `
        "STAMMDATEN_STORE=$existingStore" `
        "STAMMDATEN_SP_LIST_NAME=SchuelerAnmeldungen" | Out-Null

if (-not $SkipPublish) {
    Write-Host "npm install …"
    if (Test-Path "package-lock.json") {
        npm ci --omit=dev
    } else {
        npm install --omit=dev
    }

    Write-Host "Publish → $FunctionAppName …"
    if (-not (Get-Command func -ErrorAction SilentlyContinue)) {
        throw "Azure Functions Core Tools (func) erforderlich."
    }
    func azure functionapp publish $FunctionAppName --javascript
}

# Host-CORS: Azure '*' ist KEIN Wildcard. Unbekannte Origin → OPTIONS ohne ACAO → Failed to fetch.
# Explizite lokale Vite-Ports + Prod. Zusätzlich Vite-Proxy /__ssp-gw für localhost.
Write-Host "Host-CORS setzen (Vite-Ports + Prod) …"
try {
    $siteId = az functionapp show --name $FunctionAppName --resource-group $ResourceGroup --query id -o tsv
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($p in 5173..5200) {
        $list.Add("http://localhost:$p")
        $list.Add("http://127.0.0.1:$p")
    }
    @(
        "null",
        "http://localhost:5500",
        "http://127.0.0.1:5500",
        "http://localhost:3000",
        "http://127.0.0.1:3000",
        "https://app.ms365.schule",
        "https://ms365.schule",
        "https://kurtrocks.github.io",
        "https://kurtsoeser.github.io"
    ) + @($CorsExtra -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) | ForEach-Object { $list.Add($_) }
    $uniq = $list | Select-Object -Unique
    $jsonOrigins = ($uniq | ForEach-Object { '"' + ($_ -replace '\\', '\\\\' -replace '"', '\"') + '"' }) -join ','
    $corsBody = '{"properties":{"cors":{"allowedOrigins":[' + $jsonOrigins + '],"supportCredentials":false}}}'
    $tmp = [System.IO.Path]::GetTempFileName()
    [System.IO.File]::WriteAllText($tmp, $corsBody)
    az rest --method PUT --uri "$siteId/config/web?api-version=2024-04-01" --body "@$tmp" | Out-Null
    Remove-Item $tmp -Force -ErrorAction SilentlyContinue
} catch {
    Write-Host "Host-CORS-Set übersprungen: $_"
}

$base = "https://$FunctionAppName.azurewebsites.net/api"
Write-Host ""
Write-Host "Fertig."
Write-Host "  Health:  $base/health"
Write-Host "  Admin:   $base/ssp-ensure-list  (Header x-admin-key)"
Write-Host ""
Write-Host "Nächste Schritte SharePoint:"
Write-Host "  1) Entra-App mit Application Permission Sites.Selected (oder Sites.ReadWrite.All)"
Write-Host "  2) App Settings: STAMMDATEN_STORE=sharepoint + GRAPH_* + STAMMDATEN_SP_SITE_URL"
Write-Host "  3) POST $base/ssp-ensure-list mit x-admin-key = STAMMDATEN_ADMIN_KEY"
Write-Host "  4) Im Admin-Tool Modus=remote, API-Base=$base, Function Key = STAMMDATEN_ADMIN_KEY"
