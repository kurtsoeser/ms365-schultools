# Deploy der Election-API nach Azure Functions (Consumption).
#
# Beispiel (neu anlegen + deployen):
#   cd backend\election-api
#   .\scripts\ensure-entra-election-scope.ps1
#   .\scripts\deploy-azure.ps1 -ResourceGroup "rg-ms365schule-kursteams-dev" `
#       -FunctionAppName "func-ms365-election-dev" -CreateResources
#
# Frontend: MS365_ELECTION_API.baseUrl = https://<FunctionAppName>.azurewebsites.net/api/election

param(
    [Parameter(Mandatory = $true)]
    [string] $ResourceGroup,

    [Parameter(Mandatory = $true)]
    [string] $FunctionAppName,

    [string] $Location = "westeurope",

    [string] $StorageAccountName = "",

    [string] $ApiAppId = "12e0cfe2-8337-4b35-93e8-542faf658eb3",

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

if ($CreateResources) {
    if (-not $StorageAccountName) {
        $StorageAccountName = "stms365elec" + (New-RandomAlnum 6)
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
        Write-Host "Function App: $FunctionAppName"
        az functionapp create `
            --resource-group $ResourceGroup `
            --consumption-plan-location $Location `
            --runtime node `
            --runtime-version 22 `
            --functions-version 4 `
            --name $FunctionAppName `
            --storage-account $StorageAccountName `
            --os-type Linux | Out-Null
        if ($LASTEXITCODE -ne 0) { throw "Function App konnte nicht erstellt werden." }
    } else {
        Write-Host "Function App existiert: $FunctionAppName"
    }
}

# App Settings (Pepper nur setzen, wenn noch nicht vorhanden)
$audiences = "$ApiAppId,api://$ApiAppId"
$existingPepper = az functionapp config appsettings list `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --query "[?name=='ELECTION_VOTER_PEPPER'].value | [0]" -o tsv 2>$null
$pepper = if ($existingPepper) {
    Write-Host "ELECTION_VOTER_PEPPER bleibt unverändert."
    $existingPepper
} else {
    Write-Host "ELECTION_VOTER_PEPPER neu erzeugen …"
    -join ((48..57 + 65..90 + 97..122) | Get-Random -Count 48 | ForEach-Object { [char]$_ })
}

Write-Host "App Settings setzen …"
az functionapp config appsettings set `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --settings `
        "AZURE_CLIENT_ID=$ApiAppId" `
        "ELECTION_TOKEN_AUDIENCES=$audiences" `
        "ELECTION_REQUIRED_SCOPES=Election.Access,License.Access" `
        "ELECTION_VOTER_PEPPER=$pepper" `
        "FUNCTIONS_WORKER_RUNTIME=node" | Out-Null

if (-not $SkipPublish) {
    Write-Host "npm ci …"
    if (Test-Path "package-lock.json") {
        npm ci --omit=dev
    } else {
        npm install --omit=dev
    }

    Write-Host "Publish (zip) → $FunctionAppName …"
    if (-not (Get-Command func -ErrorAction SilentlyContinue)) {
        throw "Azure Functions Core Tools (func) erforderlich für: func azure functionapp publish"
    }
    func azure functionapp publish $FunctionAppName --javascript
}

$corsOrigins = @(
    "http://localhost:5173",
    "http://localhost:5174",
    "http://localhost:5175",
    "http://localhost:4173",
    "http://127.0.0.1:5173",
    "http://127.0.0.1:5174",
    "http://127.0.0.1:5175",
    "https://kurtrocks.github.io",
    "https://kurtsoeser.github.io",
    "https://app.ms365.schule"
)
Write-Host "CORS-Origins sicherstellen …"
az functionapp cors add `
    --name $FunctionAppName `
    --resource-group $ResourceGroup `
    --allowed-origins $corsOrigins 2>$null | Out-Null

$base = "https://$FunctionAppName.azurewebsites.net/api/election"
Write-Host ""
Write-Host "Fertig."
Write-Host "  Health:  $base/health"
Write-Host "  baseUrl: $base"
Write-Host "  Scope:   api://$ApiAppId/Election.Access"
Write-Host ""
Write-Host "Frontend (ms365-config.js):"
Write-Host "  window.MS365_ELECTION_API = { baseUrl: '$base', scope: 'api://$ApiAppId/Election.Access', fallbackScope: 'api://$ApiAppId/License.Access' };"
