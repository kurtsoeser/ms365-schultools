# Deploy der License-API nach Azure Functions (Consumption).
# Voraussetzung: Azure CLI (az), angemeldet; Functions Core Tools (func) optional für Zip-Deploy.
#
# Beispiel:
#   cd backend\license-api
#   .\scripts\deploy-azure.ps1 -ResourceGroup "rg-ms365schule" -FunctionAppName "func-ms365-license-dev"
#
# Danach App Settings setzen (Portal oder az) – Werte aus local.settings.json:
#   AZURE_TENANT_ID, AZURE_CLIENT_ID, AZURE_CLIENT_SECRET
#   LICENSE_SITE_WEB_URL, LICENSE_LIST_NAME, LICENSE_SPA_CLIENT_ID, LICENSE_TOKEN_AUDIENCES
#
# Frontend: MS365_LICENSE_API.baseUrl = https://<FunctionAppName>.azurewebsites.net/api/license
# GitHub Secret: LICENSE_API_BASE_URL (gleiche URL) für Pages-Deploy.

param(
    [Parameter(Mandatory = $true)]
    [string] $ResourceGroup,

    [Parameter(Mandatory = $true)]
    [string] $FunctionAppName,

    [string] $Location = "westeurope",

    [string] $StorageAccountName = "",

    [switch] $CreateResources
)

$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $PSScriptRoot
Set-Location $root

if (-not (Get-Command az -ErrorAction SilentlyContinue)) {
    throw "Azure CLI (az) nicht gefunden. Bitte installieren und 'az login' ausführen."
}

if ($CreateResources) {
    if (-not $StorageAccountName) {
        $StorageAccountName = ("stms365lic" + -join ((48..57 + 97..122) | Get-Random -Count 6 | ForEach-Object { [char]$_ }))
    }
    Write-Host "Resource Group: $ResourceGroup ($Location)"
    az group create --name $ResourceGroup --location $Location | Out-Null

    Write-Host "Storage Account: $StorageAccountName"
    az storage account create `
        --name $StorageAccountName `
        --resource-group $ResourceGroup `
        --location $Location `
        --sku Standard_LRS | Out-Null

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
}

Write-Host "Publish (zip) → $FunctionAppName …"
if (-not (Get-Command func -ErrorAction SilentlyContinue)) {
    Write-Host "Hinweis: 'func' fehlt – nutze az functionapp deployment source config-zip nach manuellem Zip."
    throw "Azure Functions Core Tools (func) erforderlich für: func azure functionapp publish"
}

func azure functionapp publish $FunctionAppName --javascript

Write-Host ""
Write-Host "Fertig. Health-Check:"
Write-Host "  https://$FunctionAppName.azurewebsites.net/api/license/health"
Write-Host "baseUrl für Frontend / GitHub Secret LICENSE_API_BASE_URL:"
Write-Host "  https://$FunctionAppName.azurewebsites.net/api/license"
