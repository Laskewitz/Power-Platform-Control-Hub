<#
.SYNOPSIS
    Deploy Power Platform Control Hub to a target environment.

.PARAMETER Environment
    The environment alias to deploy to: 'dev' or 'prod'.
    Connection IDs and environment config are read from env-config.json.

.EXAMPLE
    .\Deploy.ps1 -Environment prod
    .\Deploy.ps1 -Environment dev
#>
param(
    [Parameter(Mandatory)]
    [ValidateSet('dev', 'prod')]
    [string]$Environment
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$rootDir       = Resolve-Path "$PSScriptRoot\.."
$configPath    = "$PSScriptRoot\env-config.json"
$powerCfgPath  = "$rootDir\power.config.json"
$solutionId    = "PowerPlatformControlHub"

# ── Load environment config ───────────────────────────────────────────────────
$config  = Get-Content $configPath -Raw | ConvertFrom-Json
$envCfg  = $config.environments.$Environment

Write-Host ""
Write-Host "  Deploying to: $($envCfg.name)" -ForegroundColor Cyan
Write-Host "  Environment:  $($envCfg.environmentId)" -ForegroundColor Cyan
Write-Host ""

# ── Patch power.config.json ───────────────────────────────────────────────────
$powerCfg = Get-Content $powerCfgPath -Raw | ConvertFrom-Json
$powerCfg.environmentId = $envCfg.environmentId
$powerCfg.appId         = if ($envCfg.appId) { $envCfg.appId } else { "" }
$powerCfg | ConvertTo-Json -Depth 20 | Set-Content $powerCfgPath -Encoding UTF8
Write-Host "  power.config.json updated." -ForegroundColor DarkGray

# ── Add / refresh each data source ───────────────────────────────────────────
$connections = $envCfg.connections
foreach ($apiId in $connections.PSObject.Properties.Name) {
    $connId = $connections.$apiId
    if ($connId -like "<fill-in*>") {
        Write-Warning "  Skipping $apiId — connection ID not configured in env-config.json"
        continue
    }
    Write-Host "  Configuring: $apiId ($connId)" -ForegroundColor DarkGray
    & npx power-apps add-data-source --api-id $apiId --connection-id $connId --non-interactive
    if ($LASTEXITCODE -ne 0) { throw "add-data-source failed for $apiId" }
}

# ── Build ─────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  Building app..." -ForegroundColor Cyan
Push-Location $rootDir
try {
    & npm run build
    if ($LASTEXITCODE -ne 0) { throw "build failed" }
} finally {
    Pop-Location
}

# ── Push ──────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "  Pushing app..." -ForegroundColor Cyan
& npx power-apps push --solution-id $solutionId
if ($LASTEXITCODE -ne 0) { throw "push failed" }

# ── Save updated appId back to env-config.json ───────────────────────────────
$updatedPowerCfg = Get-Content $powerCfgPath -Raw | ConvertFrom-Json
$config.environments.$Environment.appId = $updatedPowerCfg.appId
$config | ConvertTo-Json -Depth 10 | Set-Content $configPath -Encoding UTF8

Write-Host ""
Write-Host "  Done! appId saved: $($updatedPowerCfg.appId)" -ForegroundColor Green
Write-Host "  Play URL: https://apps.powerapps.com/play/e/$($envCfg.environmentId)/app/$($updatedPowerCfg.appId)" -ForegroundColor Green
Write-Host ""
