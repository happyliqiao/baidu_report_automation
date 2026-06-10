param()

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot

function Write-Step {
    param([string]$Message)
    Write-Host ('[TEST] ' + $Message)
}

function Assert-PathExists {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        throw ('Missing required path: ' + $Path)
    }
}

function Test-PowerShellSyntax {
    param([string]$Path)

    $tokens = $null
    $errors = $null
    [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors) | Out-Null
    if ($errors -and $errors.Count -gt 0) {
        $messages = @($errors | ForEach-Object { $_.Message }) -join '; '
        throw ('PowerShell syntax failed for ' + $Path + ': ' + $messages)
    }
}

Write-Step 'Check required files and directories.'
$requiredPaths = @(
    'README.md',
    'config.json',
    'package.json',
    'package-lock.json',
    'run_daily_report.ps1',
    'configure_daily_schedule.ps1',
    'baidu_report.ps1',
    'ad_report.ps1',
    'refresh_baidu_token.ps1',
    'sync_shimo_account_sheets.ps1',
    'run_scheduled_full_report.ps1',
    'run_scheduled_test_report.ps1',
    'scripts\shimo_account_sheet_sync.js',
    'node_modules\playwright-core',
    'backups\baidu_report_automation_versions_merged'
)
foreach ($relativePath in $requiredPaths) {
    Assert-PathExists (Join-Path $root $relativePath)
}

Write-Step 'Validate config.json.'
$configPath = Join-Path $root 'config.json'
$config = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $config.accounts) {
    throw 'config.json missing accounts.'
}
if (-not $config.providers -or -not $config.providers.qihoo360 -or -not $config.providers.qihoo360.accounts) {
    throw 'config.json missing providers.qihoo360.accounts.'
}
if (-not $config.shimoAccountSheets) {
    throw 'config.json missing shimoAccountSheets.'
}

Write-Step 'Validate package.json.'
$package = Get-Content -LiteralPath (Join-Path $root 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $package.dependencies -or -not $package.dependencies.'playwright-core') {
    throw 'package.json missing dependency: playwright-core.'
}

Write-Step 'Check PowerShell syntax.'
$psScripts = @(
    'run_daily_report.ps1',
    'configure_daily_schedule.ps1',
    'baidu_report.ps1',
    'ad_report.ps1',
    'refresh_baidu_token.ps1',
    'sync_shimo_account_sheets.ps1',
    'run_scheduled_full_report.ps1',
    'run_scheduled_test_report.ps1'
)
foreach ($script in $psScripts) {
    Test-PowerShellSyntax (Join-Path $root $script)
}

Write-Step 'Check Node script syntax.'
Push-Location $root
try {
    node --check (Join-Path $root 'scripts\shimo_account_sheet_sync.js')
    if ($LASTEXITCODE -ne 0) {
        throw 'node --check failed for scripts\shimo_account_sheet_sync.js.'
    }
} finally {
    Pop-Location
}

Write-Step 'All safe checks passed.'
