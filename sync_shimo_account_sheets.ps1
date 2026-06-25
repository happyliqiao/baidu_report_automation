param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
    [string]$Date = '',
    [string]$AccountUsername = '',
    [string]$SourceCsvPaths = '',
    [string]$StartDate = '',
    [string]$EndDate = '',
    [int]$Days = 0,
    [string]$OutputJson = '',
    [switch]$Setup,
    [switch]$DryRun,
    [switch]$ScanMissing
)

$ErrorActionPreference = 'Stop'

function Write-Log {
    param([string]$Message)

    $line = '{0} {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Message
    $logRoots = @(
        (Join-Path $PSScriptRoot 'logs'),
        (Join-Path $env:TEMP 'baidu_report_automation\logs')
    )
    foreach ($logDir in $logRoots) {
        try {
            New-Item -ItemType Directory -Path $logDir -Force | Out-Null
            Add-Content -Path (Join-Path $logDir 'shimo_account_sheet_sync.log') -Value $line -Encoding UTF8
            Write-Host $line
            return
        } catch {
            continue
        }
    }

    Write-Warning ('Failed to write Shimo sync log to project and temp directories: ' + $line)
    Write-Host $line
}

if (-not (Test-Path $ConfigPath)) {
    throw ('Config file not found: ' + $ConfigPath)
}

$nodeScript = Join-Path $PSScriptRoot 'scripts\shimo_account_sheet_sync.js'
if (-not (Test-Path $nodeScript)) {
    throw ('Node script not found: ' + $nodeScript)
}

$args = @(
    $nodeScript,
    ('--config=' + $ConfigPath)
)
if ($Setup) { $args += '--setup' }
if ($DryRun) { $args += '--dry-run' }
if ($ScanMissing) { $args += '--scan-missing' }
if (-not [string]::IsNullOrWhiteSpace($Date)) { $args += ('--date=' + $Date) }
if (-not [string]::IsNullOrWhiteSpace($AccountUsername)) { $args += ('--account=' + $AccountUsername) }
if (-not [string]::IsNullOrWhiteSpace($SourceCsvPaths)) { $args += ('--source-csvs=' + $SourceCsvPaths) }
if (-not [string]::IsNullOrWhiteSpace($StartDate)) { $args += ('--start-date=' + $StartDate) }
if (-not [string]::IsNullOrWhiteSpace($EndDate)) { $args += ('--end-date=' + $EndDate) }
if ($Days -gt 0) { $args += ('--days=' + $Days) }
if (-not [string]::IsNullOrWhiteSpace($OutputJson)) { $args += ('--output-json=' + $OutputJson) }

Write-Log 'Start Shimo account-sheet sync.'
node @args
if ($LASTEXITCODE -ne 0) {
    throw 'Shimo account-sheet sync failed.'
}
Write-Log 'Done Shimo account-sheet sync.'
