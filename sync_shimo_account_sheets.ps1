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

    $logDir = Join-Path $PSScriptRoot 'logs'
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    $line = '{0} {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Message
    $logPath = Join-Path $logDir 'shimo_account_sheet_sync.log'
    $written = $false
    for ($attempt = 1; $attempt -le 5 -and -not $written; $attempt++) {
        try {
            Add-Content -Path $logPath -Value $line -Encoding UTF8
            $written = $true
        } catch {
            Start-Sleep -Milliseconds (200 * $attempt)
        }
    }
    if (-not $written) {
        $fallbackPath = Join-Path $logDir ('shimo_account_sheet_sync_' + $PID + '.log')
        try {
            Add-Content -Path $fallbackPath -Value $line -Encoding UTF8
        } catch {
            Write-Warning ('Failed to write Shimo sync log: ' + $_.Exception.Message)
        }
    }
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
