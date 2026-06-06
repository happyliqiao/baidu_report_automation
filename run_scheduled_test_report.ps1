param()

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$logDir = Join-Path $root 'logs'
New-Item -ItemType Directory -Path $logDir -Force | Out-Null

$logPath = Join-Path $logDir ('scheduled_test_report_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.log')

Start-Transcript -Path $logPath -Append | Out-Null
try {
    Set-Location $root
    Write-Host ('Scheduled test report started at ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
    & (Join-Path $root 'run_daily_report.ps1')
    Write-Host ('Scheduled test report finished at ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
} finally {
    Stop-Transcript | Out-Null
}
