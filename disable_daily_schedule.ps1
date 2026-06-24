param(
    [string]$TaskName = 'BaiduReportAutomation_Daily_0850'
)

$ErrorActionPreference = 'Stop'

try {
    $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
    Disable-ScheduledTask -InputObject $task | Out-Null
    Write-Host ('Disabled scheduled task: ' + $TaskName)
} catch {
    $message = $_.Exception.Message
    if ($message -match 'cannot find' -or $message -match 'No MSFT_ScheduledTask objects found') {
        Write-Host ('Scheduled task not found: ' + $TaskName)
        exit 0
    }

    throw
}
