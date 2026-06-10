param(
    [string]$TaskName = 'BaiduReportAutomation_Daily_0850'
)

$ErrorActionPreference = 'Stop'

$root = $PSScriptRoot
$runner = Join-Path $root 'run_scheduled_full_report.ps1'

if (-not (Test-Path -LiteralPath $runner)) {
    throw ('Missing scheduled runner: ' + $runner)
}

function Test-TaskUsesProject {
    param([object]$Task)

    foreach ($action in @($Task.Actions)) {
        $workingDirectory = [string]$action.WorkingDirectory
        $arguments = [string]$action.Arguments

        if ($workingDirectory.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }

        if ($arguments.IndexOf($root, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
            return $true
        }
    }

    return $false
}

$existingProjectTasks = Get-ScheduledTask | Where-Object { Test-TaskUsesProject $_ }
foreach ($task in @($existingProjectTasks)) {
    Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false
}

$action = New-ScheduledTaskAction `
    -Execute 'powershell.exe' `
    -Argument ('-NoProfile -ExecutionPolicy Bypass -File "' + $runner + '"') `
    -WorkingDirectory $root

$trigger = New-ScheduledTaskTrigger -Daily -At '08:50'
$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -StartWhenAvailable

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $action `
    -Trigger $trigger `
    -Settings $settings `
    -Description 'Run baidu report automation once daily at 08:50.' `
    -Force | Out-Null

Write-Host ('Configured scheduled task: ' + $TaskName + ' at 08:50 daily.')
