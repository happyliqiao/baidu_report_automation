param(
    [string]$Date = '',
    [string]$StartDate = '',
    [string]$EndDate = '',
    [int]$Days = 0,
    [switch]$UseShimoMissingDates,
    [switch]$SkipShimoSync
)

$ErrorActionPreference = 'Stop'

$useShimoMissingDatesExplicit = $PSBoundParameters.ContainsKey('UseShimoMissingDates')
if (-not $useShimoMissingDatesExplicit) {
    $UseShimoMissingDates = $true
}

function Set-DefaultDirectReportDate {
    if ([string]::IsNullOrWhiteSpace($script:Date) -and
        [string]::IsNullOrWhiteSpace($script:StartDate) -and
        [string]::IsNullOrWhiteSpace($script:EndDate) -and
        $script:Days -le 0 -and
        -not $script:UseShimoMissingDates) {
        $script:Date = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd')
    }
}

Set-DefaultDirectReportDate

$baiduDir = $PSScriptRoot
$qihooDir = $PSScriptRoot

$refreshScript = Join-Path $baiduDir 'refresh_baidu_token.ps1'
$baiduRunScript = Join-Path $baiduDir 'baidu_report.ps1'
$qihooRunScript = Join-Path $qihooDir 'ad_report.ps1'
$shimoSyncScript = Join-Path $baiduDir 'sync_shimo_account_sheets.ps1'
$shimoConfigPath = Join-Path $baiduDir 'config.json'

if (-not (Test-Path $shimoConfigPath)) {
    throw ('Config file not found: ' + $shimoConfigPath)
}

$rootConfig = Get-Content -Path $shimoConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$shimoEnabled = $false
if ($rootConfig.shimoAccountSheets) {
    $shimoEnabled = [bool]$rootConfig.shimoAccountSheets.enabled
}

if ($UseShimoMissingDates -and -not $shimoEnabled) {
    if ($useShimoMissingDatesExplicit) {
        throw 'UseShimoMissingDates was explicitly requested, but shimoAccountSheets.enabled is false in config.json.'
    }
    Write-Host 'Shimo account-sheet sync is disabled in config.json. Fall back to direct date-based report fetch.'
    $UseShimoMissingDates = $false
    Set-DefaultDirectReportDate
}

function Get-NormalizedDate {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ''
    }
    ([datetime]::Parse($Value)).ToString('yyyy-MM-dd')
}

function Get-ReportDateList {
    param(
        [string]$Date,
        [string]$StartDate,
        [string]$EndDate,
        [int]$Days
    )

    $format = 'yyyy-MM-dd'
    if ($Days -gt 0) {
        if ([string]::IsNullOrWhiteSpace($EndDate)) {
            $end = (Get-Date).Date
        } else {
            $end = [datetime]::ParseExact((Get-NormalizedDate $EndDate), $format, [Globalization.CultureInfo]::InvariantCulture)
        }
        $start = $end.AddDays(-1 * ($Days - 1))
    } elseif (-not [string]::IsNullOrWhiteSpace($StartDate) -or -not [string]::IsNullOrWhiteSpace($EndDate)) {
        if ([string]::IsNullOrWhiteSpace($StartDate)) {
            $StartDate = $EndDate
        }
        if ([string]::IsNullOrWhiteSpace($EndDate)) {
            $EndDate = $StartDate
        }
        $start = [datetime]::ParseExact((Get-NormalizedDate $StartDate), $format, [Globalization.CultureInfo]::InvariantCulture)
        $end = [datetime]::ParseExact((Get-NormalizedDate $EndDate), $format, [Globalization.CultureInfo]::InvariantCulture)
    } else {
        if ([string]::IsNullOrWhiteSpace($Date)) {
            $Date = (Get-Date).ToString($format)
        }
        $start = [datetime]::ParseExact((Get-NormalizedDate $Date), $format, [Globalization.CultureInfo]::InvariantCulture)
        $end = $start
    }

    if ($start -gt $end) {
        throw ('StartDate must be earlier than or equal to EndDate: ' + $start.ToString($format) + ' > ' + $end.ToString($format))
    }

    $dates = @()
    for ($cursor = $start; $cursor -le $end; $cursor = $cursor.AddDays(1)) {
        $dates += $cursor.ToString($format)
    }
    $dates
}

function Get-AccountSourceMap {
    param([string]$BaiduConfigPath, [string]$QihooConfigPath)

    $result = @{}
    $baiduConfig = Get-Content -Path $BaiduConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($account in @($baiduConfig.accounts | Where-Object { $_.enabled })) {
        $key = ([string]$account.username).Trim()
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $result[$key] = 'baidu'
        }
    }

    $qihooConfig = Get-Content -Path $QihooConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    foreach ($account in @($qihooConfig.providers.qihoo360.accounts | Where-Object { $_.enabled })) {
        $key = ([string]$account.username).Trim()
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $result[$key] = '360'
        }
    }
    $result
}

function Get-EnabledBaiduAccounts {
    param($Config)

    @($Config.accounts | Where-Object { $_.enabled })
}

function Test-BaiduAccountHasRefreshInput {
    param($Account)

    foreach ($field in @('refreshToken', 'authorizationCode')) {
        if ($Account.PSObject.Properties.Name -contains $field -and
            -not [string]::IsNullOrWhiteSpace([string]$Account.$field)) {
            return $true
        }
    }

    $false
}

function Test-BaiduAccountHasAccessToken {
    param($Account)

    if ($Account.PSObject.Properties.Name -contains 'accessToken' -and
        -not [string]::IsNullOrWhiteSpace([string]$Account.accessToken)) {
        return $true
    }

    $false
}

function Get-ConfigAccountByUsername {
    param(
        $Config,
        [string]$Username
    )

    foreach ($account in @(Get-EnabledBaiduAccounts -Config $Config)) {
        if ([string]$account.username -eq $Username) {
            return $account
        }
    }

    $null
}

function Invoke-ReportScriptForDate {
    param(
        [string]$Name,
        [string]$ScriptPath,
        [string]$WorkDir,
        [string]$ReportDate,
        [string]$ConfigPath,
        [string]$AccountUsername = ''
    )

    if (-not (Test-Path $ScriptPath)) {
        throw ($Name + ' script not found: ' + $ScriptPath)
    }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-ConfigPath', $ConfigPath,
        '-Date', $ReportDate
    )
    if (-not [string]::IsNullOrWhiteSpace($AccountUsername)) {
        $args += @('-AccountUsername', $AccountUsername)
    }

    Push-Location $WorkDir
    try {
        Write-Host ('=== Run ' + $Name + ' report for ' + $ReportDate + ' ===')
        powershell.exe @args
        if ($LASTEXITCODE -ne 0) {
            throw ($Name + ' report failed for ' + $ReportDate)
        }
    } finally {
        Pop-Location
    }
}

function Invoke-ReportScriptForDateRange {
    param(
        [string]$Name,
        [string]$ScriptPath,
        [string]$WorkDir,
        [string]$StartDate,
        [string]$EndDate,
        [string]$ConfigPath
    )

    if (-not (Test-Path $ScriptPath)) {
        throw ($Name + ' script not found: ' + $ScriptPath)
    }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-ConfigPath', $ConfigPath,
        '-StartDate', $StartDate,
        '-EndDate', $EndDate
    )

    Push-Location $WorkDir
    try {
        if ($StartDate -eq $EndDate) {
            Write-Host ('=== Run ' + $Name + ' reports for ' + $StartDate + ' ===')
        } else {
            Write-Host ('=== Run ' + $Name + ' reports for ' + $StartDate + ' to ' + $EndDate + ' ===')
        }
        powershell.exe @args
        if ($LASTEXITCODE -ne 0) {
            throw ($Name + ' report failed for ' + $StartDate + ' to ' + $EndDate)
        }
    } finally {
        Pop-Location
    }
}

function Invoke-BaiduTokenRefresh {
    param(
        [string]$ScriptPath,
        [string]$ConfigPath,
        [string]$AccountUsername = ''
    )

    if (-not (Test-Path $ScriptPath)) {
        throw ('Refresh script not found: ' + $ScriptPath)
    }
    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-ConfigPath', $ConfigPath,
        '-UpdateConfig',
        '-SkipMissingRefreshToken',
        '-FailOnAccountError'
    )
    if (-not [string]::IsNullOrWhiteSpace($AccountUsername)) {
        $args += @('-AccountUsername', $AccountUsername)
    }
    powershell.exe @args
    if ($LASTEXITCODE -ne 0) {
        throw 'Token refresh failed.'
    }
}

function Invoke-ShimoSyncForDate {
    param(
        [string]$ScriptPath,
        [string]$ConfigPath,
        [string]$ReportDate,
        [string]$AccountUsername = ''
    )

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-ConfigPath', $ConfigPath,
        '-Date', $ReportDate
    )
    if (-not [string]::IsNullOrWhiteSpace($AccountUsername)) {
        $args += @('-AccountUsername', $AccountUsername)
    }
    Write-Host ('=== Sync Shimo account sheets for ' + $ReportDate + ' ===')
    powershell.exe @args
    if ($LASTEXITCODE -ne 0) {
        throw ('Shimo account-sheet sync failed for ' + $ReportDate)
    }
}

function Invoke-ShimoSyncForDateRange {
    param(
        [string]$ScriptPath,
        [string]$ConfigPath,
        [string]$StartDate,
        [string]$EndDate
    )

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-ConfigPath', $ConfigPath,
        '-StartDate', $StartDate,
        '-EndDate', $EndDate
    )
    if ($StartDate -eq $EndDate) {
        Write-Host ('=== Sync Shimo account sheets for ' + $StartDate + ' ===')
    } else {
        Write-Host ('=== Sync Shimo account sheets for ' + $StartDate + ' to ' + $EndDate + ' ===')
    }
    powershell.exe @args
    if ($LASTEXITCODE -ne 0) {
        throw ('Shimo account-sheet sync failed for ' + $StartDate + ' to ' + $EndDate)
    }
}

function Get-DateBounds {
    param([object[]]$Runs)

    $dates = @($Runs | ForEach-Object { Get-NormalizedDate ([string]$_.Date) } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Sort-Object)
    if ($dates.Count -eq 0) {
        return $null
    }

    [pscustomobject]@{
        StartDate = $dates[0]
        EndDate = $dates[$dates.Count - 1]
    }
}

function Invoke-ReportScript {
    param(
        [string]$Name,
        [string]$ScriptPath,
        [string]$WorkDir,
        [switch]$SkipSync
    )

    if (-not (Test-Path $ScriptPath)) {
        throw ($Name + ' script not found: ' + $ScriptPath)
    }

    $args = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $ScriptPath,
        '-ConfigPath', $shimoConfigPath
    )
    if (-not [string]::IsNullOrWhiteSpace($Date)) {
        $args += @('-Date', $Date)
    }
    if (-not [string]::IsNullOrWhiteSpace($StartDate)) {
        $args += @('-StartDate', $StartDate)
    }
    if (-not [string]::IsNullOrWhiteSpace($EndDate)) {
        $args += @('-EndDate', $EndDate)
    }
    if ($Days -gt 0) {
        $args += @('-Days', $Days)
    }
    Push-Location $WorkDir
    try {
        Write-Host ('=== Run ' + $Name + ' report ===')
        powershell.exe @args
        if ($LASTEXITCODE -ne 0) {
            throw ($Name + ' report failed.')
        }
    } finally {
        Pop-Location
    }
}

if ($UseShimoMissingDates) {
    if ($SkipShimoSync) {
        throw 'UseShimoMissingDates requires Shimo access; do not combine it with -SkipShimoSync.'
    }
    if (-not (Test-Path $shimoSyncScript)) {
        throw ('Shimo account-sheet sync script not found: ' + $shimoSyncScript)
    }

    $planDir = Join-Path $baiduDir 'data'
    New-Item -ItemType Directory -Path $planDir -Force | Out-Null
    $planPath = Join-Path $planDir ('shimo_missing_plan_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.json')
    $scanArgs = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $shimoSyncScript,
        '-ConfigPath', $shimoConfigPath,
        '-ScanMissing',
        '-OutputJson', $planPath
    )
    if (-not [string]::IsNullOrWhiteSpace($Date)) { $scanArgs += @('-Date', $Date) }
    if (-not [string]::IsNullOrWhiteSpace($StartDate)) { $scanArgs += @('-StartDate', $StartDate) }
    if (-not [string]::IsNullOrWhiteSpace($EndDate)) { $scanArgs += @('-EndDate', $EndDate) }
    if ($Days -gt 0) { $scanArgs += @('-Days', $Days) }

    Write-Host '=== Scan Shimo missing dates ==='
    powershell.exe @scanArgs
    if ($LASTEXITCODE -ne 0) {
        throw 'Shimo missing-date scan failed.'
    }

    $plan = Get-Content -Path $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $sourceMap = Get-AccountSourceMap -BaiduConfigPath $shimoConfigPath -QihooConfigPath $shimoConfigPath
    $missingItems = @($plan.items | Where-Object { $_.matched -and $_.missingDates -and @($_.missingDates).Count -gt 0 })
    if ($missingItems.Count -eq 0) {
        Write-Host 'No Shimo missing dates found. Nothing to fetch or sync.'
        exit 0
    }

    $runKeys = @{}
    foreach ($item in $missingItems) {
        $account = ([string]$item.account).Trim()
        if (-not $sourceMap.ContainsKey($account)) {
            Write-Host ('Skip missing account without enabled local config: ' + $account)
            continue
        }
        foreach ($missingDate in @($item.missingDates)) {
            $dateText = Get-NormalizedDate ([string]$missingDate)
            $key = $sourceMap[$account] + '|' + $account + '|' + $dateText
            $runKeys[$key] = [pscustomobject]@{
                Source = $sourceMap[$account]
                Account = $account
                Date = $dateText
            }
        }
    }

    $needsBaidu = @($runKeys.Values | Where-Object { $_.Source -eq 'baidu' }).Count -gt 0
    if ($needsBaidu) {
        $baiduRefreshAccounts = @(
            $runKeys.Values |
                Where-Object { $_.Source -eq 'baidu' } |
                ForEach-Object { [string]$_.Account } |
                Sort-Object -Unique
        )
        foreach ($account in $baiduRefreshAccounts) {
            $configAccount = Get-ConfigAccountByUsername -Config $rootConfig -Username $account
            if ($null -eq $configAccount) {
                throw ('Baidu account not found in config: ' + $account)
            }

            if (Test-BaiduAccountHasRefreshInput $configAccount) {
                Invoke-BaiduTokenRefresh -ScriptPath $refreshScript -ConfigPath $shimoConfigPath -AccountUsername $account
            } elseif (Test-BaiduAccountHasAccessToken $configAccount) {
                Write-Host ('No refresh token for Baidu account ' + $account + '. Use existing accessToken for missing-date fetch.')
            } else {
                throw ('Baidu account ' + $account + ' has neither refresh input nor accessToken.')
            }
        }
    }

    $baiduRuns = @($runKeys.Values | Where-Object { $_.Source -eq 'baidu' })
    $qihooRuns = @($runKeys.Values | Where-Object { $_.Source -eq '360' })
    $allRuns = @($runKeys.Values)

    $baiduBounds = Get-DateBounds -Runs $baiduRuns
    if ($baiduBounds) {
        Invoke-ReportScriptForDateRange -Name 'Baidu' -ScriptPath $baiduRunScript -WorkDir $baiduDir -StartDate $baiduBounds.StartDate -EndDate $baiduBounds.EndDate -ConfigPath $shimoConfigPath
    }

    $qihooBounds = Get-DateBounds -Runs $qihooRuns
    if ($qihooBounds) {
        Invoke-ReportScriptForDateRange -Name '360' -ScriptPath $qihooRunScript -WorkDir $qihooDir -StartDate $qihooBounds.StartDate -EndDate $qihooBounds.EndDate -ConfigPath $shimoConfigPath
    }

    $syncBounds = Get-DateBounds -Runs $allRuns
    if ($syncBounds) {
        Invoke-ShimoSyncForDateRange -ScriptPath $shimoSyncScript -ConfigPath $shimoConfigPath -StartDate $syncBounds.StartDate -EndDate $syncBounds.EndDate
    }
} else {
    $enabledBaiduAccounts = @(Get-EnabledBaiduAccounts -Config $rootConfig)
    $refreshableBaiduAccounts = @($enabledBaiduAccounts | Where-Object { Test-BaiduAccountHasRefreshInput $_ })
    $hasAnyBaiduAccessToken = @($enabledBaiduAccounts | Where-Object { Test-BaiduAccountHasAccessToken $_ }).Count -gt 0

    if ($refreshableBaiduAccounts.Count -gt 0) {
        Invoke-BaiduTokenRefresh -ScriptPath $refreshScript -ConfigPath $shimoConfigPath
    } elseif ($enabledBaiduAccounts.Count -gt 0) {
        if ($hasAnyBaiduAccessToken) {
            Write-Host 'No refreshable Baidu account found in config.json. Skip token refresh and use existing accessToken values.'
        } else {
            Write-Host 'No refreshable Baidu account found in config.json. Skip Baidu token refresh.'
        }
    }

    Invoke-ReportScript -Name 'Baidu' -ScriptPath $baiduRunScript -WorkDir $baiduDir -SkipSync
    Invoke-ReportScript -Name '360' -ScriptPath $qihooRunScript -WorkDir $qihooDir -SkipSync
}

if (-not $SkipShimoSync -and -not $UseShimoMissingDates) {
    if (-not (Test-Path $shimoSyncScript)) {
        throw ('Shimo account-sheet sync script not found: ' + $shimoSyncScript)
    }

    Write-Host '=== Sync Shimo account sheets ==='
    $syncArgs = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', $shimoSyncScript,
        '-ConfigPath', $shimoConfigPath
    )
    if (-not [string]::IsNullOrWhiteSpace($Date)) {
        $syncArgs += @('-Date', $Date)
    }
    powershell.exe @syncArgs
    if ($LASTEXITCODE -ne 0) {
        throw 'Shimo account-sheet sync failed.'
    }
}

Write-Host '=== All reports done ==='
