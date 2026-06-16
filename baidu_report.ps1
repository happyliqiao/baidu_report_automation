param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
    [string]$Date = '',
    [string]$StartDate = '',
    [string]$EndDate = '',
    [int]$Days = 0,
    [string]$AccountUsername = ''
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Write-Log {
    param([string]$Message)

    $logDir = Join-Path $PSScriptRoot 'logs'
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    $line = '{0} {1}' -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $Message
    $logPath = Join-Path $logDir 'baidu_report.log'
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
        $fallbackPath = Join-Path $logDir ('baidu_report_' + $PID + '.log')
        Add-Content -Path $fallbackPath -Value $line -Encoding UTF8
    }
    Write-Host $line
}

function Invoke-BaiduJson {
    param(
        [string]$Uri,
        [object]$Body,
        [int]$MaxAttempts = 4,
        [int]$InitialDelaySeconds = 2
    )

    $json = $Body | ConvertTo-Json -Depth 30 -Compress

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return Invoke-RestMethod -Uri $Uri -Method Post -Body $json -ContentType 'application/json' -TimeoutSec 60
        } catch {
            $detail = 'POST ' + $Uri + ' failed: ' + $_.Exception.Message
            if ($_.Exception.Response) {
                try {
                    $reader = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())
                    $detail = $detail + ' Response: ' + $reader.ReadToEnd()
                } catch {
                    $detail = $detail + ' Response body could not be read.'
                }
            }

            if ($attempt -ge $MaxAttempts) {
                throw $detail
            }

            $delay = [math]::Min(30, $InitialDelaySeconds * [math]::Pow(2, $attempt - 1))
            Write-Log ('POST retry {0}/{1} in {2}s: {3}' -f ($attempt + 1), $MaxAttempts, $delay, $detail)
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-BaiduFailures {
    param($Response)

    if ($Response -and
        $Response.PSObject.Properties.Name -contains 'header' -and
        $Response.header -and
        $Response.header.PSObject.Properties.Name -contains 'failures' -and
        $Response.header.failures) {
        return @($Response.header.failures)
    }

    @()
}

function Test-BaiduEmptyRequestFailure {
    param($Failures)

    foreach ($failure in @($Failures)) {
        if ($failure.PSObject.Properties.Name -contains 'code' -and [string]$failure.code -eq '901600') {
            return $true
        }
        if ($failure.PSObject.Properties.Name -contains 'message' -and [string]$failure.message -like '*Request could not be empty*') {
            return $true
        }
    }

    $false
}

function Test-BaiduRetryableReportShapeFailure {
    param($Failures)

    foreach ($failure in @($Failures)) {
        if ($failure.PSObject.Properties.Name -contains 'code' -and @('901600', '99912') -contains [string]$failure.code) {
            return $true
        }
        if ($failure.PSObject.Properties.Name -contains 'message') {
            $message = [string]$failure.message
            if ($message -like '*Request could not be empty*' -or $message -like '*Json message*invalid*') {
                return $true
            }
        }
    }

    $false
}

function Invoke-BaiduReport {
    param(
        [string]$Uri,
        [object]$Header,
        [object]$RequestType
    )

    $payloads = @(
        [ordered]@{
            name = 'realTimeRequestType'
            value = [ordered]@{
                header = $Header
                body = [ordered]@{
                    realTimeRequestType = $RequestType
                }
            }
        },
        [ordered]@{
            name = 'realTimeRequestTypes'
            value = [ordered]@{
                header = $Header
                body = [ordered]@{
                    realTimeRequestTypes = @($RequestType)
                }
            }
        },
        [ordered]@{
            name = 'directBody'
            value = [ordered]@{
                header = $Header
                body = $RequestType
            }
        }
    )

    $lastResponse = $null
    foreach ($payload in $payloads) {
        $response = Invoke-BaiduJson -Uri $Uri -Body $payload.value
        $lastResponse = $response
        $failures = Get-BaiduFailures -Response $response
        if (-not $failures -or $failures.Count -eq 0) {
            Write-Log ('Baidu report request shape: ' + $payload.name)
            return $response
        }
        if (-not (Test-BaiduRetryableReportShapeFailure -Failures $failures)) {
            Write-Log ('Baidu report request shape failed: ' + $payload.name)
            return $response
        }
        Write-Log ('Baidu report request shape retrying: ' + $payload.name)
    }

    $lastResponse
}

function Get-Number {
    param($Value)

    if ($null -eq $Value) { return 0 }
    if ($Value -is [array]) {
        if ($Value.Count -eq 0) { return 0 }
        return Get-Number -Value $Value[0]
    }

    $text = ([string]$Value).Trim()
    if ([string]::IsNullOrWhiteSpace($text) -or $text -eq '--') { return 0 }

    $number = 0.0
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::InvariantCulture, [ref]$number)) {
        return $number
    }
    if ([double]::TryParse($text, [Globalization.NumberStyles]::Any, [Globalization.CultureInfo]::CurrentCulture, [ref]$number)) {
        return $number
    }
    0
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
            $end = (Get-Date).Date.AddDays(-1)
        } else {
            $end = [datetime]::ParseExact($EndDate, $format, [Globalization.CultureInfo]::InvariantCulture)
        }
        $start = $end.AddDays(-1 * ($Days - 1))
    } elseif (-not [string]::IsNullOrWhiteSpace($StartDate) -or -not [string]::IsNullOrWhiteSpace($EndDate)) {
        if ([string]::IsNullOrWhiteSpace($StartDate)) {
            $StartDate = $EndDate
        }
        if ([string]::IsNullOrWhiteSpace($EndDate)) {
            $EndDate = $StartDate
        }
        $start = [datetime]::ParseExact($StartDate, $format, [Globalization.CultureInfo]::InvariantCulture)
        $end = [datetime]::ParseExact($EndDate, $format, [Globalization.CultureInfo]::InvariantCulture)
    } else {
        if ([string]::IsNullOrWhiteSpace($Date)) {
            $Date = (Get-Date).AddDays(-1).ToString($format)
        }
        $start = [datetime]::ParseExact($Date, $format, [Globalization.CultureInfo]::InvariantCulture)
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

function Get-ConfigValue {
    param(
        $Account,
        $Config,
        [string]$Name,
        [object]$DefaultValue = $null
    )

    if ($Account.PSObject.Properties.Name -contains $Name -and -not [string]::IsNullOrWhiteSpace([string]$Account.$Name)) {
        return $Account.$Name
    }
    if ($Config.PSObject.Properties.Name -contains $Name -and -not [string]::IsNullOrWhiteSpace([string]$Config.$Name)) {
        return $Config.$Name
    }
    $DefaultValue
}

function Get-Header {
    param($Account)

    $header = [ordered]@{}
    if (-not [string]::IsNullOrWhiteSpace([string]$Account.userId)) {
        $header.userid = [long]$Account.userId
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Account.baiduUsername)) {
        $header.username = [string]$Account.baiduUsername
        $header.userName = [string]$Account.baiduUsername
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Account.accessToken)) {
        $header.accessToken = [string]$Account.accessToken
    }
    $header
}

function Get-TongjiHeader {
    param($Account)

    $header = [ordered]@{}
    if (-not [string]::IsNullOrWhiteSpace([string]$Account.tongjiUserName)) {
        $header.userName = [string]$Account.tongjiUserName
    } elseif (-not [string]::IsNullOrWhiteSpace([string]$Account.baiduUsername)) {
        $header.userName = [string]$Account.baiduUsername
    } elseif (-not [string]::IsNullOrWhiteSpace([string]$Account.userId)) {
        $header.userid = [long]$Account.userId
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Account.accessToken)) {
        $header.accessToken = [string]$Account.accessToken
    }
    $header
}

function Get-ResultNode {
    param($Response)

    if ($Response.PSObject.Properties.Name -contains 'body' -and $Response.body) {
        $body = $Response.body
        if ($body.PSObject.Properties.Name -contains 'data' -and $body.data) {
            $data = @($body.data)
            if ($data.Count -gt 0) {
                $first = $data[0]
                if ($first.PSObject.Properties.Name -contains 'result' -and $first.result) {
                    return $first.result
                }
                return $data
            }
        }
        if ($body.PSObject.Properties.Name -contains 'result' -and $body.result) {
            return $body.result
        }
        return $body
    }

    if ($Response.PSObject.Properties.Name -contains 'result' -and $Response.result) {
        return $Response.result
    }

    $Response
}

function Get-UniqueStrings {
    param([object[]]$Values)

    $seen = @{}
    $result = @()
    foreach ($value in $Values) {
        foreach ($item in @($value)) {
            $text = ([string]$item).Trim()
            if ([string]::IsNullOrWhiteSpace($text)) {
                continue
            }
            if (-not $seen.ContainsKey($text)) {
                $seen[$text] = $true
                $result += $text
            }
        }
    }
    $result
}

function Get-ConfiguredStringArray {
    param(
        $Account,
        $Config,
        [string]$Name,
        [string[]]$DefaultValue
    )

    if ($Account.PSObject.Properties.Name -contains $Name -and $Account.$Name) {
        return @(Get-UniqueStrings -Values @($Account.$Name))
    }
    if ($Config.PSObject.Properties.Name -contains $Name -and $Config.$Name) {
        return @(Get-UniqueStrings -Values @($Config.$Name))
    }
    @($DefaultValue)
}

function Add-MetricValue {
    param(
        $Totals,
        [string]$Metric,
        $Value,
        [string[]]$ConversionMetrics
    )

    $number = Get-Number -Value $Value
    if (@('impression', 'show_count', 'showCount', 'impressions') -contains $Metric) {
        $Totals.impression += $number
    } elseif (@('click', 'clk_count', 'clkCount', 'clicks') -contains $Metric) {
        $Totals.click += $number
    } elseif (@('cost', 'cost_count', 'costCount') -contains $Metric) {
        $Totals.cost += $number
    } elseif (@($ConversionMetrics) -contains $Metric -or @('conversion', 'trans_count', 'convCount', 'conversions') -contains $Metric) {
        $Totals.conversion += $number
    }
}

function Convert-ResultToTotals {
    param($Result, [string[]]$PerformanceData, [string[]]$ConversionMetrics)

    $totals = [ordered]@{
        impression = 0
        click = 0
        cost = 0
        conversion = 0
    }

    $resultItems = @($Result)
    if ($resultItems.Count -gt 0 -and
        $resultItems[0].PSObject.Properties.Name -contains 'fields' -and
        $resultItems[0].PSObject.Properties.Name -contains 'sum') {
        $Result = $resultItems[0]
        $fields = @($Result.fields)
        $sums = @($Result.sum)
        $sumRow = @()
        if ($sums.Count -gt 0) {
            $sumRow = @($sums[0])
        }

        for ($i = 0; $i -lt $fields.Count -and $i -lt $sumRow.Count; $i++) {
            Add-MetricValue -Totals $totals -Metric ([string]$fields[$i]) -Value $sumRow[$i] -ConversionMetrics $ConversionMetrics
        }
        foreach ($metric in @($ConversionMetrics)) {
            $index = [Array]::IndexOf($fields, $metric)
            if ($index -ge 0 -and $index -lt $sumRow.Count) {
                Add-MetricValue -Totals $totals -Metric $metric -Value $sumRow[$index] -ConversionMetrics $ConversionMetrics
            }
        }
        return $totals
    }

    $rows = @()
    if ($resultItems.Count -gt 1) {
        $rows = $resultItems
    } elseif ($Result.PSObject.Properties.Name -contains 'rows' -and $Result.rows) {
        $rows = @($Result.rows)
    } elseif ($Result.PSObject.Properties.Name -contains 'row' -and $Result.row) {
        $rows = @($Result.row)
    } elseif ($Result.PSObject.Properties.Name -contains 'data' -and $Result.data) {
        $rows = @($Result.data)
    } else {
        $rows = @($Result)
    }

    foreach ($row in $rows) {
        if ($row.PSObject.Properties.Name -contains 'kpis') {
            $kpis = @($row.kpis)
            for ($i = 0; $i -lt $PerformanceData.Count -and $i -lt $kpis.Count; $i++) {
                Add-MetricValue -Totals $totals -Metric ([string]$PerformanceData[$i]) -Value $kpis[$i] -ConversionMetrics $ConversionMetrics
            }
        } elseif ($row.PSObject.Properties.Name -contains 'kpi') {
            $kpi = $row.kpi
            $totals.impression += Get-Number -Value $kpi.impression
            $totals.click += Get-Number -Value $kpi.click
            $totals.cost += Get-Number -Value $kpi.cost
            $totals.conversion += Get-Number -Value $kpi.conversion
            foreach ($metric in @($ConversionMetrics)) {
                if ($kpi.PSObject.Properties.Name -contains $metric) {
                    Add-MetricValue -Totals $totals -Metric $metric -Value $kpi.$metric -ConversionMetrics $ConversionMetrics
                }
            }
        } else {
            foreach ($property in @($row.PSObject.Properties.Name)) {
                Add-MetricValue -Totals $totals -Metric $property -Value $row.$property -ConversionMetrics $ConversionMetrics
            }
        }
    }

    $totals
}

function Get-TongjiMetricTotal {
    param($Result, [string]$Metric)

    if (-not $Result) {
        return $null
    }

    if ($Result.PSObject.Properties.Name -contains 'fields' -and $Result.PSObject.Properties.Name -contains 'sum') {
        $fields = @($Result.fields)
        $sums = @($Result.sum)
        if ($sums.Count -gt 0) {
            $sumRow = @($sums[0])
            $index = [Array]::IndexOf($fields, $Metric)
            if ($index -ge 0 -and $index -lt $sumRow.Count) {
                return Get-Number -Value $sumRow[$index]
            }
        }
    }

    $total = 0
    $found = $false
    if ($Result.PSObject.Properties.Name -contains 'items' -and $Result.items) {
        foreach ($item in @($Result.items)) {
            if ($item.PSObject.Properties.Name -contains $Metric) {
                $total += Get-Number -Value $item.$Metric
                $found = $true
            }
        }
    }
    if ($Result.PSObject.Properties.Name -contains $Metric) {
        $total += Get-Number -Value $Result.$Metric
        $found = $true
    }

    if ($found) { return $total }
    $null
}

function Get-TongjiTargetConversions {
    param($Account, $Config, [string]$ReportDate)

    $siteId = Get-ConfigValue -Account $Account -Config $Config -Name 'tongjiSiteId' -DefaultValue ''
    if ([string]::IsNullOrWhiteSpace([string]$siteId)) {
        return $null
    }

    $endpoint = Get-ConfigValue -Account $Account -Config $Config -Name 'tongjiEndpoint' -DefaultValue 'https://api.baidu.com/json/tongji/v1/ReportService/getData'
    $metric = Get-ConfigValue -Account $Account -Config $Config -Name 'tongjiConversionMetric' -DefaultValue 'trans_count'
    $body = [ordered]@{
        site_id = [string]$siteId
        start_date = $ReportDate.Replace('-', '')
        end_date = $ReportDate.Replace('-', '')
        metrics = [string]$metric
        method = Get-ConfigValue -Account $Account -Config $Config -Name 'tongjiMethod' -DefaultValue 'pro/hour/a'
    }
    $adProduct = Get-ConfigValue -Account $Account -Config $Config -Name 'tongjiAdProduct' -DefaultValue '1,0'
    if (-not [string]::IsNullOrWhiteSpace([string]$adProduct)) {
        $body.adProduct = [string]$adProduct
    }

    $payload = [ordered]@{
        header = Get-TongjiHeader -Account $Account
        body = $body
    }

    $response = Invoke-BaiduJson -Uri $endpoint -Body $payload
    $failures = Get-BaiduFailures -Response $response
    if ($failures -and @($failures).Count -gt 0) {
        throw ('Baidu Tongji API failures: ' + ($failures | ConvertTo-Json -Compress))
    }

    $result = Get-ResultNode -Response $response
    Get-TongjiMetricTotal -Result $result -Metric ([string]$metric)
}

function Get-ReportRows {
    param($Account, $Config, [string]$ReportDate)

    if ([string]::IsNullOrWhiteSpace([string]$Account.accessToken)) {
        Write-Log ('Skip account without accessToken: ' + $Account.username)
        return $null
    }

    $endpoint = Get-ConfigValue -Account $Account -Config $Config -Name 'reportEndpoint' -DefaultValue 'https://api.baidu.com/json/sms/service/ReportService/getRealTimeData'
    $basePerformanceData = Get-ConfiguredStringArray -Account $Account -Config $Config -Name 'performanceData' -DefaultValue @('impression', 'click', 'cost')
    $conversionMetrics = Get-ConfiguredStringArray -Account $Account -Config $Config -Name 'conversionMetrics' -DefaultValue @('ocpcConversionsDetail10')
    $performanceData = Get-UniqueStrings -Values @($basePerformanceData, $conversionMetrics)

    $requestType = [ordered]@{
        performanceData = @($performanceData)
        startDate = $ReportDate + ' 00:00:00'
        endDate = $ReportDate + ' 23:59:59'
        levelOfDetails = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'levelOfDetails' -DefaultValue 2)
        reportType = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'reportType' -DefaultValue 2)
        statRange = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'statRange' -DefaultValue 2)
        unitOfTime = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'unitOfTime' -DefaultValue 5)
        device = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'device' -DefaultValue 0)
        platform = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'platform' -DefaultValue 0)
        number = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'number' -DefaultValue 1000)
        pageIndex = [int](Get-ConfigValue -Account $Account -Config $Config -Name 'pageIndex' -DefaultValue 1)
        order = [bool](Get-ConfigValue -Account $Account -Config $Config -Name 'order' -DefaultValue $true)
    }

    if ($Account.PSObject.Properties.Name -contains 'statIds' -and $Account.statIds) {
        $requestType.statIds = @($Account.statIds)
    }

    $response = Invoke-BaiduReport -Uri $endpoint -Header (Get-Header -Account $Account) -RequestType $requestType
    $failures = Get-BaiduFailures -Response $response
    if ($failures -and @($failures).Count -gt 0) {
        throw ('Baidu API failures: ' + ($failures | ConvertTo-Json -Compress))
    }
    if ($null -ne $response.header.status -and [int]$response.header.status -ne 0) {
        throw ('Baidu API status is not 0: ' + ($response.header | ConvertTo-Json -Compress))
    }

    $result = Get-ResultNode -Response $response
    if (-not $result) {
        throw 'Baidu API returned no report data.'
    }

    $totals = Convert-ResultToTotals -Result $result -PerformanceData $performanceData -ConversionMetrics $conversionMetrics
    $conversionSource = 'ReportService: ' + ($conversionMetrics -join '+')
    try {
        $tongjiConversions = Get-TongjiTargetConversions -Account $Account -Config $Config -ReportDate $ReportDate
        if ($null -ne $tongjiConversions) {
            $totals.conversion = $tongjiConversions
            $conversionSource = 'Tongji: trans_count'
        }
    } catch {
        Write-Log ('Baidu Tongji conversion failed for account ' + $Account.username + '; ' + $_.Exception.Message)
    }

    [pscustomobject]@{
        date = $ReportDate
        platform = 'baidu'
        account = [string]$Account.username
        device = 'all'
        impressions = [int][math]::Round([double]$totals.impression, 0)
        clicks = [int][math]::Round([double]$totals.click, 0)
        cost = [math]::Round([double]$totals.cost, 2)
        conversions = [int][math]::Round([double]$totals.conversion, 0)
        note = 'ReportService/getRealTimeData; conversion=' + $conversionSource
    }
}

function Save-ReportRows {
    param([object[]]$Rows, [string]$OutputCsv)

    if (-not $Rows -or $Rows.Count -eq 0) {
        Write-Log 'No rows to write.'
        return
    }

    $outputDir = Split-Path -Parent $OutputCsv
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
    $baseName = [IO.Path]::GetFileNameWithoutExtension($OutputCsv)
    $extension = [IO.Path]::GetExtension($OutputCsv)
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $datedOutputCsv = Join-Path $outputDir ($baseName + '_' + $timestamp + $extension)

    $Rows |
        Sort-Object date, platform, account, device |
        Export-Csv -Path $datedOutputCsv -NoTypeInformation -Encoding UTF8
    Write-Log ('Wrote report CSV: ' + $datedOutputCsv)

    $Rows |
        Sort-Object date, platform, account, device |
        Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Log ('Wrote latest report CSV: ' + $OutputCsv)
}

if (-not (Test-Path $ConfigPath)) {
    throw ('Config file not found: ' + $ConfigPath)
}

$config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$accounts = @($config.accounts | Where-Object { $_.enabled })
if (-not [string]::IsNullOrWhiteSpace($AccountUsername)) {
    $accounts = @($accounts | Where-Object { $_.username -eq $AccountUsername })
}

if ($accounts.Count -eq 0) {
    Write-Log 'No enabled accounts configured. Nothing to export.'
    Write-Log 'Done.'
    return
}

$reportDates = @(Get-ReportDateList -Date $Date -StartDate $StartDate -EndDate $EndDate -Days $Days)
$rows = @()
$failedAccounts = @()

foreach ($reportDate in $reportDates) {
    foreach ($account in $accounts) {
        try {
            Write-Log ('Start fetching Baidu data for account: ' + $account.username + '; date: ' + $reportDate)
            $result = Get-ReportRows -Account $account -Config $config -ReportDate $reportDate
            if ($null -eq $result) {
                $accountRows = @()
            } else {
                $accountRows = @($result)
            }
            if ($accountRows.Count -gt 0) {
                $rows += $accountRows
                Write-Log ('Baidu rows for account ' + $account.username + ' date ' + $reportDate + ': ' + $accountRows.Count)
            }
        } catch {
            $failureMessage = 'Baidu account failed: ' + $account.username + '; date: ' + $reportDate + '; ' + $_.Exception.Message
            $failedAccounts += $failureMessage
            Write-Log $failureMessage
        }
    }
}

Write-Log ('Baidu total rows: ' + $rows.Count)
$outputCsv = $config.outputCsv
if ($config.PSObject.Properties.Name -contains 'outputCsvs' -and
    $config.outputCsvs.PSObject.Properties.Name -contains 'baidu' -and
    -not [string]::IsNullOrWhiteSpace([string]$config.outputCsvs.baidu)) {
    $outputCsv = [string]$config.outputCsvs.baidu
}
Save-ReportRows -Rows $rows -OutputCsv $outputCsv
if ($failedAccounts.Count -gt 0) {
    throw ('Baidu report had account failures: ' + ($failedAccounts -join ' | '))
}
Write-Log 'Done.'
