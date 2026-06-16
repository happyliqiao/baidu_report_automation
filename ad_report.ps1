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
    $logPath = Join-Path $logDir '360_report.log'
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
        $fallbackPath = Join-Path $logDir ('360_report_' + $PID + '.log')
        Add-Content -Path $fallbackPath -Value $line -Encoding UTF8
    }
    Write-Host $line
}

function ConvertTo-FormBody {
    param([hashtable]$Data)
    $pairs = foreach ($key in $Data.Keys) {
        '{0}={1}' -f [uri]::EscapeDataString([string]$key), [uri]::EscapeDataString([string]$Data[$key])
    }
    $pairs -join '&'
}

function Get-Md5Hex {
    param([string]$Text)
    $bytes = [Text.Encoding]::UTF8.GetBytes($Text)
    $hash = [Security.Cryptography.MD5]::Create().ComputeHash($bytes)
    -join ($hash | ForEach-Object { $_.ToString('x2') })
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

function ConvertTo-QihooEncryptedPassword {
    param(
        [string]$Password,
        [string]$ApiSecret
    )

    if ($ApiSecret.Length -lt 32) {
        throw 'Qihoo360 apiSecret must be at least 32 characters.'
    }

    $md5 = Get-Md5Hex $Password
    $plainBytes = [Text.Encoding]::UTF8.GetBytes($md5)
    $keyBytes = [Text.Encoding]::UTF8.GetBytes($ApiSecret.Substring(0, 16))
    $ivBytes = [Text.Encoding]::UTF8.GetBytes($ApiSecret.Substring(16, 16))

    $aes = [Security.Cryptography.Aes]::Create()
    $aes.Mode = [Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [Security.Cryptography.PaddingMode]::None
    $aes.Key = $keyBytes
    $aes.IV = $ivBytes
    $encryptor = $aes.CreateEncryptor()
    $cipher = $encryptor.TransformFinalBlock($plainBytes, 0, $plainBytes.Length)
    -join ($cipher | ForEach-Object { $_.ToString('x2') })
}

function Invoke-FormPost {
    param(
        [string]$Uri,
        [hashtable]$Headers,
        [hashtable]$Body,
        [int]$MaxAttempts = 4,
        [int]$InitialDelaySeconds = 2
    )

    $form = ConvertTo-FormBody $Body

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return Invoke-RestMethod -Uri $Uri -Method Post -Headers $Headers -Body $form -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 60
        } catch {
            $detail = $_.Exception.Message
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
            Write-Log ('POST retry {0}/{1} in {2}s: {3}; {4}' -f ($attempt + 1), $MaxAttempts, $delay, $Uri, $detail)
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-QihooToken {
    param($Provider)
    if ([string]::IsNullOrWhiteSpace($Provider.apiKey) -or [string]::IsNullOrWhiteSpace($Provider.username)) {
            throw 'Qihoo360 config is missing apiKey or username.'
    }

    $passwd = $Provider.passwordEncrypted
    if ([string]::IsNullOrWhiteSpace($passwd)) {
        if ([string]::IsNullOrWhiteSpace($Provider.passwordPlain) -or [string]::IsNullOrWhiteSpace($Provider.apiSecret)) {
            throw 'Qihoo360 config is missing passwordEncrypted, or passwordPlain/apiSecret for local encryption.'
        }
        $passwd = ConvertTo-QihooEncryptedPassword -Password $Provider.passwordPlain -ApiSecret $Provider.apiSecret
    }

    $response = Invoke-FormPost `
        -Uri 'https://api.e.360.cn/uc/account/clientLogin' `
        -Headers @{ apiKey = $Provider.apiKey } `
        -Body @{ username = $Provider.username; passwd = $passwd }

    if (-not $response.accessToken) {
        throw ('Qihoo360 login did not return accessToken: ' + ($response | ConvertTo-Json -Compress))
    }
    $response.accessToken
}

function Get-QihooAccountRows {
    param($Provider, [string]$ReportDate, [string]$AccessToken)
    $rows = @()
    $deviceTypes = @('all')
    if ($Provider.deviceTypes -and $Provider.deviceTypes.Count -gt 0) {
        $deviceTypes = @($Provider.deviceTypes)
    }

    foreach ($type in $deviceTypes) {
        $body = @{
            startDate = $ReportDate
            endDate = $ReportDate
        }
        if ($type -and $type -ne 'all') {
            $body.type = $type
        }

        $response = Invoke-FormPost `
            -Uri 'https://api.e.360.cn/dianjing/report/accountDaily' `
            -Headers @{ apiKey = $Provider.apiKey; accessToken = $AccessToken } `
            -Body $body

        foreach ($item in @($response.dailyList)) {
            $rows += [pscustomobject]@{
                date = [string]$item.date
                platform = '360'
                account = [string]$Provider.username
                device = [string]$item.type
                impressions = [string]$item.views
                clicks = [string]$item.clicks
                cost = [string]$item.totalCost
                conversions = ''
                note = 'accountDaily'
            }
        }
    }
    $rows
}

function Get-QihooConversionRows {
    param($Provider, [string]$ReportDate, [string]$AccessToken)
    $totalConverts = 0
    $totalFactConverts = 0
    $page = 1

    while ($true) {
        $response = Invoke-FormPost `
            -Uri 'https://api.e.360.cn/dianjing/report/Ocpc' `
            -Headers @{ apiKey = $Provider.apiKey; accessToken = $AccessToken } `
            -Body @{
                startDate = $ReportDate
                endDate = $ReportDate
                groupDay = 'day'
                page = $page
            }

        $data = @($response.data)
        if ($data.Count -eq 0) { break }
        foreach ($item in $data) {
            $totalConverts += [int]([decimal]('0' + [string]$item.converts))
            $totalFactConverts += [int]([decimal]('0' + [string]$item.factConverts))
        }
        if ($data.Count -lt 1000) { break }
        $page += 1
    }

    [pscustomobject]@{
        Converts = $totalConverts
        FactConverts = $totalFactConverts
    }
}

function Merge-ConversionsIntoRows {
    param([object[]]$Rows, $Conversions)
    if (-not $Rows -or $Rows.Count -eq 0) { return $Rows }
    $first = $true
    foreach ($row in $Rows) {
        if ($first) {
            $row.conversions = [string]$Conversions.Converts
            $row.note = $row.note + '; ocpc converts total, fact=' + [string]$Conversions.FactConverts
            $first = $false
        }
    }
    $Rows
}

function Get-QihooRows {
    param($Provider, [string]$ReportDate)

    $accounts = @()
    if ($Provider.accounts) {
        $accounts = @($Provider.accounts | Where-Object { $_.enabled })
    } else {
        $accounts = @($Provider)
    }

    $allRows = @()
    foreach ($account in $accounts) {
        try {
            Write-Log ('Start fetching Qihoo360 data for account: ' + $account.username)
            $effective = [pscustomobject]@{
                enabled = $true
                apiKey = if ($account.apiKey) { $account.apiKey } else { $Provider.apiKey }
                apiSecret = if ($account.apiSecret) { $account.apiSecret } else { $Provider.apiSecret }
                username = $account.username
                passwordPlain = $account.passwordPlain
                passwordEncrypted = $account.passwordEncrypted
                passwordEncryption = if ($account.passwordEncryption) { $account.passwordEncryption } else { $Provider.passwordEncryption }
                deviceTypes = if ($account.deviceTypes) { $account.deviceTypes } else { $Provider.deviceTypes }
                includeOcpcConversions = if ($null -ne $account.includeOcpcConversions) { $account.includeOcpcConversions } else { $Provider.includeOcpcConversions }
            }

            $token = Get-QihooToken -Provider $effective
            $rows = @(Get-QihooAccountRows -Provider $effective -ReportDate $ReportDate -AccessToken $token)
            if ($effective.includeOcpcConversions) {
                $conversions = Get-QihooConversionRows -Provider $effective -ReportDate $ReportDate -AccessToken $token
                $rows = @(Merge-ConversionsIntoRows -Rows $rows -Conversions $conversions)
            }
            Write-Log ('Qihoo360 rows for account ' + $effective.username + ': ' + $rows.Count)
            $allRows += $rows
        } catch {
            Write-Log ('Qihoo360 account failed: ' + $account.username + '; ' + $_.Exception.Message)
        }
    }

    Write-Log ('Qihoo360 total rows: ' + $allRows.Count)
    $allRows
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
        Sort-Object date,platform,account,device |
        Export-Csv -Path $datedOutputCsv -NoTypeInformation -Encoding UTF8
    Write-Log ('Wrote report CSV: ' + $datedOutputCsv)

    $Rows |
        Sort-Object date,platform,account,device |
        Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Log ('Wrote latest report CSV: ' + $OutputCsv)
}

if (-not (Test-Path $ConfigPath)) {
    throw ('Config file not found: ' + $ConfigPath)
}

$config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$allRows = @()
$reportDates = @(Get-ReportDateList -Date $Date -StartDate $StartDate -EndDate $EndDate -Days $Days)

foreach ($reportDate in $reportDates) {
    if ($config.providers.qihoo360.enabled) {
        $provider = $config.providers.qihoo360
        if (-not [string]::IsNullOrWhiteSpace($AccountUsername)) {
            $matchedAccounts = @($provider.accounts | Where-Object { $_.enabled -and ([string]$_.username).Trim() -eq $AccountUsername.Trim() })
            $provider = [pscustomobject]@{
                enabled = $provider.enabled
                apiKey = $provider.apiKey
                apiSecret = $provider.apiSecret
                passwordEncryption = $provider.passwordEncryption
                deviceTypes = $provider.deviceTypes
                includeOcpcConversions = $provider.includeOcpcConversions
                accounts = $matchedAccounts
            }
        }
        $allRows += @(Get-QihooRows -Provider $provider -ReportDate $reportDate)
    }
}

$outputCsv = $config.outputCsv
if ($config.PSObject.Properties.Name -contains 'outputCsvs' -and
    $config.outputCsvs.PSObject.Properties.Name -contains 'qihoo360' -and
    -not [string]::IsNullOrWhiteSpace([string]$config.outputCsvs.qihoo360)) {
    $outputCsv = [string]$config.outputCsvs.qihoo360
}

Save-ReportRows -Rows $allRows -OutputCsv $outputCsv
Write-Log 'Done.'
