param(
    [string]$ConfigPath = (Join-Path $PSScriptRoot 'config.json'),
    [string]$AccountUsername,
    [string]$Code,
    [string]$RedirectUri,
    [switch]$Interactive,
    [switch]$PrintAuthUrl,
    [switch]$UpdateConfig,
    [switch]$SkipMissingRefreshToken
)

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

function Get-TargetAccounts {
    param($Config, [string]$Username)

    $accounts = @($Config.accounts | Where-Object { $_.enabled })
    if (-not [string]::IsNullOrWhiteSpace($Username)) {
        $accounts = @($accounts | Where-Object { $_.username -eq $Username })
    }

    if ($accounts.Count -eq 0) {
        throw ('No enabled account matched username: ' + $Username)
    }

    $accounts
}

function Assert-OAuthFields {
    param($Account)

    if ([string]::IsNullOrWhiteSpace((Get-ClientId -Account $Account))) {
        throw ('Account ' + $Account.username + ' is missing apiKey/clientId/appKey/appId.')
    }
    if ([string]::IsNullOrWhiteSpace([string]$Account.secretKey)) {
        throw ('Account ' + $Account.username + ' is missing secretKey.')
    }
}

function Get-ClientId {
    param($Account)

    foreach ($name in @('apiKey', 'clientId', 'appKey', 'appId')) {
        if ($Account.PSObject.Properties.Name -contains $name -and -not [string]::IsNullOrWhiteSpace([string]$Account.$name)) {
            return [string]$Account.$name
        }
    }
    ''
}

function Add-NotePropertyIfMissing {
    param(
        $Object,
        [string]$Name,
        [object]$Value
    )

    if (-not ($Object.PSObject.Properties.Name -contains $Name)) {
        $Object | Add-Member -NotePropertyName $Name -NotePropertyValue $Value
    }
}

function Get-QueryString {
    param([hashtable]$Query)

    $pairs = foreach ($key in $Query.Keys) {
        '{0}={1}' -f [uri]::EscapeDataString($key), [uri]::EscapeDataString([string]$Query[$key])
    }
    $pairs -join '&'
}

function Get-AuthUrl {
    param(
        $Account,
        [string]$ClientId,
        [string]$RedirectUri
    )

    if ($Account.PSObject.Properties.Name -contains 'authorizationUrl' -and
        -not [string]::IsNullOrWhiteSpace([string]$Account.authorizationUrl)) {
        return [string]$Account.authorizationUrl
    }
    if ($Account.PSObject.Properties.Name -contains 'authUrl' -and
        -not [string]::IsNullOrWhiteSpace([string]$Account.authUrl)) {
        return [string]$Account.authUrl
    }

    'https://openapi.baidu.com/oauth/2.0/authorize?response_type=code&client_id={0}&redirect_uri={1}&scope=basic&display=popup' -f [uri]::EscapeDataString($ClientId), [uri]::EscapeDataString($RedirectUri)
}

function Get-CallbackCode {
    param($Request)

    foreach ($name in @('authCode', 'auth_code', 'code', 'authorization_code', 'oauthCode')) {
        $value = $Request.QueryString[$name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }
    ''
}

function Get-CallbackUserId {
    param($Request)

    foreach ($name in @('userId', 'user_id', 'ucid', 'uid', 'masterUid')) {
        $value = $Request.QueryString[$name]
        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }
    }
    ''
}

function Get-QueryDebugText {
    param($Request)

    $parts = @()
    foreach ($key in $Request.QueryString.AllKeys) {
        if ([string]::IsNullOrWhiteSpace($key)) {
            continue
        }
        $parts += ($key + '=' + $Request.QueryString[$key])
    }
    if ($parts.Count -eq 0) {
        return 'no query parameters'
    }
    $parts -join '; '
}

function Invoke-MarketingTokenRequest {
    param(
        [string]$Uri,
        [hashtable]$Body,
        [int]$MaxAttempts = 4,
        [int]$InitialDelaySeconds = 2
    )

    $json = $Body | ConvertTo-Json -Depth 20 -Compress

    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            return Invoke-RestMethod -Uri $Uri -Method Post -Body $json -ContentType 'application/json' -TimeoutSec 60
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
            Write-Warning ('Token request retry {0}/{1} in {2}s: {3}' -f ($attempt + 1), $MaxAttempts, $delay, $detail)
            Start-Sleep -Seconds $delay
        }
    }
}

function Get-TokenField {
    param($Token, [string[]]$Names)

    foreach ($name in $Names) {
        if ($Token.PSObject.Properties.Name -contains $name -and -not [string]::IsNullOrWhiteSpace([string]$Token.$name)) {
            return $Token.$name
        }
    }

    if ($Token.PSObject.Properties.Name -contains 'data' -and $Token.data) {
        foreach ($name in $Names) {
            if ($Token.data.PSObject.Properties.Name -contains $name -and -not [string]::IsNullOrWhiteSpace([string]$Token.data.$name)) {
                return $Token.data.$name
            }
        }
    }

    ''
}

function Get-EffectiveRedirectUri {
    param($Account, [string]$ExplicitRedirectUri)

    if (-not [string]::IsNullOrWhiteSpace($ExplicitRedirectUri)) {
        return $ExplicitRedirectUri
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Account.redirectUri)) {
        return [string]$Account.redirectUri
    }
    'http://127.0.0.1:8787/baidu_oauth_callback/'
}

function Read-AuthorizationCode {
    param($Account, [string]$ExplicitCode)

    if (-not [string]::IsNullOrWhiteSpace($ExplicitCode)) {
        return $ExplicitCode
    }
    if (-not [string]::IsNullOrWhiteSpace([string]$Account.authorizationCode)) {
        return [string]$Account.authorizationCode
    }
    ''
}

if (-not (Test-Path $ConfigPath)) {
    throw ('Config file not found: ' + $ConfigPath)
}

$config = Get-Content -Path $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
$accounts = @(Get-TargetAccounts -Config $config -Username $AccountUsername)

if ($accounts.Count -eq 0) {
    Write-Host 'No enabled accounts. Nothing to refresh.'
    return
}

if ($PrintAuthUrl) {
    foreach ($account in $accounts) {
        Assert-OAuthFields -Account $account
        $accountRedirectUri = Get-EffectiveRedirectUri -Account $account -ExplicitRedirectUri $RedirectUri
        Write-Host ('Account: ' + $account.username)
        Write-Host ('RedirectUri: ' + $accountRedirectUri)
        Write-Host (Get-AuthUrl -Account $account -ClientId (Get-ClientId -Account $account) -RedirectUri $accountRedirectUri)
        Write-Host ''
    }
    return
}

foreach ($account in $accounts) {
    try {
        Assert-OAuthFields -Account $account
        $accountRedirectUri = Get-EffectiveRedirectUri -Account $account -ExplicitRedirectUri $RedirectUri
        $authorizationCode = Read-AuthorizationCode -Account $account -ExplicitCode $Code

        if ($Interactive) {
            $listener = New-Object System.Net.HttpListener
            $listener.Prefixes.Add($accountRedirectUri)
            try {
                $listener.Start()
            } catch {
                throw ('Failed to start local callback listener on ' + $accountRedirectUri + '. Use -Code instead.')
            }

            $authorizeUrl = Get-AuthUrl -Account $account -ClientId (Get-ClientId -Account $account) -RedirectUri $accountRedirectUri
            Start-Process $authorizeUrl
            Write-Host ('Waiting for authorization code for account: ' + $account.username)

            $async = $listener.BeginGetContext($null, $null)
            if (-not $async.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds(300))) {
                $listener.Stop()
                throw 'Timed out waiting for OAuth callback.'
            }

            $context = $listener.EndGetContext($async)
            $request = $context.Request
            $response = $context.Response
            $authorizationCode = Get-CallbackCode -Request $request
            $callbackUserId = Get-CallbackUserId -Request $request
            $errorCode = $request.QueryString['error']
            $errorDescription = $request.QueryString['error_description']
            $debugText = Get-QueryDebugText -Request $request

            if ([string]::IsNullOrWhiteSpace($authorizationCode)) {
                $message = 'Authorization did not return code. Callback query: ' + $debugText
                if ($errorCode) { $message += ' Error: ' + $errorCode }
                if ($errorDescription) { $message += ' ' + $errorDescription }
                $bytes = [Text.Encoding]::UTF8.GetBytes($message)
                $response.StatusCode = 400
                $response.OutputStream.Write($bytes, 0, $bytes.Length)
                $response.Close()
                $listener.Stop()
                throw $message
            }

            Add-NotePropertyIfMissing -Object $account -Name 'userId' -Value ''
            if (-not [string]::IsNullOrWhiteSpace($callbackUserId)) {
                $account.userId = [string]$callbackUserId
            }

            $successHtml = '<html><head><meta charset="utf-8"></head><body>Authorization received. You can close this window.</body></html>'
            $successBytes = [Text.Encoding]::UTF8.GetBytes($successHtml)
            $response.ContentType = 'text/html; charset=utf-8'
            $response.OutputStream.Write($successBytes, 0, $successBytes.Length)
            $response.Close()
            $listener.Stop()
        }

        if (-not [string]::IsNullOrWhiteSpace($authorizationCode)) {
            $token = Invoke-MarketingTokenRequest -Uri 'https://u.baidu.com/oauth/accessToken' -Body @{
                appId = [string]$account.appId
                authCode = $authorizationCode
                secretKey = [string]$account.secretKey
                grantType = 'access_token'
                userId = if ([string]::IsNullOrWhiteSpace([string]$account.userId)) { $null } else { [long]$account.userId }
            }
        } elseif (-not [string]::IsNullOrWhiteSpace([string]$account.refreshToken)) {
            $token = Invoke-MarketingTokenRequest -Uri 'https://u.baidu.com/oauth/refreshToken' -Body @{
                appId = [string]$account.appId
                refreshToken = [string]$account.refreshToken
                secretKey = [string]$account.secretKey
                userId = if ([string]::IsNullOrWhiteSpace([string]$account.userId)) { $null } else { [long]$account.userId }
            }
        } elseif ($SkipMissingRefreshToken) {
            Write-Host ('Skip account without refreshToken: ' + $account.username)
            continue
        } else {
            throw ('Account ' + $account.username + ' has no refreshToken. Run with -PrintAuthUrl first, authorize in browser, then rerun with -Code CODE, or use -SkipMissingRefreshToken.')
        }

        $accessToken = Get-TokenField -Token $token -Names @('accessToken', 'access_token')
        if (-not $accessToken) {
            throw ('Token response did not include access_token: ' + ($token | ConvertTo-Json -Compress))
        }

        Add-NotePropertyIfMissing -Object $account -Name 'userId' -Value ''
        Add-NotePropertyIfMissing -Object $account -Name 'refreshToken' -Value ''
        Add-NotePropertyIfMissing -Object $account -Name 'tokenExpiresAt' -Value ''
        Add-NotePropertyIfMissing -Object $account -Name 'authorizationCode' -Value ''

        $account.accessToken = [string]$accessToken
        $account.authorizationCode = ''

        $refreshToken = Get-TokenField -Token $token -Names @('refreshToken', 'refresh_token')
        if ($refreshToken) {
            $account.refreshToken = [string]$refreshToken
        }

        $userId = Get-TokenField -Token $token -Names @('userId', 'user_id')
        if ($userId) {
            $account.userId = [string]$userId
        }

        $expiresIn = Get-TokenField -Token $token -Names @('expiresIn', 'expires_in')
        $expiresTime = Get-TokenField -Token $token -Names @('expiresTime', 'expires_time')
        if ($expiresIn) {
            $account.tokenExpiresAt = (Get-Date).AddSeconds([int]$expiresIn).ToString('yyyy-MM-dd HH:mm:ss')
        } elseif ($expiresTime) {
            $account.tokenExpiresAt = [string]$expiresTime
        }

        Write-Host ('Refreshed accessToken for account: ' + $account.username)
    } catch {
        Write-Warning ($_.Exception.Message)
    }
}

if ($UpdateConfig) {
    $config | ConvertTo-Json -Depth 20 | Set-Content -Path $ConfigPath -Encoding UTF8
    Write-Host ('Updated config: ' + $ConfigPath)
} else {
    Write-Host 'Dry run only. Add -UpdateConfig to save token fields into config.json.'
}
