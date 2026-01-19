$ErrorActionPreference = "Stop"

while (-not $Global:MicrosoftEntraIDAccessToken) {
    $Global:MicrosoftEntraIDAccessToken = Read-Host "Paste a Microsoft Graph access token"
}

$GraphBaseUri = "https://graph.microsoft.com/v1.0"
$GraphBetaBaseUri = "https://graph.microsoft.com/beta"

function Request-GraphAccessToken {
    $newToken = Read-Host "Graph access token expired. Paste a new Microsoft Graph access token"
    if (-not $newToken) {
        throw "Token refresh aborted: no token provided."
    }

    $Global:MicrosoftEntraIDAccessToken = $newToken
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory)]
        [ValidateSet("GET", "POST", "PATCH", "DELETE")]
        [string]$Method,
        [Parameter(Mandatory)]
        [string]$Uri
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $headers = @{
        Authorization = "Bearer $Global:MicrosoftEntraIDAccessToken"
    }

    $attempt = 0
    $maxAttempts = 5
    $tokenRefreshed = $false
    while ($true) {
        try {
            return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -TimeoutSec 120
        }
        catch {
            $attempt++
            $response = $_.Exception.Response
            $statusCode = $null
            if ($response) {
                $statusCode = [int]$response.StatusCode
            }

            $errorMessage = $null
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                $errorMessage = $_.ErrorDetails.Message
            }
            $isTokenExpired = $statusCode -eq 401 -and ($errorMessage -match "expired|InvalidAuthenticationToken|token" -or $_.Exception.Message -match "expired|InvalidAuthenticationToken|token")

            if ($isTokenExpired -and -not $tokenRefreshed) {
                Request-GraphAccessToken
                $headers.Authorization = "Bearer $Global:MicrosoftEntraIDAccessToken"
                $tokenRefreshed = $true
                continue
            }

            $isTransientHttp = $statusCode -in 408, 429, 500, 502, 503, 504
            $isConnectionReset = $_.Exception.Message -match "forcibly closed|transport connection"

            if (($isTransientHttp -or $isConnectionReset) -and $attempt -lt $maxAttempts) {
                $retryAfter = $response.Headers["Retry-After"]
                if (-not $retryAfter) {
                    $retryAfter = [Math]::Min(30, [Math]::Pow(2, $attempt))
                }
                Start-Sleep -Seconds ([int]$retryAfter)
                continue
            }
            throw
        }
    }
}

function Get-GraphPaged {
    param(
        [Parameter(Mandatory)]
        [string]$Uri
    )

    $items = @()
    $nextLink = $Uri
    while ($nextLink) {
        $response = Invoke-GraphRequest -Method GET -Uri $nextLink
        if ($response.value) {
            $items += $response.value
        }
        $nextLink = $response.'@odata.nextLink'
    }

    return $items
}

$managedDevicesUri = "$GraphBaseUri/deviceManagement/managedDevices?`$filter=operatingSystem eq 'Windows'&`$select=id,deviceName,complianceState,azureADDeviceId,operatingSystem"
$managedDevices = Get-GraphPaged -Uri $managedDevicesUri

$bitlockerByDeviceId = @{}
foreach ($device in $managedDevices) {
    if (-not $device.azureADDeviceId) {
        continue
    }

    Write-Host ("[BitLocker] Processing {0} ({1})" -f $device.deviceName, $device.azureADDeviceId)
    $escapedDeviceId = $device.azureADDeviceId.Replace("'", "''")
    $bitlockerKeysUri = "$GraphBetaBaseUri/informationProtection/bitlocker/recoveryKeys?`$filter=deviceId eq '$escapedDeviceId'&`$select=id,deviceId,createdDateTime,volumeType"
    $bitlockerKeys = Get-GraphPaged -Uri $bitlockerKeysUri
    if (-not $bitlockerKeys) {
        Write-Host ("[BitLocker] No recovery keys found for {0}" -f $device.azureADDeviceId)
        continue
    }

    $latestKey = $bitlockerKeys | Sort-Object -Property createdDateTime -Descending | Select-Object -First 1
    if ($latestKey) {
        $bitlockerByDeviceId[$device.azureADDeviceId] = [pscustomobject]@{
            Status          = "Present"
            BackupTimestamp = $latestKey.createdDateTime
        }
        Write-Host ("[BitLocker] Latest backup {0}" -f $latestKey.createdDateTime)
    }
}

$lapsByDeviceId = @{}
foreach ($device in $managedDevices) {
    if (-not $device.azureADDeviceId) {
        continue
    }

    Write-Host ("[LAPS] Processing {0} ({1})" -f $device.deviceName, $device.azureADDeviceId)
    $escapedDeviceId = $device.azureADDeviceId.Replace("'", "''")
    $lapsUri = "$GraphBaseUri/directory/deviceLocalCredentials/$escapedDeviceId"
    $lapsEntry = Invoke-GraphRequest -Method GET -Uri $lapsUri
    if (-not $lapsEntry) {
        Write-Host ("[LAPS] No credentials found for {0}" -f $device.azureADDeviceId)
        continue
    }

    $lapsByDeviceId[$device.azureADDeviceId] = $lapsEntry.lastPasswordRotationDateTime
    Write-Host ("[LAPS] Last rotation {0}" -f $lapsEntry.lastPasswordRotationDateTime)
}

$Results = foreach ($device in $managedDevices) {
    Write-Host ("[Result] Building output for {0} ({1})" -f $device.deviceName, $device.azureADDeviceId)
    $aadDeviceId = $device.azureADDeviceId
    $bitlockerInfo = $null
    if ($aadDeviceId -and $bitlockerByDeviceId.ContainsKey($aadDeviceId)) {
        $bitlockerInfo = $bitlockerByDeviceId[$aadDeviceId]
    }

    $lapsRotation = $null
    if ($aadDeviceId -and $lapsByDeviceId.ContainsKey($aadDeviceId)) {
        $lapsRotation = $lapsByDeviceId[$aadDeviceId]
    }

    [pscustomobject]@{
        DeviceName                              = $device.deviceName
        IntuneDeviceId                          = $device.id
        ComplianceStatus                        = $device.complianceState
        BitlockerRecoveryKeyStatus              = if ($bitlockerInfo) { $bitlockerInfo.Status } else { "Missing" }
        BitlockerRecoveryKeyBackupTimestamp     = if ($bitlockerInfo) { $bitlockerInfo.BackupTimestamp } else { $null }
        LocalAdminPasswordLastRotationTimestamp = $lapsRotation
    }
}

$Results
