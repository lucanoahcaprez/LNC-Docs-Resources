<#
.SYNOPSIS
Sync mailbox contacts to a shared mailbox "All Contacts" folder and back using Microsoft Graph.

.DESCRIPTION
Step 1: Sync contacts from each licensed mailbox user to the shared mailbox "All Contacts" folder.
Step 2: Sync the shared mailbox "All Contacts" folder back to each user's mailbox contacts.
Adds a SyncId and CreatedBy stamp in personalNotes to match contacts across runs.
Requires application permissions and admin consent.

.EXAMPLE
.\automation.ps1
#>
# Load configuration from GitLab environment variables (preferred)
$rawConfigJson = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_CONFIG_JSON')
if ($rawConfigJson -and $rawConfigJson.Trim().Length -gt 0) {
    try {
        $ConfigObject = $rawConfigJson | ConvertFrom-Json
        Write-Host "Loaded configuration from env var AUTOMATION_SYNCCONTACTS_CONFIG_JSON" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to parse AUTOMATION_SYNCCONTACTS_CONFIG_JSON: $_"
        exit 1
    }
}
else {
    $ConfigObject = [PSCustomObject]@{
        TenantId                = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_TENANT_ID')
        ClientId                = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_CLIENT_ID')
        ClientSecret            = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_CLIENT_SECRET')
        GlobalAddressBookUserId = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_GLOBAL_ADDRESS_BOOK_USER_ID')
        DryRun                  = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_DRY_RUN')
    }

    Write-Host "Loaded configuration from individual environment variables" -ForegroundColor Green
}

$TenantId = $ConfigObject.TenantId
$ClientId = $ConfigObject.ClientId
$ClientSecret = $ConfigObject.ClientSecret
$GlobalAddressBookUserId = $ConfigObject.GlobalAddressBookUserId
$PageSize = 100
$UpdateExisting = $false
function ConvertTo-Boolean {
    param([object]$Value, [bool]$Default = $false)

    if ($null -eq $Value) { return $Default }
    if ($Value -is [bool]) { return $Value }

    $text = $Value.ToString().Trim().ToLowerInvariant()
    if ($text -in @("true", "1", "yes", "y")) { return $true }
    if ($text -in @("false", "0", "no", "n")) { return $false }

    return $Default
}

$DryRun = ConvertTo-Boolean -Value $ConfigObject.DryRun -Default $false

$TestUserList = @(
    "admin_luca1@yybpm.onmicrosoft.com",
    "luca_test1@yybpm.onmicrosoft.com"
)

function Get-GraphAccessToken {
    param(
        [Parameter(Mandatory = $true)][string]$TenantId,
        [Parameter(Mandatory = $true)][string]$ClientId,
        [Parameter(Mandatory = $true)][string]$ClientSecret
    )

    $tokenUri = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
    $body = @{ 
        grant_type    = "client_credentials"
        scope         = "https://graph.microsoft.com/.default"
        client_id     = $ClientId
        client_secret = $ClientSecret
    }

    $tokenResponse = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body
    return $tokenResponse.access_token
}

function Invoke-GraphRequest {
    param(
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [hashtable]$Body,
        [hashtable]$AdditionalHeaders
    )

    $headers = @{ Authorization = "Bearer $AccessToken" }
    if ($AdditionalHeaders) {
        foreach ($key in $AdditionalHeaders.Keys) {
            $headers[$key] = $AdditionalHeaders[$key]
        }
    }
    if ($Body) {
        $json = $Body | ConvertTo-Json -Depth 6
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType "application/json" -Body $json
    }

    return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers
}

function Get-GraphPaged {
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [hashtable]$AdditionalHeaders
    )

    $results = @()
    $next = $Uri

    while ($next) {
        $response = Invoke-GraphRequest -Method "GET" -Uri $next -AccessToken $AccessToken -AdditionalHeaders $AdditionalHeaders
        if ($response.value) {
            $results += $response.value
        }
        $next = $response."@odata.nextLink"
    }

    return $results
}

function Get-UserContactFolders {
    param(
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )

    $rootUri = "https://graph.microsoft.com/v1.0/users/$UserId/contactFolders?`$top=200&`$select=id,displayName"
    $rootFolders = Get-GraphPaged -Uri $rootUri -AccessToken $AccessToken

    $allFolders = @()
    $queue = New-Object System.Collections.Generic.Queue[object]
    foreach ($folder in $rootFolders) {
        $allFolders += $folder
        $queue.Enqueue($folder)
    }

    while ($queue.Count -gt 0) {
        $current = $queue.Dequeue()
        $childUri = "https://graph.microsoft.com/v1.0/users/$UserId/contactFolders/$($current.id)/childFolders?`$top=200&`$select=id,displayName"
        $children = Get-GraphPaged -Uri $childUri -AccessToken $AccessToken
        foreach ($child in $children) {
            $allFolders += $child
            $queue.Enqueue($child)
        }
    }

    return $allFolders
}

function Get-UserContacts {
    param(
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $true)][int]$PageSize,
        [Parameter(Mandatory = $true)][string]$Select,
        [string]$FolderId,
        [switch]$AllFolders
    )

    if ($FolderId) {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contactFolders/$FolderId/contacts?`$top=$PageSize&`$select=$Select"
        return Get-GraphPaged -Uri $uri -AccessToken $AccessToken
    }

    if (-not $AllFolders) {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contacts?`$top=$PageSize&`$select=$Select"
        return Get-GraphPaged -Uri $uri -AccessToken $AccessToken
    }

    $contacts = @()
    $folders = Get-UserContactFolders -UserId $UserId -AccessToken $AccessToken
    foreach ($folder in $folders) {
        $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contactFolders/$($folder.id)/contacts?`$top=$PageSize&`$select=$Select"
        $contacts += Get-GraphPaged -Uri $uri -AccessToken $AccessToken
    }

    return $contacts
}

function Get-ContactFolderByName {
    param(
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $true)][string]$FolderName
    )

    $folders = Get-UserContactFolders -UserId $UserId -AccessToken $AccessToken
    foreach ($folder in $folders) {
        if ($folder.displayName -eq $FolderName) {
            return $folder
        }
    }

    return $null
}

function Ensure-ContactFolder {
    param(
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $true)][string]$FolderName
    )

    $existing = Get-ContactFolderByName -UserId $UserId -AccessToken $AccessToken -FolderName $FolderName
    if ($existing) { return $existing }

    $uri = "https://graph.microsoft.com/v1.0/users/$UserId/contactFolders"
    $payload = @{ displayName = $FolderName }
    return Invoke-GraphRequest -Method "POST" -Uri $uri -AccessToken $AccessToken -Body $payload
}

function Get-LicensedUsers {
    param(
        [Parameter(Mandatory = $true)][string]$AccessToken
    )

    $headers = @{ "ConsistencyLevel" = "eventual" }
    $uri = "https://graph.microsoft.com/v1.0/users?`$select=id,userPrincipalName&`$filter=assignedLicenses/any(x:x/skuId ne null)&`$top=999"
    $response = Get-GraphPaged -Uri $uri -AccessToken $AccessToken -AdditionalHeaders $headers

    return $response
}

function Normalize-Email {
    param([string]$Email)
    if ([string]::IsNullOrWhiteSpace($Email)) { return $null }
    return $Email.Trim().ToLowerInvariant()
}

function Get-PrimaryEmail {
    param($Contact)
    if ($null -eq $Contact.emailAddresses) { return $null }
    foreach ($entry in $Contact.emailAddresses) {
        if ($entry.address) {
            return (Normalize-Email -Email $entry.address)
        }
    }
    return $null
}

function Add-IfValue {
    param(
        [hashtable]$Target,
        [string]$Key,
        $Value
    )

    if ($null -eq $Value) { return }
    if ($Value -is [string] -and [string]::IsNullOrWhiteSpace($Value)) { return }
    if ($Value -is [System.Collections.IEnumerable] -and -not ($Value -is [string])) {
        if ($Value.Count -eq 0) { return }
    }

    $Target[$Key] = $Value
}

function Get-SyncIdFromNotes {
    param([string]$Notes)

    if ([string]::IsNullOrWhiteSpace($Notes)) { return $null }
    $match = [regex]::Match($Notes, '(?i)SyncId\s*=\s*([0-9a-f-]{36})')
    if ($match.Success) { return $match.Groups[1].Value }
    return $null
}

function Get-CreatedByFromNotes {
    param([string]$Notes)

    if ([string]::IsNullOrWhiteSpace($Notes)) { return $null }
    $match = [regex]::Match($Notes, '(?i)CreatedBy\s*=\s*([^;]+)')
    if ($match.Success) { return $match.Groups[1].Value.Trim() }
    return $null
}

function Build-SyncNotes {
    param(
        [string]$ExistingNotes,
        [string]$SyncId,
        [string]$CreatedBy
    )

    $cleanedLines = @()
    if (-not [string]::IsNullOrWhiteSpace($ExistingNotes)) {
        foreach ($line in ($ExistingNotes -split "`r?`n")) {
            if ($line -match '(?i)SyncId\s*=' -or $line -match '(?i)CreatedBy\s*=') { continue }
            if (-not [string]::IsNullOrWhiteSpace($line)) { $cleanedLines += $line }
        }
    }

    $tagLine = "SyncId=$SyncId;CreatedBy=$CreatedBy"
    if ($cleanedLines.Count -eq 0) { return $tagLine }

    return ($cleanedLines -join "`r`n") + "`r`n" + $tagLine
}

function Sync-Contacts {
    param(
        [Parameter(Mandatory = $true)][string]$SourceUserId,
        [Parameter(Mandatory = $true)][string]$TargetUserId,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $true)][int]$PageSize,
        [string]$SourceFolderId,
        [string]$TargetFolderId,
        [switch]$SourceAllFolders,
        [switch]$TargetAllFolders,
        [switch]$UpdateExisting,
        [switch]$DryRun
    )

    $sourceSelect = "id,displayName,givenName,surname,companyName,jobTitle,department,businessPhones,mobilePhone,homePhones,emailAddresses,imAddresses,personalNotes,categories"
    $targetSelect = "id,displayName,emailAddresses,personalNotes"

    $sourceContacts = Get-UserContacts -UserId $SourceUserId -AccessToken $AccessToken -PageSize $PageSize -Select $sourceSelect -FolderId $SourceFolderId -AllFolders:$SourceAllFolders
    $targetContacts = Get-UserContacts -UserId $TargetUserId -AccessToken $AccessToken -PageSize $PageSize -Select $targetSelect -FolderId $TargetFolderId -AllFolders:$TargetAllFolders

    $targetBySyncId = @{}
    $targetByEmail = @{}
    $targetByDisplayName = @{}

    foreach ($contact in $targetContacts) {
        $syncId = Get-SyncIdFromNotes -Notes $contact.personalNotes
        if ($syncId -and -not $targetBySyncId.ContainsKey($syncId)) {
            $targetBySyncId[$syncId] = $contact
        }

        $email = Get-PrimaryEmail -Contact $contact
        if ($email) {
            $targetByEmail[$email] = $contact
        } elseif ($contact.displayName -and -not $targetByDisplayName.ContainsKey($contact.displayName)) {
            $targetByDisplayName[$contact.displayName] = $contact
        }
    }

    $created = 0
    $updated = 0
    $skipped = 0

    foreach ($contact in $sourceContacts) {
        $email = Get-PrimaryEmail -Contact $contact
        $displayName = $contact.displayName
        $syncId = Get-SyncIdFromNotes -Notes $contact.personalNotes

        $existing = $null
        if ($syncId -and $targetBySyncId.ContainsKey($syncId)) {
            $existing = $targetBySyncId[$syncId]
        } elseif ($email -and $targetByEmail.ContainsKey($email)) {
            $existing = $targetByEmail[$email]
        } elseif (-not $email -and $displayName -and $targetByDisplayName.ContainsKey($displayName)) {
            $existing = $targetByDisplayName[$displayName]
        }

        if (-not $syncId -and $existing) {
            $syncId = Get-SyncIdFromNotes -Notes $existing.personalNotes
        }
        if (-not $syncId) {
            $syncId = [guid]::NewGuid().ToString()
        }

        $createdByValue = Get-CreatedByFromNotes -Notes $contact.personalNotes
        if (-not $createdByValue) {
            $createdByValue = $SourceUserId
        }

        $notes = Build-SyncNotes -ExistingNotes $contact.personalNotes -SyncId $syncId -CreatedBy $createdByValue
        $notesChanged = $false
        if ($existing) {
            $notesChanged = ($notes -ne $existing.personalNotes)
        } else {
            $notesChanged = $true
        }

        $payload = @{}
        Add-IfValue -Target $payload -Key "displayName" -Value $contact.displayName
        Add-IfValue -Target $payload -Key "givenName" -Value $contact.givenName
        Add-IfValue -Target $payload -Key "surname" -Value $contact.surname
        Add-IfValue -Target $payload -Key "companyName" -Value $contact.companyName
        Add-IfValue -Target $payload -Key "jobTitle" -Value $contact.jobTitle
        Add-IfValue -Target $payload -Key "department" -Value $contact.department
        Add-IfValue -Target $payload -Key "businessPhones" -Value $contact.businessPhones
        Add-IfValue -Target $payload -Key "mobilePhone" -Value $contact.mobilePhone
        Add-IfValue -Target $payload -Key "homePhones" -Value $contact.homePhones
        Add-IfValue -Target $payload -Key "emailAddresses" -Value $contact.emailAddresses
        Add-IfValue -Target $payload -Key "imAddresses" -Value $contact.imAddresses
        Add-IfValue -Target $payload -Key "personalNotes" -Value $notes
        Add-IfValue -Target $payload -Key "categories" -Value $contact.categories

        if ($existing) {
            if ($UpdateExisting -or $notesChanged) {
                if ($DryRun) {
                    Write-Host "Would update: $displayName ($email)"
                } else {
                    $updateUri = "https://graph.microsoft.com/v1.0/users/$TargetUserId/contacts/$($existing.id)"
                    Invoke-GraphRequest -Method "PATCH" -Uri $updateUri -AccessToken $AccessToken -Body $payload | Out-Null
                    $updated++
                    Write-Host "Updated: $displayName ($email)"
                }
            } else {
                $skipped++
                Write-Host "Skipped (exists): $displayName ($email)"
            }
            continue
        }

        if ($DryRun) {
            Write-Host "Would create: $displayName ($email)"
        } else {
            if ($TargetFolderId) {
                $createUri = "https://graph.microsoft.com/v1.0/users/$TargetUserId/contactFolders/$TargetFolderId/contacts"
            } else {
                $createUri = "https://graph.microsoft.com/v1.0/users/$TargetUserId/contacts"
            }
            Invoke-GraphRequest -Method "POST" -Uri $createUri -AccessToken $AccessToken -Body $payload | Out-Null
            $created++
            Write-Host "Created: $displayName ($email)"
        }
    }

    return [pscustomobject]@{
        Created = $created
        Updated = $updated
        Skipped = $skipped
    }
}

if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
    throw "TenantId, ClientId, and ClientSecret are required."
}

if (-not $GlobalAddressBookUserId) {
    throw "GlobalAddressBookUserId is required."
}

$accessToken = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

if ($TestUserList -and $TestUserList.Count -gt 0) {
    $UserList = $TestUserList | Where-Object { $_ }
    Write-Host "Using test user list with $($UserList.Count) users." -ForegroundColor Yellow
} else {
    $licensedUsers = Get-LicensedUsers -AccessToken $accessToken
    if (-not $licensedUsers -or $licensedUsers.Count -eq 0) {
        throw "No licensed users found."
    }

    $UserList = @()
    foreach ($user in $licensedUsers) {
        if ($user.userPrincipalName) {
            $UserList += $user.userPrincipalName
        } elseif ($user.id) {
            $UserList += $user.id
        }
    }

    $UserList = $UserList | Where-Object { $_ }
    if ($UserList.Count -eq 0) {
        throw "No licensed users available."
    }
}

$UserList = $UserList | Where-Object { $_ -ne $GlobalAddressBookUserId }

$allContactsFolder = Ensure-ContactFolder -UserId $GlobalAddressBookUserId -AccessToken $accessToken -FolderName "All Contacts"
$allContactsFolderId = $allContactsFolder.id

Write-Host "Using shared mailbox All Contacts folder id: $allContactsFolderId"

Write-Host "Step 1: Sync user contacts to shared mailbox All Contacts"
$step1Created = 0
$step1Updated = 0
$step1Skipped = 0

foreach ($user in $UserList) {
    Write-Host "Syncing $user -> $GlobalAddressBookUserId"
    $result = Sync-Contacts -SourceUserId $user -TargetUserId $GlobalAddressBookUserId -AccessToken $accessToken -PageSize $PageSize -TargetFolderId $allContactsFolderId -UpdateExisting:$UpdateExisting -DryRun:$DryRun
    $step1Created += $result.Created
    $step1Updated += $result.Updated
    $step1Skipped += $result.Skipped
}

Write-Host "Step 2: Sync shared mailbox All Contacts to user contacts"
$step2Created = 0
$step2Updated = 0
$step2Skipped = 0

foreach ($user in $UserList) {
    Write-Host "Syncing $GlobalAddressBookUserId -> $user"
    $result = Sync-Contacts -SourceUserId $GlobalAddressBookUserId -TargetUserId $user -AccessToken $accessToken -PageSize $PageSize -SourceFolderId $allContactsFolderId -UpdateExisting:$UpdateExisting -DryRun:$DryRun
    $step2Created += $result.Created
    $step2Updated += $result.Updated
    $step2Skipped += $result.Skipped
}

Write-Host "Done. Step1 - Created: $step1Created, Updated: $step1Updated, Skipped: $step1Skipped"
Write-Host "Done. Step2 - Created: $step2Created, Updated: $step2Updated, Skipped: $step2Skipped"
