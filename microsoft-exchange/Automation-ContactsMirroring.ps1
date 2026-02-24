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
$rawTestUserListByTenant = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_TEST_USER_LIST_BY_TENANT_JSON')
$testUserListByTenant = $null

if ($rawTestUserListByTenant -and $rawTestUserListByTenant.Trim().Length -gt 0) {
    try {
        $testUserListByTenant = $rawTestUserListByTenant | ConvertFrom-Json
        Write-Host "Loaded tenant test user map from env var AUTOMATION_SYNCCONTACTS_TEST_USER_LIST_BY_TENANT_JSON" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to parse AUTOMATION_SYNCCONTACTS_TEST_USER_LIST_BY_TENANT_JSON: $_"
        exit 1
    }
}

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
    $rawTenants = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_MICROSOFT_TENANTS_JSON')
    $tenants = @()

    if ($rawTenants -and $rawTenants.Trim().Length -gt 0) {
        try {
            $tenants = $rawTenants | ConvertFrom-Json
        }
        catch {
            Write-Error "Failed to parse AUTOMATION_SYNCCONTACTS_MICROSOFT_TENANTS_JSON: $_"
            exit 1
        }
    }
    else {
        # Backward compatible fallback for single-tenant env vars.
        $singleTenantId = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_TENANT_ID')
        $singleClientId = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_CLIENT_ID')
        $singleClientSecret = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_CLIENT_SECRET')
        $singleGlobalAddressBookUserId = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_GLOBAL_ADDRESS_BOOK_USER_ID')

        if ($singleTenantId -and $singleClientId -and $singleClientSecret -and $singleGlobalAddressBookUserId) {
            $tenants = @(
                [PSCustomObject]@{
                    TenantId                = $singleTenantId
                    ClientId                = $singleClientId
                    ClientSecret            = $singleClientSecret
                    GlobalAddressBookUserId = $singleGlobalAddressBookUserId
                }
            )
        }
    }

    $ConfigObject = [PSCustomObject]@{
        MicrosoftTenants     = $tenants
        DryRun               = [System.Environment]::GetEnvironmentVariable('AUTOMATION_SYNCCONTACTS_DRY_RUN')
        TestUserListByTenant = $testUserListByTenant
    }

    Write-Host "Loaded configuration from individual environment variables" -ForegroundColor Green
}

# Allow a dedicated env var input field for tenant test users, even when full config JSON is used.
if ($testUserListByTenant) {
    $ConfigObject | Add-Member -NotePropertyName "TestUserListByTenant" -NotePropertyValue $testUserListByTenant -Force
}

$TenantConfigs = @()
if ($ConfigObject.MicrosoftTenants -and $ConfigObject.MicrosoftTenants.Count -gt 0) {
    $TenantConfigs = @($ConfigObject.MicrosoftTenants)
}
elseif ($ConfigObject.TenantId -and $ConfigObject.ClientId -and $ConfigObject.ClientSecret -and $ConfigObject.GlobalAddressBookUserId) {
    # Backward compatibility for older config JSON shape.
    $TenantConfigs = @(
        [PSCustomObject]@{
            TenantId                = $ConfigObject.TenantId
            ClientId                = $ConfigObject.ClientId
            ClientSecret            = $ConfigObject.ClientSecret
            GlobalAddressBookUserId = $ConfigObject.GlobalAddressBookUserId
        }
    )
}

$PageSize = 100
$UpdateExisting = $true

# Ensure console/log output uses UTF-8 so umlauts are displayed correctly.
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
$OutputEncoding = [Console]::OutputEncoding

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

function Get-TenantTestUsers {
    param(
        [Parameter(Mandatory = $true)]$TenantConfig,
        [Parameter(Mandatory = $true)]$ConfigObject
    )

    if ($TenantConfig.TestUserList -and $TenantConfig.TestUserList.Count -gt 0) {
        return @($TenantConfig.TestUserList | Where-Object { $_ })
    }

    $candidateKeys = @(
        $TenantConfig.TenantId,
        $TenantConfig.TenantID,
        $TenantConfig.PrimaryDomain,
        $TenantConfig.TenantDomain
    ) | Where-Object { $_ } | Select-Object -Unique

    if ($ConfigObject.TestUserListByTenant) {
        foreach ($key in $candidateKeys) {
            $property = $ConfigObject.TestUserListByTenant.PSObject.Properties[$key]
            if ($property -and $property.Value) {
                return @($property.Value | Where-Object { $_ })
            }
        }
    }

    return @()
}

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
        $utf8Json = [System.Text.Encoding]::UTF8.GetBytes($json)
        return Invoke-RestMethod -Method $Method -Uri $Uri -Headers $headers -ContentType "application/json; charset=utf-8" -Body $utf8Json
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

function Resolve-ContactFolderLabel {
    param(
        [Parameter(Mandatory = $true)][string]$UserId,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [string]$FolderId
    )

    if ([string]::IsNullOrWhiteSpace($FolderId)) {
        return "Default Contacts"
    }

    $folders = Get-UserContactFolders -UserId $UserId -AccessToken $AccessToken
    foreach ($folder in $folders) {
        if ($folder.id -eq $FolderId) {
            return $folder.displayName
        }
    }

    return "FolderId:$FolderId"
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

function Normalize-NotesText {
    param([string]$Notes)

    if ($null -eq $Notes) { return "" }

    # Compare notes content consistently regardless of line ending differences.
    return (($Notes -replace "`r`n", "`n" -replace "`r", "`n").Trim())
}

function Test-ContactNeedsUpdate {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Payload,
        [Parameter(Mandatory = $true)]$ExistingContact,
        [string[]]$IgnoreKeys = @()
    )

    $changedProperties = @()
    foreach ($key in ($Payload.Keys | Sort-Object)) {
        if ($IgnoreKeys -contains $key) { continue }
        $newValue = $Payload[$key]
        $existingValue = $null
        if ($ExistingContact.PSObject.Properties[$key]) {
            $existingValue = $ExistingContact.$key
        }

        $newComparable = $null
        $existingComparable = $null
        if ($key -eq "personalNotes") {
            $newComparable = Normalize-NotesText -Notes ([string]$newValue)
            $existingComparable = Normalize-NotesText -Notes ([string]$existingValue)
        }
        else {
            $newComparable = $newValue | ConvertTo-Json -Depth 6 -Compress
            $existingComparable = $existingValue | ConvertTo-Json -Depth 6 -Compress
        }

        if ($newComparable -ne $existingComparable) {
            $changedProperties += $key
        }
    }

    return $changedProperties
}

function Format-DeltaValue {
    param($Value)

    if ($null -eq $Value) { return "<null>" }
    if ($Value -is [string]) {
        if ([string]::IsNullOrWhiteSpace($Value)) { return "<empty>" }
        return $Value
    }

    return ($Value | ConvertTo-Json -Depth 6 -Compress)
}

function Get-NamePropertyDeltaText {
    param(
        [Parameter(Mandatory = $true)][hashtable]$Payload,
        [Parameter(Mandatory = $true)]$ExistingContact,
        [Parameter(Mandatory = $true)][string[]]$ChangedProperties
    )

    $nameKeys = @("displayName", "givenName", "surname")
    $segments = @()

    foreach ($key in $nameKeys) {
        if (-not ($ChangedProperties -contains $key)) { continue }
        if (-not $Payload.ContainsKey($key)) { continue }

        $newValue = $Payload[$key]
        $existingValue = $null
        if ($ExistingContact.PSObject.Properties[$key]) {
            $existingValue = $ExistingContact.$key
        }

        $segments += ("{0}: '{1}' -> '{2}'" -f $key, (Format-DeltaValue -Value $existingValue), (Format-DeltaValue -Value $newValue))
    }

    return ($segments -join "; ")
}

function Test-TargetContactOlderThanSource {
    param(
        [Parameter(Mandatory = $true)]$SourceContact,
        [Parameter(Mandatory = $true)]$TargetContact
    )

    $sourceLastModified = $null
    if ($SourceContact.PSObject.Properties["lastModifiedDateTime"]) {
        $sourceLastModified = $SourceContact.lastModifiedDateTime
    }

    $targetLastModified = $null
    if ($TargetContact.PSObject.Properties["lastModifiedDateTime"]) {
        $targetLastModified = $TargetContact.lastModifiedDateTime
    }

    if ([string]::IsNullOrWhiteSpace([string]$sourceLastModified)) {
        return $false
    }

    try {
        $sourceTimestamp = [datetimeoffset]$sourceLastModified
    }
    catch {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$targetLastModified)) {
        return $true
    }

    try {
        $targetTimestamp = [datetimeoffset]$targetLastModified
    }
    catch {
        return $true
    }

    return ($targetTimestamp -lt $sourceTimestamp)
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

function Get-LastUpdatedByFromNotes {
    param([string]$Notes)

    if ([string]::IsNullOrWhiteSpace($Notes)) { return $null }
    $match = [regex]::Match($Notes, '(?i)LastUpdatedBy\s*=\s*([^;]+)')
    if ($match.Success) { return $match.Groups[1].Value.Trim() }
    return $null
}

function Get-FirstNonSharedUserId {
    param([string[]]$Candidates)

    foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $trimmed = $candidate.Trim()
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        if ($GlobalAddressBookUserId -and $trimmed -eq $GlobalAddressBookUserId) { continue }
        return $trimmed
    }

    return $null
}

function Get-LastUpdatedAtFromNotes {
    param([string]$Notes)

    if ([string]::IsNullOrWhiteSpace($Notes)) { return $null }
    $match = [regex]::Match($Notes, '(?i)LastUpdatedAt\s*=\s*([^;]+)')
    if ($match.Success) { return $match.Groups[1].Value.Trim() }
    return $null
}

function Build-SyncNotes {
    param(
        [string]$ExistingNotes,
        [string]$SyncId,
        [string]$CreatedBy,
        [string]$LastUpdatedBy
    )

    $cleanedLines = @()
    if (-not [string]::IsNullOrWhiteSpace($ExistingNotes)) {
        foreach ($line in ($ExistingNotes -split "`r?`n")) {
            if ($line -match '(?i)SyncId\s*=' -or $line -match '(?i)CreatedBy\s*=' -or $line -match '(?i)LastUpdatedAt\s*=' -or $line -match '(?i)LastUpdatedBy\s*=') { continue }
            if (-not [string]::IsNullOrWhiteSpace($line)) { $cleanedLines += $line }
        }
    }

    # Always stamp current execution time when notes metadata is generated.
    $lastUpdatedAt = (Get-Date).ToUniversalTime().ToString("o")
    $tagLine = "SyncId=$SyncId;CreatedBy=$CreatedBy;LastUpdatedAt=$lastUpdatedAt;LastUpdatedBy=$LastUpdatedBy"
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
        [switch]$SyncSourceNotesToTarget,
        [switch]$SyncGeneratedNotesToSource,
        [switch]$UpdateExisting,
        [switch]$DryRun
    )

    $sourceSelect = "id,displayName,givenName,surname,companyName,jobTitle,department,businessPhones,mobilePhone,homePhones,emailAddresses,imAddresses,personalNotes,categories,lastModifiedDateTime,createdDateTime"
    $targetSelect = "id,displayName,givenName,surname,companyName,jobTitle,department,businessPhones,mobilePhone,homePhones,emailAddresses,imAddresses,personalNotes,categories,lastModifiedDateTime,createdDateTime"

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
        }
        elseif ($contact.displayName -and -not $targetByDisplayName.ContainsKey($contact.displayName)) {
            $targetByDisplayName[$contact.displayName] = $contact
        }
    }

    $created = 0
    $updated = 0
    $skipped = 0
    $createdRecords = @()
    $targetFolderLabel = Resolve-ContactFolderLabel -UserId $TargetUserId -AccessToken $AccessToken -FolderId $TargetFolderId
    function Update-SourceContactNotesIfNeeded {
        param(
            [Parameter(Mandatory = $true)]$SourceContact,
            [Parameter(Mandatory = $true)][string]$DesiredNotes
        )

        if (-not $SyncGeneratedNotesToSource) { return }
        if ($DryRun) { return }
        if (-not $SourceContact -or -not $SourceContact.id) { return }

        $currentNotes = $null
        if ($SourceContact.PSObject.Properties["personalNotes"]) {
            $currentNotes = $SourceContact.personalNotes
        }

        if ($currentNotes -eq $DesiredNotes) { return }

        $sourcePatchUri = "https://graph.microsoft.com/v1.0/users/$SourceUserId/contacts/$($SourceContact.id)"
        Invoke-GraphRequest -Method "PATCH" -Uri $sourcePatchUri -AccessToken $AccessToken -Body @{ personalNotes = $DesiredNotes } | Out-Null
        $SourceContact.personalNotes = $DesiredNotes
    }

    foreach ($contact in $sourceContacts) {
        $email = Get-PrimaryEmail -Contact $contact
        $displayName = $contact.displayName
        $syncId = Get-SyncIdFromNotes -Notes $contact.personalNotes

        $existing = $null
        if ($syncId -and $targetBySyncId.ContainsKey($syncId)) {
            $existing = $targetBySyncId[$syncId]
        }
        elseif ($email -and $targetByEmail.ContainsKey($email)) {
            $existing = $targetByEmail[$email]
        }
        elseif (-not $email -and $displayName -and $targetByDisplayName.ContainsKey($displayName)) {
            $existing = $targetByDisplayName[$displayName]
        }

        $existingSyncId = $null
        if ($existing) {
            $existingSyncId = Get-SyncIdFromNotes -Notes $existing.personalNotes
        }

        if (-not $syncId -and $existingSyncId) {
            $syncId = $existingSyncId
        }
        if (-not $syncId) {
            $syncId = [guid]::NewGuid().ToString()
        }

        $createdByValue = Get-CreatedByFromNotes -Notes $contact.personalNotes
        if ($existing) {
            $existingCreatedByValue = Get-CreatedByFromNotes -Notes $existing.personalNotes
            if ($existingCreatedByValue) {
                $createdByValue = $existingCreatedByValue
            }
        }

        if (-not $createdByValue) {
            $createdByValue = $SourceUserId
        }
        # Always stamp current loop source user as last updater.
        $lastUpdatedByValue = $SourceUserId

        $lastUpdatedByValue

        # Preserve the existing sync identifier when updating an existing target contact.
        $syncIdForNotes = if ($existingSyncId) { $existingSyncId } else { $syncId }
        $notesSource = $null
        if ($existing) {
            # Keep target description text; only regenerate sync metadata tags.
            $notesSource = $existing.personalNotes
        }
        $notes = Build-SyncNotes -ExistingNotes $notesSource -SyncId $syncIdForNotes -CreatedBy $createdByValue -LastUpdatedBy $lastUpdatedByValue

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
        if (-not $existing) {
            Add-IfValue -Target $payload -Key "personalNotes" -Value $notes
        }
        Add-IfValue -Target $payload -Key "categories" -Value $contact.categories

        if ($existing) {
            $updateAllowed = $UpdateExisting
            $changedProperties = @()
            $hasDataChanges = $false
            $targetIsOlderThanSource = $false
            $needsUpdate = $false
            if ($updateAllowed) {
                $changedProperties = @(Test-ContactNeedsUpdate -Payload $payload -ExistingContact $existing)
                if ($changedProperties.Count -gt 0) {
                    # Entry changed: regenerate sync metadata in notes for this write.
                    $payload["personalNotes"] = $notes
                    $changedProperties = @(Test-ContactNeedsUpdate -Payload $payload -ExistingContact $existing)
                }
                $hasDataChanges = ($changedProperties.Count -gt 0)
                if ($hasDataChanges) {
                    $targetIsOlderThanSource = Test-TargetContactOlderThanSource -SourceContact $contact -TargetContact $existing
                }
                $needsUpdate = ($hasDataChanges -and $targetIsOlderThanSource)
            }

            if ($updateAllowed -and $needsUpdate) {
                $changedPropertiesText = ($changedProperties -join ", ")
                $namePropertyDeltaText = Get-NamePropertyDeltaText -Payload $payload -ExistingContact $existing -ChangedProperties $changedProperties
                $deltaSuffix = ""
                if (-not [string]::IsNullOrWhiteSpace($namePropertyDeltaText)) {
                    $deltaSuffix = " | Delta: $namePropertyDeltaText"
                }
                if ($DryRun) {
                    Write-Host ("Would update: NotesId={0} | {1} ({2}) | Changed: {3}{4}" -f $syncIdForNotes, $displayName, $email, $changedPropertiesText, $deltaSuffix)
                }
                else {
                    $updateUri = "https://graph.microsoft.com/v1.0/users/$TargetUserId/contacts/$($existing.id)"
                    Invoke-GraphRequest -Method "PATCH" -Uri $updateUri -AccessToken $AccessToken -Body $payload | Out-Null
                    Update-SourceContactNotesIfNeeded -SourceContact $contact -DesiredNotes $notes
                    $updated++
                    Write-Host ("Updated: NotesId={0} | {1} ({2}) | Changed: {3}{4}" -f $syncIdForNotes, $displayName, $email, $changedPropertiesText, $deltaSuffix)
                }
            }
            else {
                $skipped++
                # if (-not $updateAllowed) {
                #     Write-Host ("Skipped (exists, updates disabled): NotesId={0} | ContactId={1} | {2} ({3})" -f $syncIdForNotes, $existing.id, $displayName, $email)
                # }
                # else {
                #     Write-Host ("Skipped (no changes): NotesId={0} | ContactId={1} | {2} ({3})" -f $syncIdForNotes, $existing.id, $displayName, $email)
                # }
            }
            continue
        }

        if ($DryRun) {
            Write-Host ("Would create: NotesId={0} | {1} ({2})" -f $syncIdForNotes, $displayName, $email)
            $createdRecords += [pscustomobject]@{
                Action       = "WouldCreate"
                DisplayName  = $displayName
                Email        = $email
                SyncId       = $syncIdForNotes
                CreatedBy    = $createdByValue
                SourceUser   = $SourceUserId
                TargetUser   = $TargetUserId
                TargetFolder = $targetFolderLabel
            }
        }
        else {
            if ($TargetFolderId) {
                $createUri = "https://graph.microsoft.com/v1.0/users/$TargetUserId/contactFolders/$TargetFolderId/contacts"
            }
            else {
                $createUri = "https://graph.microsoft.com/v1.0/users/$TargetUserId/contacts"
            }
            $createdContact = Invoke-GraphRequest -Method "POST" -Uri $createUri -AccessToken $AccessToken -Body $payload
            if ($createdContact -and $payload.ContainsKey("personalNotes") -and -not [string]::IsNullOrWhiteSpace($payload.personalNotes)) {
                $createdNotes = $null
                if ($createdContact.PSObject.Properties["personalNotes"]) {
                    $createdNotes = $createdContact.personalNotes
                }

                # Graph may ignore personalNotes during create for some mailboxes; enforce it immediately after create.
                if ($createdNotes -ne $payload.personalNotes) {
                    $notesPatchUri = "https://graph.microsoft.com/v1.0/users/$TargetUserId/contacts/$($createdContact.id)"
                    Invoke-GraphRequest -Method "PATCH" -Uri $notesPatchUri -AccessToken $AccessToken -Body @{ personalNotes = $payload.personalNotes } | Out-Null
                }
            }
            Update-SourceContactNotesIfNeeded -SourceContact $contact -DesiredNotes $notes
            $created++
            Write-Host ("Created: NotesId={0} | {1} ({2})" -f $syncIdForNotes, $displayName, $email)
            $createdRecords += [pscustomobject]@{
                Action       = "Created"
                DisplayName  = $displayName
                Email        = $email
                SyncId       = $syncIdForNotes
                CreatedBy    = $createdByValue
                SourceUser   = $SourceUserId
                TargetUser   = $TargetUserId
                TargetFolder = $targetFolderLabel
            }
        }
    }

    return [pscustomobject]@{
        Created        = $created
        Updated        = $updated
        Skipped        = $skipped
        CreatedRecords = $createdRecords
    }
}

function Write-CreationSummary {
    param(
        [Parameter(Mandatory = $true)][string]$Title,
        [Parameter(Mandatory = $true)][AllowNull()][AllowEmptyCollection()][array]$CreatedRecords
    )

    Write-Host "$Title creation summary" -ForegroundColor Cyan
    if (-not $CreatedRecords -or $CreatedRecords.Count -eq 0) {
        Write-Host "  No contacts were created."
        return
    }

    $groups = $CreatedRecords | Group-Object -Property CreatedBy, TargetUser, TargetFolder | Sort-Object -Property Count -Descending
    foreach ($group in $groups) {
        $first = $group.Group[0]
        Write-Host ("  {0} contact(s) | CreatedBy: {1} | Target: {2} | Folder: {3}" -f $group.Count, $first.CreatedBy, $first.TargetUser, $first.TargetFolder)
        foreach ($item in ($group.Group | Sort-Object -Property DisplayName)) {
            $name = if ($item.DisplayName) { $item.DisplayName } else { "(no displayName)" }
            $email = if ($item.Email) { $item.Email } else { "no-email" }
            Write-Host ("    - [{0}] {1} <{2}> | SyncId: {3}" -f $item.Action, $name, $email, $item.SyncId)
        }
    }
}

if (-not $TenantConfigs -or $TenantConfigs.Count -eq 0) {
    throw "MicrosoftTenants is required. Provide AUTOMATION_SYNCCONTACTS_CONFIG_JSON with MicrosoftTenants or AUTOMATION_SYNCCONTACTS_MICROSOFT_TENANTS_JSON."
}

foreach ($TenantConfig in $TenantConfigs) {
    $TenantId = $TenantConfig.TenantId
    $ClientId = $TenantConfig.ClientId
    $ClientSecret = $TenantConfig.ClientSecret
    $GlobalAddressBookUserId = $TenantConfig.GlobalAddressBookUserId

    if (-not $TenantId -or -not $ClientId -or -not $ClientSecret) {
        throw "TenantId, ClientId, and ClientSecret are required for each tenant."
    }

    if (-not $GlobalAddressBookUserId) {
        throw "GlobalAddressBookUserId is required for each tenant."
    }

    Write-Host "Running synccontacts for tenant $TenantId" -ForegroundColor Cyan
    $accessToken = Get-GraphAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret

    $TenantTestUserList = Get-TenantTestUsers -TenantConfig $TenantConfig -ConfigObject $ConfigObject
    if ($TenantTestUserList -and $TenantTestUserList.Count -gt 0) {
        $UserList = $TenantTestUserList | Where-Object { $_ }
        Write-Host "Using tenant test user list with $($UserList.Count) users." -ForegroundColor Yellow
    }
    else {
        $licensedUsers = Get-LicensedUsers -AccessToken $accessToken
        if (-not $licensedUsers -or $licensedUsers.Count -eq 0) {
            throw "No licensed users found."
        }

        $UserList = @()
        foreach ($user in $licensedUsers) {
            if ($user.userPrincipalName) {
                $UserList += $user.userPrincipalName
            }
            elseif ($user.id) {
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
    $step1CreatedRecords = @()

    foreach ($user in $UserList) {
        Write-Host "Syncing $user -> $GlobalAddressBookUserId"
        $result = Sync-Contacts -SourceUserId $user -TargetUserId $GlobalAddressBookUserId -AccessToken $accessToken -PageSize $PageSize -TargetFolderId $allContactsFolderId -SyncGeneratedNotesToSource -UpdateExisting:$UpdateExisting -DryRun:$DryRun
        $step1Created += $result.Created
        $step1Updated += $result.Updated
        $step1Skipped += $result.Skipped
        if ($result.CreatedRecords) {
            $step1CreatedRecords += $result.CreatedRecords
        }
    }

    Write-Host "Step 2: Sync shared mailbox All Contacts to user contacts"
    $step2Created = 0
    $step2Updated = 0
    $step2Skipped = 0
    $step2CreatedRecords = @()

    foreach ($user in $UserList) {
        Write-Host "Syncing $GlobalAddressBookUserId -> $user"
        $result = Sync-Contacts -SourceUserId $GlobalAddressBookUserId -TargetUserId $user -AccessToken $accessToken -PageSize $PageSize -SourceFolderId $allContactsFolderId -SyncSourceNotesToTarget -UpdateExisting:$UpdateExisting -DryRun:$DryRun
        $step2Created += $result.Created
        $step2Updated += $result.Updated
        $step2Skipped += $result.Skipped
        if ($result.CreatedRecords) {
            $step2CreatedRecords += $result.CreatedRecords
        }
    }

    Write-Host "Done for tenant $TenantId. Step1 - Created: $step1Created, Updated: $step1Updated, Skipped: $step1Skipped"
    Write-Host "Done for tenant $TenantId. Step2 - Created: $step2Created, Updated: $step2Updated, Skipped: $step2Skipped"
    Write-CreationSummary -Title "Tenant $TenantId / Step 1 (users -> shared mailbox)" -CreatedRecords $step1CreatedRecords
    Write-CreationSummary -Title "Tenant $TenantId / Step 2 (shared mailbox -> users)" -CreatedRecords $step2CreatedRecords
}