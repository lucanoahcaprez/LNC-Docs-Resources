<#
.SYNOPSIS
    Finds all objects in a Microsoft Entra ID tenant that use the memberOf operator.

.DESCRIPTION
    All Graph calls are made directly with Invoke-RestMethod using a bearer token, no
    module is required by default. Two ways to obtain that token are supported:
      - Paste an access token directly (no module install, no sign-in flow).
      - Sign in via the Microsoft Graph PowerShell module (Connect-MgGraph). The module
        is only used to obtain the token, every actual Graph call still goes through
        Invoke-RestMethod exactly as in the token-paste path.

    Three check modes are supported:
      - WholeTenant: everything below, both the Entra ID directory objects/assignment
        consumers and the Intune check.
      - EntraOnly: dynamic membership groups, administrative units, restricted
        management administrative units, entitlement management assignment policies,
        Conditional Access policies and enterprise application (service principal)
        assignments. The Intune check does not run.
      - IntuneOnly: dynamic groups are still evaluated internally (they are the only
        thing an assignment can target), but are not shown or reported. Nothing else
        on the Entra ID side is checked in this mode. Only Intune objects that use an
        affected group are reported.

    Two kinds of Entra ID objects are covered:
      - Objects that DEFINE a memberOf rule directly: dynamic membership groups,
        administrative units, restricted management administrative units,
        entitlement management assignment policies.
      - Objects that CONSUME (assign to) an affected dynamic group: Conditional
        Access policies (include/exclude group conditions) and enterprise
        application assignments (a group assigned an app role on a service
        principal).

    The Intune and Entra ID "consumer" checks take every dynamic group found to be
    affected and search for any assignment that targets one of those groups. Only the
    affected consumer objects are reported, the groups themselves are never repeated
    as rows in the report. Intune coverage:
      - Device configuration profiles, settings catalog profiles, administrative
        templates, compliance policies
      - Enrollment configurations, Autopilot deployment profiles
      - Windows PowerShell scripts, macOS shell scripts, macOS custom attribute
        scripts, proactive remediation scripts
      - Windows driver/feature/quality update profiles
      - Endpoint security policies (antivirus, disk encryption, firewall, EDR, attack
        surface reduction, account protection)
      - App assignments, app configuration policies, app protection policies
        (Android, iOS, Windows MAM, Windows WIP)

    Known gap: PIM / PIM for Groups role eligibility where a group is the principal
    is not checked, that API surface needs its own review, see .NOTES.

    The memberOf rule operator (public preview) is being retired on 03 November 2026.
    Affected objects stop updating after that date and remain in their last known state.

.PARAMETER Token
    Microsoft Graph access token (bearer token). If provided, -AuthMode defaults to
    Token and the prompt is skipped. A quick way to obtain one:
      - Graph Explorer: https://developer.microsoft.com/graph/graph-explorer (sign in, then copy the access token from the "Access token" tab)
      - Azure CLI:      az account get-access-token --resource https://graph.microsoft.com --query accessToken -o tsv

.PARAMETER AuthMode
    'Token' to paste an access token (default when -Token is supplied or left to a
    prompt), or 'Module' to sign in interactively through Connect-MgGraph. Requires
    Install-Module Microsoft.Graph.Authentication when set to Module. If omitted, the
    script prompts for a choice.

.PARAMETER Mode
    'WholeTenant', 'EntraOnly' or 'IntuneOnly', see .DESCRIPTION. If omitted, the
    script prompts for a choice.

.NOTES
    Required Graph permissions (delegated or application), regardless of auth method:
      Group.Read.All
    Additionally in WholeTenant and EntraOnly mode:
      AdministrativeUnit.Read.All
      EntitlementManagement.Read.All
      Policy.Read.All              (Conditional Access policies)
      Application.Read.All         (enterprise application assignments)
    Additionally in WholeTenant and IntuneOnly mode:
      DeviceManagementConfiguration.Read.All
      DeviceManagementApps.Read.All

    Not covered by this script, review separately if relevant:
      PIM / PIM for Groups role eligibility where a group (not a user) is the
      principal. That data lives under roleManagement/directory and
      identityGovernance/privilegedAccess/group, both need Entra ID P2 and
      RoleManagement.Read.Directory, and were left out to avoid shipping a
      half-verified implementation of a fairly involved API surface.

    No PowerShell module is required for -AuthMode Token. -AuthMode Module requires
    Install-Module Microsoft.Graph.Authentication -Scope CurrentUser, and is used
    solely to acquire an access token, not to make any of the Graph calls.

    Run the script as a file, do not paste it line by line into the console:
      .\Check-EntraMemberOfUsage.ps1
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$Token,

    [Parameter(Mandatory = $false)]
    [ValidateSet('Token', 'Module')]
    [string]$AuthMode,

    [Parameter(Mandatory = $false)]
    [ValidateSet('WholeTenant', 'EntraOnly', 'IntuneOnly')]
    [string]$Mode
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------

function Write-Section {
    param([string]$Text)
    Write-Host ''
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
    Write-Host $Text -ForegroundColor Cyan
    Write-Host ('=' * 70) -ForegroundColor DarkCyan
}

function ConvertFrom-SecureStringPlain {
    <#
        Converts a SecureString (as returned by Read-Host -AsSecureString) back to
        plain text. This is unavoidable here because the Graph bearer header needs
        the raw token value.
    #>
    param([Parameter(Mandatory = $true)][System.Security.SecureString]$SecureString)

    $bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)

    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
    }
}

function Get-JwtClaims {
    <#
        Decodes the payload of a JWT without validating its signature. Used only to
        show the caller which tenant/account the token belongs to.
    #>
    param([Parameter(Mandatory = $true)][string]$Jwt)

    $parts = $Jwt.Split('.')
    if ($parts.Count -lt 2) { return $null }

    $payload = $parts[1].Replace('-', '+').Replace('_', '/')

    switch ($payload.Length % 4) {
        2 { $payload += '==' }
        3 { $payload += '=' }
    }

    try {
        $bytes = [Convert]::FromBase64String($payload)
        $json  = [System.Text.Encoding]::UTF8.GetString($bytes)
        return $json | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Invoke-GraphRequest {
    <#
        GET wrapper around Invoke-RestMethod that adds the bearer token, retries a
        few times on throttling (HTTP 429), and, for -AuthMode Token, reacts to an
        auth-related failure (401 Unauthorized or 403 Forbidden) by asking for a
        replacement token instead of just failing that section. The actual Graph
        response is the source of truth for whether a scope is present, not the
        token's own claims, so this check happens at the point of use rather than
        upfront. This is the only way Graph is ever called in this script,
        regardless of which auth method supplied the token.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [int]$ProgressId = 0
    )

    $maxRetries = 5
    $attempt    = 0

    while ($true) {

        $attempt++

        try {
            return Invoke-RestMethod -Method GET -Uri $Uri -Headers $script:AuthHeader
        }
        catch {

            $statusCode = $null
            if ($_.Exception.Response) { $statusCode = [int]$_.Exception.Response.StatusCode }

            if ($statusCode -eq 429 -and $attempt -le $maxRetries) {
                Write-Warning ('Throttled by Microsoft Graph, waiting 10 seconds (attempt {0} of {1})' -f $attempt, $maxRetries)
                Start-Sleep -Seconds 10
                continue
            }

            if (($statusCode -eq 403 -or $statusCode -eq 401) -and $AuthMode -eq 'Token') {

                $requestPath = ($Uri -split '\?')[0]
                $reason      = if ($statusCode -eq 403) { 'insufficient privileges' } else { 'token rejected, possibly expired or missing a scope' }
                Write-Warning ('Microsoft Graph denied this request ({0}): {1}' -f $reason, $requestPath)

                # An active progress bar and an interactive prompt fighting for the
                # same console region reads as a hang on some hosts (VS Code's
                # integrated terminal included), so hide it while we ask.
                if ($ProgressId -gt 0) {
                    Write-Progress -Id $ProgressId -Activity 'Paused for input' -Completed
                }

                if (Read-YesNo -Question 'Paste a different access token and retry?') {

                    $secureToken = Read-Host -Prompt 'Paste Microsoft Graph access token' -AsSecureString
                    $newToken    = ConvertFrom-SecureStringPlain -SecureString $secureToken
                    $newToken    = $newToken.Trim() -replace '^Bearer\s+', ''

                    if (-not [string]::IsNullOrWhiteSpace($newToken)) {
                        $script:Token      = $newToken
                        $script:AuthHeader = @{ Authorization = ('Bearer {0}' -f $newToken) }
                        continue
                    }
                }
            }

            throw
        }
    }
}

function Get-GraphCollection {
    <#
        Retrieves a complete Graph collection including paging.
        Shows a progress bar while pages are being downloaded.
    #>
    param(
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$Activity,
        [int]$ProgressId = 1
    )

    $items = New-Object System.Collections.Generic.List[Object]
    $next  = $Uri
    $page  = 0

    while ($next) {

        $page++

        Write-Progress -Id $ProgressId -Activity $Activity -Status ("Downloading page {0} - {1} objects retrieved" -f $page, $items.Count) -PercentComplete -1

        $response = Invoke-GraphRequest -Uri $next -ProgressId $ProgressId

        if ($null -ne $response.value) {
            foreach ($item in $response.value) { $items.Add($item) | Out-Null }
        }

        $next = $response.'@odata.nextLink'
    }

    Write-Progress -Id $ProgressId -Activity $Activity -Completed

    return $items.ToArray()
}

function Test-MemberOfUsage {
    <#
        Checks a complete object (serialized as JSON) for the memberOf operator.
        This also covers nested structures such as specificAllowedTargets.
    #>
    param([Parameter(Mandatory = $true)]$InputObject)

    $json = $InputObject | ConvertTo-Json -Depth 30 -Compress
    return ($json -match 'memberOf')
}

function Get-AssignmentTargetGroupIds {
    <#
        Extracts the group IDs an Intune assignment array targets (include and
        exclude assignments both carry a groupId).
    #>
    param($Assignments)

    $ids = New-Object System.Collections.Generic.List[string]

    foreach ($assignment in $Assignments) {

        $target = $assignment.target
        if ($null -eq $target) { continue }

        $odataType = $target.'@odata.type'

        if ($odataType -eq '#microsoft.graph.groupAssignmentTarget' -or $odataType -eq '#microsoft.graph.exclusionGroupAssignmentTarget') {
            if ($target.groupId) { $ids.Add($target.groupId) }
        }
    }

    return $ids.ToArray()
}

function Get-RequiredGraphScopes {
    <#
        Single source of truth for which Graph scopes a check mode needs, used both
        to request consent for -AuthMode Module and to verify a pasted token for
        -AuthMode Token.
    #>
    param([Parameter(Mandatory = $true)][string]$Mode)

    switch ($Mode) {
        'IntuneOnly' { return @('Group.Read.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All') }
        'EntraOnly'  { return @('Group.Read.All', 'AdministrativeUnit.Read.All', 'EntitlementManagement.Read.All', 'Policy.Read.All', 'Application.Read.All') }
        default      { return @('Group.Read.All', 'AdministrativeUnit.Read.All', 'EntitlementManagement.Read.All', 'Policy.Read.All', 'Application.Read.All', 'DeviceManagementConfiguration.Read.All', 'DeviceManagementApps.Read.All') }
    }
}

function Read-YesNo {
    param(
        [Parameter(Mandatory = $true)][string]$Question
    )

    while ($true) {

        $answer = Read-Host ("{0} [y/n]" -f $Question)

        switch ($answer.Trim().ToLower()) {
            'y'   { return $true }
            'yes' { return $true }
            'n'   { return $false }
            'no'  { return $false }
            default { Write-Host 'Please answer with y or n.' -ForegroundColor Yellow }
        }
    }
}

function Read-Choice {
    param(
        [Parameter(Mandatory = $true)][string]$Question,
        [Parameter(Mandatory = $true)][string[]]$Options
    )

    while ($true) {

        Write-Host ''
        Write-Host $Question -ForegroundColor Cyan

        for ($i = 0; $i -lt $Options.Count; $i++) {
            Write-Host ("  [{0}] {1}" -f ($i + 1), $Options[$i])
        }

        $answer = Read-Host 'Selection'

        if ($answer -match '^\d+$') {
            $index = [int]$answer - 1
            if ($index -ge 0 -and $index -lt $Options.Count) {
                return $index
            }
        }

        Write-Host ('Please enter a number between 1 and {0}.' -f $Options.Count) -ForegroundColor Yellow
    }
}

function ConvertTo-HtmlEncoded {
    <#
        Minimal HTML-entity encoder for the handful of values written directly into
        the report markup (table cell content is set via JS textContent instead,
        which needs no encoding at all).
    #>
    param([string]$Text)

    if ([string]::IsNullOrEmpty($Text)) { return '' }

    return $Text.Replace('&', '&amp;').Replace('<', '&lt;').Replace('>', '&gt;').Replace('"', '&quot;')
}

function New-MemberOfHtmlReport {
    <#
        Writes a single self-contained HTML file (no external CSS/JS/fonts, works
        offline) with a searchable, sortable table and Entra ID / Intune workload
        tabs over the same data as the console output and CSV export.
    #>
    param(
        [Parameter(Mandatory = $true)][System.Collections.Generic.List[Object]]$Results,
        [Parameter(Mandatory = $true)][string]$Path,
        [Parameter(Mandatory = $true)][string]$Mode,
        $Claims
    )

    $entraTypes = @('Dynamic group', 'Administrative unit', 'Restricted management AU', 'Entitlement management policy', 'Conditional Access policy', 'Enterprise app assignment')

    $categoryMap = @{
        'Dynamic group'                      = 'Groups'
        'Administrative unit'                = 'Administrative units'
        'Restricted management AU'           = 'Administrative units'
        'Entitlement management policy'      = 'Entitlement management'
        'Conditional Access policy'          = 'Conditional Access'
        'Enterprise app assignment'          = 'Enterprise apps'
        'Device configuration profile'       = 'Device configuration'
        'Settings catalog profile'           = 'Device configuration'
        'Administrative template'            = 'Device configuration'
        'Compliance policy'                  = 'Compliance'
        'Enrollment configuration'           = 'Enrollment'
        'Autopilot deployment profile'       = 'Enrollment'
        'PowerShell script (Windows)'        = 'Scripts'
        'Shell script (macOS)'               = 'Scripts'
        'Custom attribute script (macOS)'    = 'Scripts'
        'Proactive remediation script'       = 'Scripts'
        'Windows driver update profile'      = 'Windows updates'
        'Windows feature update profile'     = 'Windows updates'
        'Windows quality update profile'     = 'Windows updates'
        'Endpoint security policy'           = 'Endpoint security'
        'App assignment'                     = 'Apps'
        'App configuration policy'           = 'Apps'
        'App protection policy (Android)'    = 'App protection'
        'App protection policy (iOS)'        = 'App protection'
        'App protection policy (Windows MAM)' = 'App protection'
        'App protection policy (Windows WIP)' = 'App protection'
    }

    $reportItems = @($Results | ForEach-Object {
        [PSCustomObject]@{
            Type     = $_.Type
            Name     = $_.Name
            Id       = $_.Id
            State    = $_.State
            Rule     = $_.Rule
            Workload = if ($entraTypes -contains $_.Type) { 'Entra ID' } else { 'Intune' }
            Category = if ($categoryMap.ContainsKey($_.Type)) { $categoryMap[$_.Type] } else { 'Other' }
        }
    })

    if ($reportItems.Count -eq 0) {
        $resultsJson = '[]'
    }
    else {
        $resultsJson = ConvertTo-Json -InputObject $reportItems -Depth 5 -Compress
        # ConvertTo-Json drops the [] wrapper for a single-item collection
        if ($reportItems.Count -eq 1) { $resultsJson = '[' + $resultsJson + ']' }
    }

    # Defensive only: prevent a value that happens to contain "</script" from closing
    # the embedded script block early.
    $resultsJson = $resultsJson -replace '</script', '<\/script'

    $entraCount  = ($reportItems | Where-Object { $_.Workload -eq 'Entra ID' }).Count
    $intuneCount = ($reportItems | Where-Object { $_.Workload -eq 'Intune' }).Count

    $tenantLine  = 'n/a'
    $accountLine = 'n/a'

    if ($null -ne $Claims) {

        if ($Claims.tid) { $tenantLine = $Claims.tid }

        $accountCandidate = @($Claims.upn, $Claims.preferred_username, $Claims.app_displayname, $Claims.appid) | Where-Object { $_ } | Select-Object -First 1
        if ($accountCandidate) { $accountLine = $accountCandidate }
    }

    $generatedAt = Get-Date -Format 'yyyy-MM-dd HH:mm'

    $html = @'
<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>memberOf Usage Report</title>
<style>
  :root {
    --bg: #f4f5f7;
    --panel: #ffffff;
    --border: #e2e4e9;
    --text: #1c1f26;
    --muted: #6b7280;
    --accent: #2563eb;
    --entra: #7c3aed;
    --intune: #0891b2;
    --danger: #dc2626;
    --radius: 10px;
  }
  @media (prefers-color-scheme: dark) {
    :root {
      --bg: #111318;
      --panel: #181b22;
      --border: #2a2e37;
      --text: #e7e9ee;
      --muted: #9aa1ad;
    }
  }
  * { box-sizing: border-box; }
  body {
    margin: 0;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    background: var(--bg);
    color: var(--text);
    padding: 32px 24px 64px;
  }
  .wrap { max-width: 1180px; margin: 0 auto; }
  header h1 { font-size: 22px; margin: 0 0 4px; }
  header p.meta { color: var(--muted); font-size: 13px; margin: 0 0 20px; }
  .notice {
    background: #fff7ed; border: 1px solid #fdba74; color: #9a3412;
    padding: 10px 14px; border-radius: var(--radius); font-size: 13px; margin-bottom: 24px;
    display: flex; flex-wrap: wrap; gap: 4px 10px; align-items: baseline;
  }
  .notice a { color: inherit; font-weight: 600; white-space: nowrap; }
  @media (prefers-color-scheme: dark) {
    .notice { background: #2a1c0d; border-color: #7c4a12; color: #fdba74; }
  }
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 12px; margin-bottom: 24px; }
  .card {
    background: var(--panel); border: 1px solid var(--border); border-radius: var(--radius);
    padding: 14px 16px; cursor: pointer; transition: border-color .15s;
  }
  .card:hover { border-color: var(--accent); }
  .card.active { border-color: var(--accent); box-shadow: 0 0 0 1px var(--accent); }
  .card .num { font-size: 26px; font-weight: 600; }
  .card .lbl { font-size: 12px; color: var(--muted); text-transform: uppercase; letter-spacing: .04em; }
  .card.entra .num { color: var(--entra); }
  .card.intune .num { color: var(--intune); }
  .card.total .num { color: var(--danger); }
  .toolbar { display: flex; gap: 12px; flex-wrap: wrap; align-items: center; margin-bottom: 14px; }
  .toolbar input[type=search] {
    flex: 1; min-width: 220px; padding: 9px 12px; border-radius: var(--radius);
    border: 1px solid var(--border); background: var(--panel); color: var(--text); font-size: 14px;
  }
  .toolbar select {
    padding: 9px 12px; border-radius: var(--radius); border: 1px solid var(--border);
    background: var(--panel); color: var(--text); font-size: 14px;
  }
  .toolbar button {
    padding: 9px 14px; border-radius: var(--radius); border: 1px solid var(--accent);
    background: var(--accent); color: #ffffff; font-size: 14px; font-weight: 600;
    cursor: pointer; white-space: nowrap;
  }
  .toolbar button:hover { opacity: .9; }
  .toolbar button:disabled { opacity: .5; cursor: not-allowed; }
  .count { font-size: 13px; color: var(--muted); white-space: nowrap; }
  .tablewrap {
    background: var(--panel); border: 1px solid var(--border); border-radius: var(--radius);
    overflow: auto; max-height: 65vh;
  }
  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  thead th {
    position: sticky; top: 0; background: var(--panel); text-align: left; padding: 10px 12px;
    border-bottom: 1px solid var(--border); cursor: pointer; user-select: none; white-space: nowrap;
  }
  thead th .arrow { opacity: .5; font-size: 11px; margin-left: 4px; }
  tbody td { padding: 8px 12px; border-bottom: 1px solid var(--border); vertical-align: top; word-break: break-word; }
  tbody tr:hover { background: rgba(37, 99, 235, 0.08); }
  .badge { display: inline-block; padding: 2px 8px; border-radius: 999px; font-size: 11px; font-weight: 600; white-space: nowrap; }
  .badge.entra { background: rgba(124, 58, 237, .14); color: var(--entra); }
  .badge.intune { background: rgba(8, 145, 178, .14); color: var(--intune); }
  .empty { padding: 40px; text-align: center; color: var(--muted); display: none; }
  footer { margin-top: 24px; font-size: 12px; color: var(--muted); }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1>Microsoft Entra ID &ndash; memberOf Usage Report</h1>
    <p class="meta">Generated __GENERATED_AT__ &middot; Tenant __TENANT__ &middot; Account __ACCOUNT__ &middot; Check mode: __MODE__</p>
  </header>

  <div class="notice">
    <span>The memberOf rule operator (public preview) is retired on 03 November 2026. Affected objects stop updating after that date and remain in their last known state.</span>
    <a href="https://docs.lucanoahcaprez.ch/books/azure-active-directory/page/the-end-of-the-memberof-operator-in-entra-id-dynamic-groups-find-and-remediate-impacted-objects" target="_blank" rel="noopener noreferrer">Learn more &rarr;</a>
  </div>

  <div class="cards">
    <div class="card total active" data-workload="All">
      <div class="num">__TOTAL_COUNT__</div>
      <div class="lbl">All affected objects</div>
    </div>
    <div class="card entra" data-workload="Entra ID">
      <div class="num">__ENTRA_COUNT__</div>
      <div class="lbl">Entra ID</div>
    </div>
    <div class="card intune" data-workload="Intune">
      <div class="num">__INTUNE_COUNT__</div>
      <div class="lbl">Intune</div>
    </div>
  </div>

  <div class="toolbar">
    <input type="search" id="search" placeholder="Search name, id, type, rule...">
    <select id="categoryFilter"></select>
    <button type="button" id="exportCsv">Export CSV</button>
    <span class="count" id="count"></span>
  </div>

  <div class="tablewrap">
    <table>
      <thead>
        <tr>
          <th data-key="Workload">Workload<span class="arrow"></span></th>
          <th data-key="Category">Category<span class="arrow"></span></th>
          <th data-key="Type">Type<span class="arrow"></span></th>
          <th data-key="Name">Name<span class="arrow"></span></th>
          <th data-key="Id">Id<span class="arrow"></span></th>
          <th data-key="State">State<span class="arrow"></span></th>
          <th data-key="Rule">Rule / detail<span class="arrow"></span></th>
        </tr>
      </thead>
      <tbody id="rows"></tbody>
    </table>
    <div class="empty" id="empty">No matching objects.</div>
  </div>

  <footer>Generated by Check-EntraMemberOfUsage.ps1</footer>
</div>

<script>
const DATA = __RESULTS_JSON__;

const state = { workload: 'All', category: '', query: '', sortKey: 'Type', sortDir: 'asc' };

const rowsEl           = document.getElementById('rows');
const countEl          = document.getElementById('count');
const emptyEl          = document.getElementById('empty');
const searchEl         = document.getElementById('search');
const categorySelectEl = document.getElementById('categoryFilter');
const exportBtnEl      = document.getElementById('exportCsv');

function csvEscape(value) {
  const s = (value === null || value === undefined) ? '' : String(value);
  if (/[";\r\n]/.test(s)) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function pad2(n) {
  return n < 10 ? '0' + n : '' + n;
}

function exportCsv() {
  const rows = applyFilters();

  if (rows.length === 0) {
    return;
  }

  const headers = ['Workload', 'Category', 'Type', 'Name', 'Id', 'State', 'Rule'];
  const lines = [headers.join(';')];

  rows.forEach(function (item) {
    lines.push(headers.map(function (h) { return csvEscape(item[h]); }).join(';'));
  });

  // Delimiter matches the PowerShell script's own CSV export; the leading BOM makes
  // Excel detect UTF-8 instead of mangling non-ASCII names.
  const bom = String.fromCharCode(0xFEFF);
  const csvContent = bom + lines.join('\r\n');
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);

  const d = new Date();
  const dateStr = pad2(d.getDate()) + '-' + pad2(d.getMonth() + 1) + '-' + d.getFullYear();

  const a = document.createElement('a');
  a.href = url;
  a.download = 'MemberOf-Report-' + dateStr + '.csv';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

exportBtnEl.addEventListener('click', exportCsv);

function updateCategoryOptions() {
  const relevant = state.workload === 'All' ? DATA : DATA.filter(function (d) { return d.Workload === state.workload; });
  const cats = Array.from(new Set(relevant.map(function (d) { return d.Category; }))).sort();

  categorySelectEl.innerHTML = '';

  const allOpt = document.createElement('option');
  allOpt.value = '';
  allOpt.textContent = 'All categories';
  categorySelectEl.appendChild(allOpt);

  cats.forEach(function (c) {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = c;
    categorySelectEl.appendChild(opt);
  });

  state.category = '';
  categorySelectEl.value = '';
}

function applyFilters() {
  const q = state.query.trim().toLowerCase();

  let filtered = DATA.filter(function (item) {
    if (state.workload !== 'All' && item.Workload !== state.workload) return false;
    if (state.category && item.Category !== state.category) return false;
    if (!q) return true;
    return ['Type', 'Name', 'Id', 'State', 'Rule', 'Workload', 'Category'].some(function (key) {
      return (item[key] || '').toString().toLowerCase().indexOf(q) !== -1;
    });
  });

  filtered.sort(function (a, b) {
    const av = (a[state.sortKey] || '').toString().toLowerCase();
    const bv = (b[state.sortKey] || '').toString().toLowerCase();
    if (av < bv) return state.sortDir === 'asc' ? -1 : 1;
    if (av > bv) return state.sortDir === 'asc' ? 1 : -1;
    return 0;
  });

  return filtered;
}

function render() {
  const filtered = applyFilters();

  rowsEl.innerHTML = '';

  filtered.forEach(function (item) {
    const tr = document.createElement('tr');

    const workloadTd = document.createElement('td');
    const badge = document.createElement('span');
    badge.className = 'badge ' + (item.Workload === 'Entra ID' ? 'entra' : 'intune');
    badge.textContent = item.Workload;
    workloadTd.appendChild(badge);
    tr.appendChild(workloadTd);

    ['Category', 'Type', 'Name', 'Id', 'State', 'Rule'].forEach(function (key) {
      const td = document.createElement('td');
      td.textContent = item[key] || '';
      tr.appendChild(td);
    });

    rowsEl.appendChild(tr);
  });

  countEl.textContent = 'Showing ' + filtered.length + ' of ' + DATA.length;
  emptyEl.style.display = filtered.length === 0 ? 'block' : 'none';

  exportBtnEl.textContent = 'Export CSV (' + filtered.length + ')';
  exportBtnEl.disabled = filtered.length === 0;

  document.querySelectorAll('.card').forEach(function (card) {
    card.classList.toggle('active', card.getAttribute('data-workload') === state.workload);
  });

  document.querySelectorAll('th[data-key]').forEach(function (th) {
    const arrow = th.querySelector('.arrow');
    if (th.getAttribute('data-key') === state.sortKey) {
      arrow.textContent = state.sortDir === 'asc' ? '▲' : '▼';
    } else {
      arrow.textContent = '';
    }
  });
}

document.querySelectorAll('.card').forEach(function (card) {
  card.addEventListener('click', function () {
    state.workload = card.getAttribute('data-workload');
    updateCategoryOptions();
    render();
  });
});

categorySelectEl.addEventListener('change', function () {
  state.category = categorySelectEl.value;
  render();
});

document.querySelectorAll('th[data-key]').forEach(function (th) {
  th.addEventListener('click', function () {
    const key = th.getAttribute('data-key');
    if (state.sortKey === key) {
      state.sortDir = state.sortDir === 'asc' ? 'desc' : 'asc';
    } else {
      state.sortKey = key;
      state.sortDir = 'asc';
    }
    render();
  });
});

searchEl.addEventListener('input', function () {
  state.query = searchEl.value;
  render();
});

updateCategoryOptions();
render();
</script>
</body>
</html>
'@

    $html = $html.Replace('__RESULTS_JSON__', $resultsJson)
    $html = $html.Replace('__GENERATED_AT__', (ConvertTo-HtmlEncoded -Text $generatedAt))
    $html = $html.Replace('__TENANT__', (ConvertTo-HtmlEncoded -Text $tenantLine))
    $html = $html.Replace('__ACCOUNT__', (ConvertTo-HtmlEncoded -Text $accountLine))
    $html = $html.Replace('__MODE__', (ConvertTo-HtmlEncoded -Text $Mode))
    $html = $html.Replace('__TOTAL_COUNT__', [string]$reportItems.Count)
    $html = $html.Replace('__ENTRA_COUNT__', [string]$entraCount)
    $html = $html.Replace('__INTUNE_COUNT__', [string]$intuneCount)

    Set-Content -Path $Path -Value $html -Encoding UTF8
}

# ---------------------------------------------------------------------
# Check mode
# ---------------------------------------------------------------------

if (-not $Mode) {

    $modeChoice = Read-Choice -Question 'What do you want to check?' -Options @(
        'Whole Tenant (Entra ID directory objects and Intune configuration profiles/app assignments)',
        'Entra Only (dynamic groups, administrative units, entitlement management policies) - Intune is not checked',
        'Intune Only (configuration profiles and app assignments that use an affected dynamic group) - groups, AUs and entitlement policies are not shown'
    )

    $Mode = switch ($modeChoice) {
        0 { 'WholeTenant' }
        1 { 'EntraOnly' }
        2 { 'IntuneOnly' }
    }
}

# ---------------------------------------------------------------------
# Authentication
# ---------------------------------------------------------------------

if (-not $AuthMode) {

    if ($Token) {
        $AuthMode = 'Token'
    }
    else {
        $authChoice = Read-Choice -Question 'How do you want to authenticate to Microsoft Graph?' -Options @(
            'Paste an access token (no module required)',
            'Sign in interactively via the Microsoft Graph PowerShell module (Connect-MgGraph)'
        )
        $AuthMode = if ($authChoice -eq 1) { 'Module' } else { 'Token' }
    }
}

$requiredScopes = Get-RequiredGraphScopes -Mode $Mode
$tokenValidated = $false

if ($AuthMode -eq 'Module') {

    if (-not (Get-Command Connect-MgGraph -ErrorAction SilentlyContinue)) {
        throw 'Module Microsoft.Graph.Authentication is missing. Install it with: Install-Module Microsoft.Graph.Authentication -Scope CurrentUser'
    }

    Write-Host 'Connecting to Microsoft Graph. Consent for the required scopes may be requested.' -ForegroundColor Yellow
    Connect-MgGraph -Scopes $requiredScopes -NoWelcome | Out-Null

    # Microsoft.Graph.Authentication does not expose the raw bearer token through a
    # public cmdlet. It is instead read off the authenticated request Invoke-MgGraphRequest
    # just made, by asking for the HttpResponseMessage and inspecting the Authorization
    # header it sent. This relies on internal, undocumented behavior and is not an
    # officially supported method, but it is stable in practice and widely used for
    # exactly this bridge-to-raw-REST scenario.
    $probe = Invoke-MgGraphRequest -Method GET -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' -OutputType HttpResponseMessage
    $Token = $probe.RequestMessage.Headers.Authorization.Parameter

    if ([string]::IsNullOrWhiteSpace($Token)) {
        throw 'Could not retrieve the access token from the Microsoft Graph PowerShell session.'
    }

    $tokenValidated = $true
}
elseif ([string]::IsNullOrWhiteSpace($Token)) {

    $secureToken = Read-Host -Prompt 'Paste Microsoft Graph access token' -AsSecureString
    $Token = ConvertFrom-SecureStringPlain -SecureString $secureToken
}

$Token = $Token.Trim() -replace '^Bearer\s+', ''

if ([string]::IsNullOrWhiteSpace($Token)) {
    throw 'No access token was provided.'
}

$script:AuthHeader = @{ Authorization = ('Bearer {0}' -f $Token) }

$claims = Get-JwtClaims -Jwt $Token

Write-Host ''

if ($null -ne $claims) {

    $account = @($claims.upn, $claims.preferred_username, $claims.app_displayname, $claims.appid) | Where-Object { $_ } | Select-Object -First 1

    Write-Host ('Tenant  : {0}' -f $claims.tid)
    Write-Host ('Account : {0}' -f $account)

    if ($claims.exp) {

        $expires = [DateTimeOffset]::FromUnixTimeSeconds([int64]$claims.exp).LocalDateTime

        if ($expires -lt (Get-Date)) {
            Write-Warning ('The token expired at {0}.' -f $expires)
        }
        else {
            Write-Host ('Expires : {0}' -f $expires)
        }
    }
}
else {
    Write-Warning 'The token does not look like a valid access token (JWT), continuing anyway.'
}

if (-not $tokenValidated) {

    Write-Host 'Validating token against Microsoft Graph...' -ForegroundColor Yellow

    try {
        Invoke-GraphRequest -Uri 'https://graph.microsoft.com/v1.0/organization?$select=id' | Out-Null
        Write-Host 'Token accepted.' -ForegroundColor Green
    }
    catch {
        Write-Warning ('Token validation call failed, continuing anyway: {0}' -f $_.Exception.Message)
    }
}
else {
    Write-Host 'Token accepted.' -ForegroundColor Green
}

$results          = New-Object System.Collections.Generic.List[Object]
$affectedGroupIds = @{}

# ---------------------------------------------------------------------
# Dynamic membership groups
# ---------------------------------------------------------------------

Write-Section 'Dynamic membership groups'

if ($Mode -eq 'IntuneOnly') {
    Write-Host 'Groups are evaluated to find which ones are affected, but are not shown, only Intune usage of them is reported below.' -ForegroundColor DarkGray
}

$groupUri = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,groupTypes,membershipRule,membershipRuleProcessingState&$top=999'
$groups   = @()

try {
    $groups = Get-GraphCollection -Uri $groupUri -Activity 'Reading groups' -ProgressId 1
}
catch {
    Write-Warning ('Groups could not be read: {0}' -f $_.Exception.Message)
}

$groupIndex = 0

foreach ($group in $groups) {

    $groupIndex++

    if ($groups.Count -gt 0) {
        Write-Progress -Id 1 -Activity 'Analyzing groups' -Status ("{0} of {1}" -f $groupIndex, $groups.Count) -PercentComplete (($groupIndex / $groups.Count) * 100)
    }

    if ($group.groupTypes -contains 'DynamicMembership' -and $group.membershipRule -and $group.membershipRule -match 'memberOf') {

        $affectedGroupIds[$group.id] = $group.displayName

        if ($Mode -ne 'IntuneOnly') {

            $results.Add([PSCustomObject]@{
                Type  = 'Dynamic group'
                Name  = $group.displayName
                Id    = $group.id
                State = $group.membershipRuleProcessingState
                Rule  = $group.membershipRule
            }) | Out-Null
        }
    }
}

Write-Progress -Id 1 -Activity 'Analyzing groups' -Completed

Write-Host ('Checked: {0} groups / Affected: {1}' -f $groups.Count, $affectedGroupIds.Count)

if ($Mode -ne 'IntuneOnly') {

    # ---------------------------------------------------------------------
    # Administrative units and restricted management units
    # ---------------------------------------------------------------------

    Write-Section 'Administrative units and restricted management administrative units'

    # isMemberManagementRestricted is only reliably available in the beta endpoint
    $auUri = 'https://graph.microsoft.com/beta/directory/administrativeUnits?$top=999'
    $aus   = @()

    try {
        $aus = Get-GraphCollection -Uri $auUri -Activity 'Reading administrative units' -ProgressId 2
    }
    catch {
        Write-Warning ('Administrative units could not be read: {0}' -f $_.Exception.Message)
    }

    $auIndex = 0

    foreach ($au in $aus) {

        $auIndex++

        if ($aus.Count -gt 0) {
            Write-Progress -Id 2 -Activity 'Analyzing administrative units' -Status ("{0} of {1}" -f $auIndex, $aus.Count) -PercentComplete (($auIndex / $aus.Count) * 100)
        }

        if ($au.membershipRule -and $au.membershipRule -match 'memberOf') {

            $auType = 'Administrative unit'

            if ($au.isMemberManagementRestricted -eq $true) {
                $auType = 'Restricted management AU'
            }

            $results.Add([PSCustomObject]@{
                Type  = $auType
                Name  = $au.displayName
                Id    = $au.id
                State = $au.membershipRuleProcessingState
                Rule  = $au.membershipRule
            }) | Out-Null
        }
    }

    Write-Progress -Id 2 -Activity 'Analyzing administrative units' -Completed

    $affectedAuCount = ($results | Where-Object { $_.Type -like '*administrative*' -or $_.Type -like '*AU*' }).Count
    Write-Host ('Checked: {0} administrative units / Affected: {1}' -f $aus.Count, $affectedAuCount)

    # ---------------------------------------------------------------------
    # Entitlement management assignment policies
    # ---------------------------------------------------------------------

    Write-Section 'Entitlement management assignment policies'

    # Resource names differ between v1.0 and beta, therefore a fallback list is used
    $emUris = @(
        'https://graph.microsoft.com/v1.0/identityGovernance/entitlementManagement/assignmentPolicies?$top=100',
        'https://graph.microsoft.com/beta/identityGovernance/entitlementManagement/accessPackageAssignmentPolicies?$top=100'
    )

    $policies    = @()
    $emBaseUri   = $null
    $emSucceeded = $false

    foreach ($uri in $emUris) {

        try {
            $policies    = Get-GraphCollection -Uri $uri -Activity 'Reading entitlement management policies' -ProgressId 3
            $emBaseUri   = ($uri -split '\?')[0]
            $emSucceeded = $true
            Write-Verbose ('Endpoint used: {0}' -f $uri)
            break
        }
        catch {
            Write-Verbose ('Endpoint failed: {0} ({1})' -f $uri, $_.Exception.Message)
        }
    }

    if (-not $emSucceeded) {

        Write-Warning 'Entitlement management policies could not be read.'
        Write-Warning 'Possible reasons: the token is missing EntitlementManagement.Read.All, or entitlement management is not enabled in this tenant. Microsoft Entra ID P2 is required.'
    }
    else {

        $affectedPolicyCount = 0
        $policyIndex = 0

        foreach ($policy in $policies) {

            $policyIndex++

            if ($policies.Count -gt 0) {
                Write-Progress -Id 3 -Activity 'Analyzing entitlement management policies' -Status ("{0} of {1}" -f $policyIndex, $policies.Count) -PercentComplete (($policyIndex / $policies.Count) * 100)
            }

            # The auto assignment rule is stored in specificAllowedTargets and is not
            # always returned by the list operation, therefore each policy is read individually
            $detail = $policy

            try {
                $detailUri = '{0}/{1}' -f $emBaseUri, $policy.id
                $detail = Invoke-GraphRequest -Uri $detailUri -ProgressId 3
            }
            catch {
                Write-Verbose ('Detail request failed for policy {0}' -f $policy.id)
            }

            if (Test-MemberOfUsage -InputObject $detail) {

                $affectedPolicyCount++

                $results.Add([PSCustomObject]@{
                    Type  = 'Entitlement management policy'
                    Name  = $detail.displayName
                    Id    = $detail.id
                    State = $detail.allowedTargetScope
                    Rule  = 'memberOf found in policy definition'
                }) | Out-Null
            }
        }

        Write-Progress -Id 3 -Activity 'Analyzing entitlement management policies' -Completed

        Write-Host ('Checked: {0} policies / Affected: {1}' -f $policies.Count, $affectedPolicyCount)
    }

    # ---------------------------------------------------------------------
    # Conditional Access policies
    # ---------------------------------------------------------------------

    Write-Section 'Conditional Access policies'

    if ($affectedGroupIds.Count -eq 0) {

        Write-Host 'No affected dynamic groups were found, Conditional Access policies cannot reference one.' -ForegroundColor Green
    }
    else {

        $caUri      = 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies?$select=id,displayName,state,conditions'
        $caPolicies = @()

        try {
            $caPolicies = Get-GraphCollection -Uri $caUri -Activity 'Reading Conditional Access policies' -ProgressId 4
        }
        catch {
            Write-Warning ('Conditional Access policies could not be read: {0}' -f $_.Exception.Message)
        }

        $caMatchCount = 0

        foreach ($policy in $caPolicies) {

            $referencedGroupIds = @($policy.conditions.users.includeGroups) + @($policy.conditions.users.excludeGroups) | Where-Object { $_ }
            $matchedGroupIds    = $referencedGroupIds | Where-Object { $affectedGroupIds.ContainsKey($_) } | Select-Object -Unique

            if ($matchedGroupIds) {

                $caMatchCount++
                $matchedGroupNames = $matchedGroupIds | ForEach-Object { $affectedGroupIds[$_] }

                $results.Add([PSCustomObject]@{
                    Type  = 'Conditional Access policy'
                    Name  = $policy.displayName
                    Id    = $policy.id
                    State = $policy.state
                    Rule  = ('Affected group(s): {0}' -f ($matchedGroupNames -join ', '))
                }) | Out-Null
            }
        }

        Write-Host ('Checked: {0} Conditional Access policies / Affected: {1}' -f $caPolicies.Count, $caMatchCount)
    }

    # ---------------------------------------------------------------------
    # Enterprise application (service principal) assignments
    # ---------------------------------------------------------------------

    Write-Section 'Enterprise application assignments'

    if ($affectedGroupIds.Count -eq 0) {

        Write-Host 'No affected dynamic groups were found, application assignments cannot reference one.' -ForegroundColor Green
    }
    else {

        # Queried per affected group (group -> its appRoleAssignments) rather than per
        # service principal, so the call count scales with affected groups, not with
        # the number of enterprise apps in the tenant.
        $appAssignmentMatches = @{}
        $groupsChecked        = 0

        foreach ($groupId in $affectedGroupIds.Keys) {

            $groupsChecked++
            $assignments = @()

            try {
                $assignmentUri = 'https://graph.microsoft.com/v1.0/groups/{0}/appRoleAssignments' -f $groupId
                $assignments   = Get-GraphCollection -Uri $assignmentUri -Activity 'Reading application assignments' -ProgressId 5
            }
            catch {
                Write-Verbose ('Application assignments could not be read for group {0}: {1}' -f $groupId, $_.Exception.Message)
                continue
            }

            $groupName = $affectedGroupIds[$groupId]

            foreach ($assignment in $assignments) {

                if (-not $appAssignmentMatches.ContainsKey($assignment.resourceId)) {
                    $appAssignmentMatches[$assignment.resourceId] = [PSCustomObject]@{
                        Name       = $assignment.resourceDisplayName
                        GroupNames = New-Object System.Collections.Generic.List[string]
                    }
                }

                if (-not $appAssignmentMatches[$assignment.resourceId].GroupNames.Contains($groupName)) {
                    $appAssignmentMatches[$assignment.resourceId].GroupNames.Add($groupName)
                }
            }
        }

        foreach ($resourceId in $appAssignmentMatches.Keys) {

            $entry = $appAssignmentMatches[$resourceId]

            $results.Add([PSCustomObject]@{
                Type  = 'Enterprise app assignment'
                Name  = $entry.Name
                Id    = $resourceId
                State = 'Assigned to affected dynamic group'
                Rule  = ('Affected group(s): {0}' -f ($entry.GroupNames -join ', '))
            }) | Out-Null
        }

        Write-Host ('Checked: {0} affected group(s) for application assignments / Affected apps: {1}' -f $groupsChecked, $appAssignmentMatches.Count)
    }
}

# ---------------------------------------------------------------------
# Intune workloads
# ---------------------------------------------------------------------

if ($Mode -eq 'WholeTenant' -or $Mode -eq 'IntuneOnly') {

    Write-Section 'Intune workloads'

    if ($affectedGroupIds.Count -eq 0) {

        Write-Host 'No affected dynamic groups were found, Intune assignments cannot reference one.' -ForegroundColor Green
    }
    else {

        # Assignments are pulled via $expand so each source needs a single paged read,
        # not one detail request per object. $expand=assignments is the pattern
        # Microsoft's own bulk-assignment samples use across these resource types; if
        # a given type ever stops supporting it the request fails outright (caught
        # below) rather than silently under-reporting.
        $intuneSources = @(
            @{ Type = 'Device configuration profile';        Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceConfigurations?$expand=assignments&$top=999' },
            @{ Type = 'Settings catalog profile';             Uri = 'https://graph.microsoft.com/beta/deviceManagement/configurationPolicies?$expand=assignments&$top=999' },
            @{ Type = 'Administrative template';              Uri = 'https://graph.microsoft.com/beta/deviceManagement/groupPolicyConfigurations?$expand=assignments&$top=999' },
            @{ Type = 'Compliance policy';                    Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceCompliancePolicies?$expand=assignments&$top=999' },
            @{ Type = 'Enrollment configuration';             Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceEnrollmentConfigurations?$expand=assignments&$top=999' },
            @{ Type = 'Autopilot deployment profile';         Uri = 'https://graph.microsoft.com/beta/deviceManagement/windowsAutopilotDeploymentProfiles?$expand=assignments&$top=999' },
            @{ Type = 'PowerShell script (Windows)';          Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceManagementScripts?$expand=assignments&$top=999' },
            @{ Type = 'Shell script (macOS)';                 Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceShellScripts?$expand=assignments&$top=999' },
            @{ Type = 'Custom attribute script (macOS)';      Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceCustomAttributeShellScripts?$expand=assignments&$top=999' },
            @{ Type = 'Proactive remediation script';         Uri = 'https://graph.microsoft.com/beta/deviceManagement/deviceHealthScripts?$expand=assignments&$top=999' },
            @{ Type = 'Windows driver update profile';        Uri = 'https://graph.microsoft.com/beta/deviceManagement/windowsDriverUpdateProfiles?$expand=assignments&$top=999' },
            @{ Type = 'Windows feature update profile';       Uri = 'https://graph.microsoft.com/beta/deviceManagement/windowsFeatureUpdateProfiles?$expand=assignments&$top=999' },
            @{ Type = 'Windows quality update profile';       Uri = 'https://graph.microsoft.com/beta/deviceManagement/windowsQualityUpdateProfiles?$expand=assignments&$top=999' },
            @{ Type = 'Endpoint security policy';             Uri = 'https://graph.microsoft.com/beta/deviceManagement/intents?$expand=assignments&$top=999' },
            @{ Type = 'App assignment';                       Uri = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?$expand=assignments&$top=999' },
            @{ Type = 'App configuration policy';             Uri = 'https://graph.microsoft.com/beta/deviceAppManagement/mobileAppConfigurations?$expand=assignments&$top=999' },
            @{ Type = 'App protection policy (Android)';      Uri = 'https://graph.microsoft.com/beta/deviceAppManagement/androidManagedAppProtections?$expand=assignments&$top=999' },
            @{ Type = 'App protection policy (iOS)';          Uri = 'https://graph.microsoft.com/beta/deviceAppManagement/iosManagedAppProtections?$expand=assignments&$top=999' },
            @{ Type = 'App protection policy (Windows MAM)';  Uri = 'https://graph.microsoft.com/beta/deviceAppManagement/windowsManagedAppProtections?$expand=assignments&$top=999' },
            @{ Type = 'App protection policy (Windows WIP)';  Uri = 'https://graph.microsoft.com/beta/deviceAppManagement/mdmWindowsInformationProtectionPolicies?$expand=assignments&$top=999' }
        )

        $intuneResults = New-Object System.Collections.Generic.List[Object]
        $sourceIndex   = 0

        foreach ($source in $intuneSources) {

            $sourceIndex++
            $items = @()

            try {
                $items = Get-GraphCollection -Uri $source.Uri -Activity ('Reading {0}' -f $source.Type) -ProgressId (10 + $sourceIndex)
            }
            catch {
                Write-Warning ('{0} could not be read: {1}' -f $source.Type, $_.Exception.Message)
                continue
            }

            $matchCount = 0

            foreach ($item in $items) {

                $targetGroupIds  = Get-AssignmentTargetGroupIds -Assignments $item.assignments
                $matchedGroupIds = $targetGroupIds | Where-Object { $affectedGroupIds.ContainsKey($_) } | Select-Object -Unique

                if ($matchedGroupIds) {

                    $matchCount++
                    $matchedGroupNames = $matchedGroupIds | ForEach-Object { $affectedGroupIds[$_] }
                    $name = if ($item.displayName) { $item.displayName } else { $item.name }

                    $intuneResults.Add([PSCustomObject]@{
                        Type  = $source.Type
                        Name  = $name
                        Id    = $item.id
                        State = 'Assigned to affected dynamic group'
                        Rule  = ('Affected group(s): {0}' -f ($matchedGroupNames -join ', '))
                    }) | Out-Null
                }
            }

            Write-Host ('Checked: {0} {1} / Affected: {2}' -f $items.Count, $source.Type, $matchCount)
        }

        # Merge into the main result set so the existing output/CSV logic below covers
        # everything in one report. Only the Intune objects are added, never the groups
        # they reference.
        foreach ($intuneResult in $intuneResults) {
            $results.Add($intuneResult) | Out-Null
        }
    }
}

# ---------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------

Write-Section 'RESULT'

if ($results.Count -eq 0) {

    if ($Mode -eq 'IntuneOnly') {
        Write-Host 'No Intune configuration profile or app assignment uses a dynamic group affected by the memberOf retirement.' -ForegroundColor Green
    }
    else {
        Write-Host 'No usage of the memberOf operator was found in this tenant.' -ForegroundColor Green
    }
}
else {

    Write-Host ('Affected objects in total: {0}' -f $results.Count) -ForegroundColor Red
    Write-Host ''
    Write-Host 'Summary by object type:' -ForegroundColor Yellow

    $results |
        Group-Object Type |
        Sort-Object Name |
        Select-Object @{ Name = 'Type'; Expression = { $_.Name } }, Count |
        Format-Table -AutoSize |
        Out-String -Width 4096 |
        Write-Host

    Write-Host 'See the CSV/HTML report for the individual affected objects.' -ForegroundColor DarkGray
    Write-Host ''
    Write-Host 'Action required: replace all memberOf configurations before 03 November 2026.' -ForegroundColor Yellow
}

# ---------------------------------------------------------------------
# Optional CSV report
# ---------------------------------------------------------------------

$scriptPath = $PSScriptRoot

if ([string]::IsNullOrWhiteSpace($scriptPath)) {
    $scriptPath = (Get-Location).Path
}

Write-Host ''

# A file name cannot contain slashes, therefore the day-month-year format uses dashes
$reportDate = Get-Date -Format 'dd-MM-yyyy'

if (Read-YesNo -Question ("Create CSV report under {0}?" -f $scriptPath)) {

    $reportName = 'MemberOf-Report-{0}.csv' -f $reportDate
    $reportPath = Join-Path -Path $scriptPath -ChildPath $reportName

    try {
        $results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8 -Delimiter ';'
        Write-Host ('CSV report created: {0}' -f $reportPath) -ForegroundColor Green
    }
    catch {
        Write-Warning ('CSV report could not be created: {0}' -f $_.Exception.Message)
    }
}
else {
    Write-Host 'No CSV report was created.' -ForegroundColor Yellow
}

# ---------------------------------------------------------------------
# Optional HTML report
# ---------------------------------------------------------------------

if (Read-YesNo -Question ("Create HTML report under {0}?" -f $scriptPath)) {

    $htmlReportName = 'MemberOf-Report-{0}.html' -f $reportDate
    $htmlReportPath = Join-Path -Path $scriptPath -ChildPath $htmlReportName

    try {
        New-MemberOfHtmlReport -Results $results -Path $htmlReportPath -Mode $Mode -Claims $claims
        Write-Host ('HTML report created: {0}' -f $htmlReportPath) -ForegroundColor Green
    }
    catch {
        Write-Warning ('HTML report could not be created: {0}' -f $_.Exception.Message)
    }
}
else {
    Write-Host 'No HTML report was created.' -ForegroundColor Yellow
}
