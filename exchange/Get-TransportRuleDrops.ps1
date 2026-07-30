<#
.SYNOPSIS
    Finds emails dropped or rejected by a specific Exchange transport rule and exports
    sender, recipient, subject, date, and action to CSV.
.DESCRIPTION
    Connects to Exchange Online, prompts for a date range and transport rule name, then
    queries Message Trace and Message Trace Detail to identify messages affected by the rule.
    Splits date ranges into 10-day windows automatically (Exchange Online limitation).
    Results are exported to a timestamped CSV.

    NOTE: Message Trace supports a maximum lookback of 90 days.
    NOTE: This script calls Get-MessageTraceDetail for each message in the date range.
    For large tenants or wide date ranges, use -StatusFilter to narrow results first.
.PARAMETER StartDate
    Start of the search window. Must be within the last 90 days.
.PARAMETER EndDate
    End of the search window.
.PARAMETER RuleName
    The transport rule name to search for (partial match supported).
.PARAMETER StatusFilter
    Optional. Pre-filter messages by delivery status to reduce API calls.
    Recommended: 'Failed' for reject rules, 'All' to check every message.
    Defaults to prompting interactively.
.PARAMETER AdminUPN
    Admin UPN used to connect if no active Exchange Online session exists.
.PARAMETER OutputCsv
    Path for the output CSV. Defaults to .\TransportRuleDrops_<timestamp>.csv
.PARAMETER TenantEnvironment
    Exchange Online environment. Defaults to USGovGCC.
.EXAMPLE
    .\Get-TransportRuleDrops.ps1
.EXAMPLE
    .\Get-TransportRuleDrops.ps1 -StartDate (Get-Date).AddDays(-7) -RuleName "Block External Fwd" -StatusFilter Failed
#>

#Requires -Modules ExchangeOnlineManagement

[CmdletBinding()]
param(
    [Parameter()]
    [datetime]$StartDate,

    [Parameter()]
    [datetime]$EndDate,

    [Parameter()]
    [string]$RuleName,

    [Parameter()]
    [ValidateSet('All', 'Failed', 'Quarantined', 'FilteredAsSpam')]
    [string]$StatusFilter,

    [Parameter()]
    [string]$AdminUPN,

    [Parameter()]
    [string]$OutputCsv = ".\TransportRuleDrops_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [ValidateSet('Commercial', 'USGovGCC', 'USGovGCCHigh')]
    [string]$TenantEnvironment = 'USGovGCC'
)

# ---------- Module guard ----------
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable | Where-Object { $_.Version -ge '3.0.0' })) {
    Write-Error 'ExchangeOnlineManagement v3+ is required. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    exit 1
}

# ---------- Connect ----------
$existingConn = Get-ConnectionInformation | Where-Object { $_.State -eq 'Connected' } | Select-Object -First 1
$connectedHere = $false

if ($existingConn) {
    Write-Host "Using existing Exchange Online connection ($($existingConn.UserPrincipalName))." -ForegroundColor Green
} else {
    if ([string]::IsNullOrWhiteSpace($AdminUPN)) {
        Write-Host "`nNo active Exchange Online connection found." -ForegroundColor Yellow
        $AdminUPN = (Read-Host 'Enter your admin UPN (e.g. admin@contoso.com)').Trim()
        if ([string]::IsNullOrWhiteSpace($AdminUPN)) { Write-Error 'No admin UPN provided.'; exit 1 }
    }
    $connectParams = @{ UserPrincipalName = $AdminUPN; ShowBanner = $false; Device = $true }
    if ($TenantEnvironment -eq 'USGovGCCHigh') { $connectParams['ExchangeEnvironmentName'] = 'O365USGovGCCHigh' }
    Write-Host "Connecting to Exchange Online as $AdminUPN ..." -ForegroundColor Cyan
    Write-Host 'A device login prompt will appear — go to https://microsoft.com/devicelogin and enter the code.' -ForegroundColor Yellow
    Connect-ExchangeOnline @connectParams
    $connectedHere = $true
}

# ---------- Prompt for date range ----------
$maxLookback = (Get-Date).AddDays(-90)

if (-not $StartDate) {
    do {
        $raw = (Read-Host "`nStart date and time (e.g. 2026-07-29 11:00, max 90 days ago)").Trim()
        $parsed = [datetime]::MinValue
        $valid  = [datetime]::TryParse($raw, [ref]$parsed) -and $parsed -ge $maxLookback -and $parsed -lt (Get-Date)
        if (-not $valid) { Write-Host "  Enter a valid date within the last 90 days." -ForegroundColor Yellow }
    } while (-not $valid)
    $StartDate = $parsed
}

if (-not $EndDate) {
    do {
        $raw = (Read-Host "End date and time   (e.g. 2026-07-29 12:00)").Trim()
        $parsed = [datetime]::MinValue
        $valid  = [datetime]::TryParse($raw, [ref]$parsed) -and $parsed -gt $StartDate
        if (-not $valid) { Write-Host "  Enter a valid date after the start date." -ForegroundColor Yellow }
    } while (-not $valid)
    $EndDate = $parsed
}

if ($StartDate -lt $maxLookback) {
    Write-Error "StartDate cannot be more than 90 days ago (limit: $($maxLookback.ToString('yyyy-MM-dd')))."
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    exit 1
}

# ---------- Prompt for rule name ----------
if ([string]::IsNullOrWhiteSpace($RuleName)) {
    do {
        $RuleName = (Read-Host "Transport rule name to search for (partial match supported)").Trim()
        if ([string]::IsNullOrWhiteSpace($RuleName)) { Write-Host "  Rule name is required." -ForegroundColor Yellow }
    } while ([string]::IsNullOrWhiteSpace($RuleName))
}

# ---------- Prompt for status filter ----------
if (-not $StatusFilter) {
    Write-Host "`nStatus filter (reduces API calls for large date ranges):" -ForegroundColor Cyan
    Write-Host "  [1] All messages (slowest — checks every message)"
    Write-Host "  [2] Failed only  (fastest — use for reject/block rules)"
    Write-Host "  [3] Quarantined  (use for quarantine rules)"
    Write-Host "  [4] FilteredAsSpam"
    do {
        $sel = (Read-Host "Selection [1-4]").Trim()
    } while ($sel -notmatch '^[1-4]$')
    $StatusFilter = @{ '1' = 'All'; '2' = 'Failed'; '3' = 'Quarantined'; '4' = 'FilteredAsSpam' }[$sel]
}

Write-Host "`nSearch parameters:" -ForegroundColor Cyan
Write-Host "  Date range   : $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))"
Write-Host "  Rule name    : $RuleName"
Write-Host "  Status filter: $StatusFilter"

# ---------- Split range into <=10-day windows ----------
$windows = [System.Collections.Generic.List[PSCustomObject]]::new()
$cursor  = $StartDate
while ($cursor -lt $EndDate) {
    $windowEnd = if ($cursor.AddDays(10) -lt $EndDate) { $cursor.AddDays(10) } else { $EndDate }
    $windows.Add([PSCustomObject]@{ Start = $cursor; End = $windowEnd })
    $cursor = $windowEnd
}
Write-Host "  Date windows : $($windows.Count) x ≤10-day window(s)`n" -ForegroundColor Gray

# ---------- Collect messages via paginated trace ----------
Write-Host 'Step 1/2 — Retrieving messages from Message Trace...' -ForegroundColor Cyan
$allMessages = [System.Collections.Generic.List[object]]::new()

foreach ($window in $windows) {
    $page = 1
    do {
        $traceParams = @{
            StartDate = $window.Start
            EndDate   = $window.End
            PageSize  = 1000
            Page      = $page
        }
        if ($StatusFilter -ne 'All') { $traceParams['Status'] = $StatusFilter }

        $batch = Get-MessageTrace @traceParams -ErrorAction SilentlyContinue
        if ($batch -and $batch.Count -gt 0) {
            foreach ($msg in $batch) { $allMessages.Add($msg) }
            $page++
        }
    } while ($batch -and $batch.Count -eq 1000)
}

Write-Host "  Retrieved $($allMessages.Count) message(s) to inspect." -ForegroundColor Gray

if ($allMessages.Count -eq 0) {
    Write-Host 'No messages found in the specified date range and status filter.' -ForegroundColor Yellow
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    exit 0
}

# ---------- Check each message for the transport rule ----------
Write-Host "Step 2/2 — Checking message trace detail for rule '$RuleName'..." -ForegroundColor Cyan
Write-Host '  (This may take several minutes for large result sets.)' -ForegroundColor Gray

$results = [System.Collections.Generic.List[PSCustomObject]]::new()
$total   = $allMessages.Count
$index   = 0

foreach ($msg in $allMessages) {
    $index++
    Write-Progress -Activity "Checking message detail" `
        -Status "$index of $total — $($msg.SenderAddress) → $($msg.RecipientAddress)" `
        -PercentComplete (($index / $total) * 100)

    try {
        $details = Get-MessageTraceDetail `
            -MessageId        $msg.MessageId `
            -RecipientAddress $msg.RecipientAddress `
            -ErrorAction SilentlyContinue

        # Match any detail event that references the rule name
        $ruleMatch = $details | Where-Object { $_.Detail -like "*$RuleName*" }

        if ($ruleMatch) {
            $action = ($ruleMatch | Select-Object -First 1).Detail
            $results.Add([PSCustomObject]@{
                DateTime  = $msg.Received
                Sender    = $msg.SenderAddress
                Recipient = $msg.RecipientAddress
                Subject   = $msg.Subject
                Status    = $msg.Status
                RuleAction = $action
                MessageId = $msg.MessageId
            })
        }
    } catch {
        Write-Warning "Could not retrieve detail for message $($msg.MessageId): $_"
    }
}

Write-Progress -Activity "Checking message detail" -Completed

# ---------- Output ----------
if ($results.Count -gt 0) {
    $results | Sort-Object DateTime -Descending | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "`nFound $($results.Count) message(s) affected by rule '$RuleName'." -ForegroundColor Green
    Write-Host "Results saved to: $OutputCsv" -ForegroundColor Green

    Write-Host "`n--- Most Recent Matches ---" -ForegroundColor Cyan
    $results | Sort-Object DateTime -Descending | Select-Object -First 15 |
        Format-Table DateTime, Sender, Recipient, Subject, Status -AutoSize
} else {
    Write-Host "`nNo messages matched rule '$RuleName' in the specified date range and status filter." -ForegroundColor Yellow
    Write-Host "Tip: Try 'All' as the status filter if the rule silently deletes rather than rejects." -ForegroundColor Gray
}

if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
