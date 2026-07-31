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
Write-Host 'Step 1/2 - Retrieving messages from Message Trace...' -ForegroundColor Cyan
$allMessages   = [System.Collections.Generic.List[object]]::new()
$cappedWindows = [System.Collections.Generic.List[PSCustomObject]]::new()

function Invoke-TraceWindow {
    param(
        [PSCustomObject]$Window,
        [string]$StatusFilter,
        [System.Collections.Generic.List[object]]$MessageList
    )
    $page = 1
    $windowCount = 0
    do {
        $traceParams = @{
            StartDate = $Window.Start
            EndDate   = $Window.End
            PageSize  = 1000
            Page      = $page
        }
        if ($StatusFilter -ne 'All') { $traceParams['Status'] = $StatusFilter }

        $batch = Get-MessageTrace @traceParams -ErrorAction SilentlyContinue
        if ($batch -and $batch.Count -gt 0) {
            foreach ($msg in $batch) { $MessageList.Add($msg) }
            $windowCount += $batch.Count
            $page++
        }
    } while ($batch -and $batch.Count -eq 1000)

    return $windowCount
}

foreach ($window in $windows) {
    $count = Invoke-TraceWindow -Window $window -StatusFilter $StatusFilter -MessageList $allMessages
    Write-Host "  Window $($window.Start.ToString('yyyy-MM-dd HH:mm')) - $($window.End.ToString('yyyy-MM-dd HH:mm')): $count message(s)" -ForegroundColor Gray

    # Exact multiple of 1000 means we may have hit Exchange Online's result cap
    if ($count -gt 0 -and $count % 1000 -eq 0) {
        Write-Host "    ^ Ended on exactly $count - possible result cap. Will offer retry." -ForegroundColor Yellow
        $cappedWindows.Add($window)
    }
}

Write-Host "  Retrieved $($allMessages.Count) message(s) to inspect." -ForegroundColor Gray

# ---------- Offer retry with smaller splits for potentially capped windows ----------
if ($cappedWindows.Count -gt 0) {
    Write-Host "`nWARNING: $($cappedWindows.Count) window(s) ended on an exact multiple of 1,000 and may have been capped by Exchange Online." -ForegroundColor Yellow
    $retry = (Read-Host "Re-query those window(s) with smaller splits to recover potentially missing messages? [Y/N]").Trim().ToUpper()

    if ($retry -eq 'Y') {
        # Build a HashSet of already-collected MessageIds to deduplicate
        $knownIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($msg in $allMessages) { [void]$knownIds.Add($msg.MessageId) }

        $retryMessages = [System.Collections.Generic.List[object]]::new()

        foreach ($cappedWindow in $cappedWindows) {
            $windowSpan = ($cappedWindow.End - $cappedWindow.Start).TotalHours

            # Adaptive split size: >24h -> 1-hour chunks; <=24h -> 10-minute chunks
            if ($windowSpan -gt 24) {
                $splitMinutes = 60
                $splitLabel   = '1-hour'
            } else {
                $splitMinutes = 10
                $splitLabel   = '10-minute'
            }

            Write-Host "  Re-querying $($cappedWindow.Start.ToString('yyyy-MM-dd HH:mm')) - $($cappedWindow.End.ToString('yyyy-MM-dd HH:mm')) in $splitLabel splits..." -ForegroundColor Cyan
            $subCursor = $cappedWindow.Start

            while ($subCursor -lt $cappedWindow.End) {
                $subEnd    = $subCursor.AddMinutes($splitMinutes)
                if ($subEnd -gt $cappedWindow.End) { $subEnd = $cappedWindow.End }
                $subWindow = [PSCustomObject]@{ Start = $subCursor; End = $subEnd }
                $subList   = [System.Collections.Generic.List[object]]::new()

                $subCount = Invoke-TraceWindow -Window $subWindow -StatusFilter $StatusFilter -MessageList $subList
                Write-Host "    $($subCursor.ToString('yyyy-MM-dd HH:mm')) - $($subEnd.ToString('yyyy-MM-dd HH:mm')): $subCount message(s)" -ForegroundColor Gray

                if ($subCount % 1000 -eq 0 -and $subCount -gt 0) {
                    Write-Host "    ^ Still capped at $subCount. Volume is very high for this window." -ForegroundColor Yellow
                }

                foreach ($msg in $subList) {
                    if ($knownIds.Add($msg.MessageId)) {
                        $retryMessages.Add($msg)
                    }
                }
                $subCursor = $subEnd
            }
        }

        if ($retryMessages.Count -gt 0) {
            Write-Host "  Recovered $($retryMessages.Count) additional unique message(s) from retry." -ForegroundColor Green
            foreach ($msg in $retryMessages) { $allMessages.Add($msg) }
        } else {
            Write-Host "  No additional messages found in retry (original results were likely complete)." -ForegroundColor Gray
        }
    }
}


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
                Source    = 'MessageTrace'
                DateTime  = $msg.Received
                Sender    = $msg.SenderAddress
                Recipient = $msg.RecipientAddress
                Subject   = $msg.Subject
                Status    = $msg.Status
                RuleAction = $action
                MessageId = $msg.MessageId
            })
        }

        # Throttle: EXO allows ~3 detail calls/sec sustained; 350ms keeps us under the limit
        # and avoids the GetResponseHeader crash in the EXO module's 429 retry handler
        Start-Sleep -Milliseconds 350

    } catch {
        Write-Warning "Could not retrieve detail for message $($msg.MessageId): $_"
    }
}

Write-Progress -Activity "Checking message detail" -Completed

# ---------- Quarantine lookup (faster than message trace — no propagation delay) ----------
Write-Host "`nStep 3/3 — Checking quarantine for transport rule matches..." -ForegroundColor Cyan
try {
    $quarantineHits = Get-QuarantineMessage -StartReceivedDate $StartDate `
                                            -EndReceivedDate $EndDate `
                                            -QuarantineTypes TransportRule `
                                            -PageSize 1000 `
                                            -ErrorAction Stop |
        Where-Object { $_.PolicyName -like "*$RuleName*" }

    if ($quarantineHits -and $quarantineHits.Count -gt 0) {
        Write-Host "  Found $($quarantineHits.Count) quarantined message(s) matching rule '$RuleName'." -ForegroundColor Gray
        foreach ($qMsg in $quarantineHits) {
            # Avoid double-counting messages already found in message trace
            $alreadyFound = $results | Where-Object {
                $_.Sender -eq $qMsg.SenderAddress -and
                $_.Recipient -eq $qMsg.RecipientAddress -and
                [math]::Abs(($_.DateTime - $qMsg.ReceivedTime).TotalMinutes) -lt 2
            }
            if (-not $alreadyFound) {
                $results.Add([PSCustomObject]@{
                    Source     = 'Quarantine'
                    DateTime   = $qMsg.ReceivedTime
                    Sender     = $qMsg.SenderAddress
                    Recipient  = $qMsg.RecipientAddress
                    Subject    = $qMsg.Subject
                    Status     = 'Quarantined'
                    RuleAction = $qMsg.PolicyName
                    MessageId  = $qMsg.Identity
                })
            }
        }
    } else {
        Write-Host '  No quarantined messages matched.' -ForegroundColor Gray
    }
} catch {
    Write-Warning "Quarantine lookup failed: $_"
}
if ($results.Count -gt 0) {
    $results | Sort-Object DateTime -Descending | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "`nFound $($results.Count) message(s) affected by rule '$RuleName'." -ForegroundColor Green
    Write-Host "Results saved to: $OutputCsv" -ForegroundColor Green

    Write-Host "`n--- Most Recent Matches ---" -ForegroundColor Cyan
    $results | Sort-Object DateTime -Descending | Select-Object -First 15 |
        Format-Table Source, DateTime, Sender, Recipient, Subject, Status -AutoSize
} else {
    Write-Host "`nNo messages matched rule '$RuleName' in the specified date range and status filter." -ForegroundColor Yellow
    Write-Host "Tip: Try 'All' as the status filter if the rule silently deletes rather than rejects." -ForegroundColor Gray
    Write-Host "Tip: If the message was very recent, message trace may not have propagated yet (5-30 min delay)." -ForegroundColor Gray
}

if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
