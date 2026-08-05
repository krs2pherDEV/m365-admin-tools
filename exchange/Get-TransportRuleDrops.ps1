<#
.SYNOPSIS
    Finds emails dropped or rejected by a specific Exchange transport rule and exports
    sender, recipient, subject, date, and action to CSV.
.DESCRIPTION
    Connects to Exchange Online, prompts for a date range and transport rule name, then
    queries the detailed transport rule report, filtered by the selected rule, to identify
    messages affected by the rule. Splits the date range into one-day windows and paginates
    each window automatically.
    Results are exported to a timestamped CSV.

    NOTE: The detailed transport rule report supports a maximum lookback of 10 days.
.PARAMETER StartDate
    Start of the search window. Must be within the last 90 days.
.PARAMETER EndDate
    End of the search window.
.PARAMETER RuleName
    The transport rule name to search for (partial match supported).
.PARAMETER AdminUPN
    Admin UPN used to connect if no active Exchange Online session exists.
.PARAMETER OutputCsv
    Path for the completed output CSV. Results are appended to a sibling .partial.csv
    file during the search, then renamed to this path after successful completion.
.PARAMETER TenantEnvironment
    Exchange Online environment. Defaults to USGovGCC.
.EXAMPLE
    .\Get-TransportRuleDrops.ps1
.EXAMPLE
    .\Get-TransportRuleDrops.ps1 -StartDate (Get-Date).AddDays(-7) -RuleName "Block External Fwd"
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
$maxLookback = (Get-Date).AddDays(-10)

if (-not $StartDate) {
    do {
        $raw = (Read-Host "`nStart date and time (e.g. 2026-07-29 11:00, max 10 days ago)").Trim()
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
    Write-Error "StartDate cannot be more than 10 days ago (transport rule report limit: $($maxLookback.ToString('yyyy-MM-dd')))."
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

# ---------- Resolve partial rule name to one exact rule ----------
try {
    $matchingRules = @(Get-TransportRule -ErrorAction Stop | Where-Object { $_.Name -like "*$RuleName*" })
} catch {
    Write-Error "Could not read transport rules: $($_.Exception.Message)"
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    exit 1
}

if ($matchingRules.Count -eq 0) {
    Write-Error "No transport rule matched '$RuleName'."
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    exit 1
}
if ($matchingRules.Count -gt 1) {
    Write-Host "More than one transport rule matched '$RuleName':" -ForegroundColor Yellow
    $matchingRules | Sort-Object Name | ForEach-Object { Write-Host "  - $($_.Name)" }
    Write-Error 'Use a more specific rule name and run the script again.'
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    exit 1
}
$RuleName = $matchingRules[0].Name

Write-Host "`nSearch parameters:" -ForegroundColor Cyan
Write-Host "  Date range: $($StartDate.ToString('yyyy-MM-dd HH:mm')) to $($EndDate.ToString('yyyy-MM-dd HH:mm'))"
Write-Host "  Rule name : $RuleName"

# ---------- Split range into one-day windows ----------
$windows = [System.Collections.Generic.List[PSCustomObject]]::new()
$cursor  = $StartDate
while ($cursor -lt $EndDate) {
    $windowEnd = if ($cursor.AddDays(1) -lt $EndDate) { $cursor.AddDays(1) } else { $EndDate }
    $windows.Add([PSCustomObject]@{ Start = $cursor; End = $windowEnd })
    $cursor = $windowEnd
}
Write-Host "  Date windows: $($windows.Count) x one-day window(s)`n" -ForegroundColor Gray

function Get-FirstPropertyValue {
    param(
        [object]$InputObject,
        [string[]]$Names
    )
    foreach ($name in $Names) {
        $property = $InputObject.PSObject.Properties[$name]
        if ($null -ne $property -and $null -ne $property.Value) { return $property.Value }
    }
    return $null
}

# ---------- Retrieve only action hits for the selected rule ----------
Write-Host "Retrieving transport rule action hits for '$RuleName'..." -ForegroundColor Cyan

$outputDirectory = Split-Path -Parent ([System.IO.Path]::GetFullPath($OutputCsv))
if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
    Write-Error "Output directory does not exist: $outputDirectory"
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    exit 1
}

$outputExtension = [System.IO.Path]::GetExtension($OutputCsv)
$partialCsv = if ($outputExtension -eq '.csv') {
    Join-Path $outputDirectory "$([System.IO.Path]::GetFileNameWithoutExtension($OutputCsv)).partial.csv"
} else {
    "$([System.IO.Path]::GetFullPath($OutputCsv)).partial.csv"
}
Remove-Item -LiteralPath $partialCsv -Force -ErrorAction SilentlyContinue

$csvInitialized = $false
$resultCount = 0
$recentResults = [System.Collections.Generic.List[PSCustomObject]]::new()
$knownResults = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
$knownResultOrder = [System.Collections.Generic.Queue[string]]::new()
$deduplicationCacheSize = 10000

foreach ($window in $windows) {
    $page = 1
    $windowCount = 0
    do {
        try {
            $batch = @(Get-MailDetailTransportRuleReport `
                -StartDate $window.Start `
                -EndDate $window.End `
                -TransportRule $RuleName `
                -EventType TransportRuleActionHits `
                -Page $page `
                -PageSize 5000 `
                -ErrorAction Stop)
        } catch {
            $partialMessage = if ($csvInitialized) { " Partial results remain at '$partialCsv'." } else { '' }
            if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
            throw "Transport rule report failed for $($window.Start.ToString('yyyy-MM-dd HH:mm')) - $($window.End.ToString('yyyy-MM-dd HH:mm')), page $page`: $($_.Exception.Message).$partialMessage"
        }

        $pageResults = [System.Collections.Generic.List[PSCustomObject]]::new()
        foreach ($item in $batch) {
            $messageTraceId = Get-FirstPropertyValue $item @('MessageTraceId', 'MessageTraceID')
            $messageId = Get-FirstPropertyValue $item @('MessageId', 'MessageID')
            $recipient = Get-FirstPropertyValue $item @('RecipientAddress', 'Recipient')
            $action = Get-FirstPropertyValue $item @('Action', 'EventType')
            $key = "$messageTraceId|$messageId|$recipient|$action"
            if ($knownResults.Add($key)) {
                $knownResultOrder.Enqueue($key)
                if ($knownResultOrder.Count -gt $deduplicationCacheSize) {
                    [void]$knownResults.Remove($knownResultOrder.Dequeue())
                }
            $pageResults.Add([PSCustomObject]@{
                    Source     = 'TransportRuleReport'
                    DateTime   = Get-FirstPropertyValue $item @('Date', 'Received', 'ReceivedTime')
                    Sender     = Get-FirstPropertyValue $item @('SenderAddress', 'Sender')
                    Recipient  = $recipient
                    Subject    = Get-FirstPropertyValue $item @('Subject')
                    Status     = Get-FirstPropertyValue $item @('EventType', 'Status')
                    RuleAction = $action
                    MessageId  = $messageId
                    MessageTraceId = $messageTraceId
                })
            }
        }

        if ($pageResults.Count -gt 0) {
            if ($csvInitialized) {
                $pageResults | Export-Csv -LiteralPath $partialCsv -NoTypeInformation -Encoding UTF8 -Append
            } else {
                $pageResults | Export-Csv -LiteralPath $partialCsv -NoTypeInformation -Encoding UTF8
                $csvInitialized = $true
                Write-Host "  Writing incremental results to: $partialCsv" -ForegroundColor Gray
            }
            $resultCount += $pageResults.Count

            foreach ($result in $pageResults) { $recentResults.Add($result) }
            if ($recentResults.Count -gt 15) {
                $latest = @($recentResults | Sort-Object DateTime -Descending | Select-Object -First 15)
                $recentResults.Clear()
                foreach ($result in $latest) { $recentResults.Add($result) }
            }
        }

        $windowCount += $batch.Count
        if ($page -eq 1000 -and $batch.Count -eq 5000) {
            if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
            throw "The report reached its 5,000,000-row limit for a one-day window. Partial results remain at '$partialCsv'. Rerun with a shorter date range."
        }
        $page++
    } while ($batch.Count -eq 5000)

    Write-Host "  $($window.Start.ToString('yyyy-MM-dd HH:mm')) - $($window.End.ToString('yyyy-MM-dd HH:mm')): $windowCount action hit(s)" -ForegroundColor Gray
}

if ($resultCount -gt 0) {
    Move-Item -LiteralPath $partialCsv -Destination $OutputCsv -Force
    Write-Host "`nFound $resultCount message(s) affected by rule '$RuleName'." -ForegroundColor Green
    Write-Host "Results saved to: $OutputCsv" -ForegroundColor Green

    Write-Host "`n--- Most Recent Matches ---" -ForegroundColor Cyan
    $recentResults | Sort-Object DateTime -Descending |
        Format-Table Source, DateTime, Sender, Recipient, Subject, Status -AutoSize
} else {
    Write-Host "`nNo action hits matched rule '$RuleName' in the specified date range." -ForegroundColor Yellow
    Write-Host 'Tip: Reporting data can take time to appear. Retry later if the message was very recent.' -ForegroundColor Gray
}

if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
