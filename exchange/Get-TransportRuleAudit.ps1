<#
.SYNOPSIS
    Searches the Unified Audit Log for Exchange transport rule changes and exports results to CSV.
.DESCRIPTION
    Connects to Exchange Online and queries the Unified Audit Log for create, modify, delete,
    enable, and disable events on transport (mail flow) rules. Parses the AuditData JSON blob
    on each event to extract the rule name, who made the change, and what parameters were set.
    Handles UAL pagination automatically to retrieve more than 5,000 records.
.PARAMETER StartDate
    Start of the search window. Defaults to 30 days ago.
.PARAMETER EndDate
    End of the search window. Defaults to now.
.PARAMETER RuleName
    Optional. Filter results to a specific transport rule name (partial match).
.PARAMETER Operations
    Optional. Filter to specific operations. Defaults to all transport rule operations.
    Valid values: New-TransportRule, Set-TransportRule, Remove-TransportRule,
                  Enable-TransportRule, Disable-TransportRule
.PARAMETER AdminUPN
    Admin UPN used to connect to Exchange Online if no active session exists.
    Requires Audit Logs or View-Only Audit Logs RBAC role in addition to Exchange admin.
.PARAMETER OutputCsv
    Path for the output CSV file. Defaults to .\TransportRuleAudit_<timestamp>.csv
.PARAMETER TenantEnvironment
    Exchange Online environment. Defaults to USGovGCC.
.EXAMPLE
    .\Get-TransportRuleAudit.ps1
.EXAMPLE
    .\Get-TransportRuleAudit.ps1 -StartDate (Get-Date).AddDays(-90) -RuleName "Block External"
.EXAMPLE
    .\Get-TransportRuleAudit.ps1 -Operations "Remove-TransportRule","Disable-TransportRule"
#>

#Requires -Modules ExchangeOnlineManagement

[CmdletBinding()]
param(
    [Parameter()]
    [datetime]$StartDate = (Get-Date).AddDays(-30),

    [Parameter()]
    [datetime]$EndDate = (Get-Date),

    [Parameter()]
    [string]$RuleName,

    [Parameter()]
    [ValidateSet(
        'New-TransportRule',
        'Set-TransportRule',
        'Remove-TransportRule',
        'Enable-TransportRule',
        'Disable-TransportRule'
    )]
    [string[]]$Operations,

    [Parameter()]
    [string]$AdminUPN,

    [Parameter()]
    [string]$OutputCsv = ".\TransportRuleAudit_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv",

    [Parameter()]
    [ValidateSet('Commercial', 'USGovGCC', 'USGovGCCHigh')]
    [string]$TenantEnvironment = 'USGovGCC'
)

# ---------- Module guard ----------
if (-not (Get-Module -Name ExchangeOnlineManagement -ListAvailable | Where-Object { $_.Version -ge '3.0.0' })) {
    Write-Error 'ExchangeOnlineManagement v3+ is required. Install it with: Install-Module ExchangeOnlineManagement -Scope CurrentUser'
    exit 1
}

# ---------- Validate date range ----------
if ($StartDate -ge $EndDate) {
    Write-Error "StartDate ($StartDate) must be before EndDate ($EndDate)."
    exit 1
}

$maxRetention = (Get-Date).AddDays(-180)
if ($StartDate -lt $maxRetention) {
    Write-Warning "StartDate is older than 180 days. UAL retention is typically 90 days (standard) or 1 year (E5). Results may be empty."
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
        if ([string]::IsNullOrWhiteSpace($AdminUPN)) {
            Write-Error 'No admin UPN provided. Exiting.'
            exit 1
        }
    }

    $connectParams = @{
        UserPrincipalName = $AdminUPN
        ShowBanner        = $false
        Device            = $true
    }
    if ($TenantEnvironment -eq 'USGovGCCHigh') {
        $connectParams['ExchangeEnvironmentName'] = 'O365USGovGCCHigh'
    }

    Write-Host "Connecting to Exchange Online as $AdminUPN ..." -ForegroundColor Cyan
    Write-Host 'A device login prompt will appear — go to https://microsoft.com/devicelogin and enter the code.' -ForegroundColor Yellow
    Connect-ExchangeOnline @connectParams
    $connectedHere = $true
}

# ---------- Prompt for rule name filter if not supplied ----------
if ([string]::IsNullOrWhiteSpace($RuleName)) {
    $input = (Read-Host "`nEnter a transport rule name to filter by (or press Enter to search all)").Trim()
    if (-not [string]::IsNullOrWhiteSpace($input)) {
        $RuleName = $input
        Write-Host "Filtering results to rules matching: '$RuleName'" -ForegroundColor Cyan
    } else {
        Write-Host "No filter applied — returning all transport rule events." -ForegroundColor Gray
    }
}

# ---------- Build search parameters ----------
# Transport rule events are logged under ExchangeAdmin, filtered by Operation name
$allTransportOps = @(
    'New-TransportRule',
    'Set-TransportRule',
    'Remove-TransportRule',
    'Enable-TransportRule',
    'Disable-TransportRule'
)

$searchParams = @{
    RecordType     = 'ExchangeAdmin'
    Operations     = if ($Operations -and $Operations.Count -gt 0) { $Operations } else { $allTransportOps }
    StartDate      = $StartDate
    EndDate        = $EndDate
    ResultSize     = 5000
    SessionCommand = 'ReturnLargeSet'
}

# ---------- Paginated UAL search ----------
Write-Host "`nSearching Unified Audit Log from $($StartDate.ToString('yyyy-MM-dd')) to $($EndDate.ToString('yyyy-MM-dd'))..." -ForegroundColor Cyan

$sessionId  = [System.Guid]::NewGuid().ToString()
$rawResults = [System.Collections.Generic.List[object]]::new()

do {
    $batch = Search-UnifiedAuditLog @searchParams -SessionId $sessionId -ErrorAction Stop
    if ($batch -and $batch.Count -gt 0) {
        foreach ($r in $batch) { $rawResults.Add($r) }
        Write-Host "  Retrieved $($rawResults.Count) record(s) so far..." -ForegroundColor Gray
    }
} while ($batch -and $batch.Count -eq 5000)

Write-Host "Total raw records retrieved: $($rawResults.Count)" -ForegroundColor Cyan

if ($rawResults.Count -eq 0) {
    Write-Host 'No transport rule audit events found in the specified date range.' -ForegroundColor Yellow
    if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
    exit 0
}

# ---------- Parse AuditData JSON ----------
Write-Host 'Parsing audit records...' -ForegroundColor Cyan
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($record in $rawResults) {
    try {
        $audit = $record.AuditData | ConvertFrom-Json

        $ruleName_field   = $audit.ObjectId
        $operation        = $audit.Operation
        $modifiedBy       = $audit.UserId
        $clientIP         = $audit.ClientIP
        $resultStatus     = $audit.ResultStatus

        # Parameters: flatten Name=Value pairs for Set-TransportRule changes
        $parameters = if ($audit.Parameters) {
            ($audit.Parameters | ForEach-Object { "$($_.Name)=$($_.Value)" }) -join '; '
        } else { '' }

        # Apply optional rule name filter
        if ($RuleName -and $ruleName_field -notlike "*$RuleName*") { continue }

        $results.Add([PSCustomObject]@{
            Timestamp   = $record.CreationDate
            Operation   = $operation
            RuleName    = $ruleName_field
            ModifiedBy  = $modifiedBy
            ClientIP    = $clientIP
            ResultStatus = $resultStatus
            Parameters  = $parameters
        })
    } catch {
        Write-Warning "Failed to parse audit record (CreationDate: $($record.CreationDate)): $_"
    }
}

# ---------- Output ----------
if ($results.Count -gt 0) {
    $results | Sort-Object Timestamp -Descending | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Host "`nFound $($results.Count) transport rule event(s)." -ForegroundColor Green
    Write-Host "Results saved to: $OutputCsv" -ForegroundColor Green

    Write-Host "`n--- Summary by Operation ---" -ForegroundColor Cyan
    $results | Group-Object Operation | Select-Object Name, Count | Sort-Object Count -Descending | Format-Table -AutoSize

    Write-Host '--- Most Recent Events ---' -ForegroundColor Cyan
    $results | Sort-Object Timestamp -Descending | Select-Object -First 10 |
        Format-Table Timestamp, Operation, RuleName, ModifiedBy -AutoSize
} else {
    Write-Host "No events matched the specified filters." -ForegroundColor Yellow
}

if ($connectedHere) { Disconnect-ExchangeOnline -Confirm:$false }
