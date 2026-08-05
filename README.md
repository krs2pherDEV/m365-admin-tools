# M365 Admin Tools

A growing collection of PowerShell scripts for Microsoft 365 administration. Scripts are organized by workload. Each script is self-contained and follows a consistent pattern: module guard, connection handling, parameterized inputs, and CSV export.

---

## Structure

```
m365-admin-tools/
  exchange/       Exchange Online administration scripts
```

---

## Prerequisites

All scripts require **ExchangeOnlineManagement v3+** unless noted otherwise.

```powershell
Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser -Force
```

---

## Scripts

### Exchange

| Script | Description |
|---|---|
| [Get-TransportRuleAudit.ps1](exchange/Get-TransportRuleAudit.ps1) | Searches the Unified Audit Log for transport rule create/modify/delete/enable/disable events and exports parsed results to CSV |
| [Get-TransportRuleDrops.ps1](exchange/Get-TransportRuleDrops.ps1) | Finds action hits for a specific transport rule during the last 10 days using the server-filtered transport rule report; exports sender, recipient, subject, and action |

`Get-TransportRuleDrops.ps1` writes each completed result page to `<OutputName>.partial.csv`. After a successful search, it renames that file to the requested output path. If Exchange returns an error or the run is interrupted, completed pages remain in the partial file for recovery.

---

## Common Parameters

All scripts share these parameters:

| Parameter | Default | Description |
|---|---|---|
| `-AdminUPN` | _(prompt)_ | Admin UPN for authentication |
| `-OutputCsv` | `.\<ScriptName>_<timestamp>.csv` | Output file path |
| `-TenantEnvironment` | `USGovGCC` | `Commercial`, `USGovGCC`, or `USGovGCCHigh` |

---

## License

MIT
