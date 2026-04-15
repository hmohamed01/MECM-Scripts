# mecm-scripts

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)
[![Version](https://img.shields.io/badge/version-1.0.0-brightgreen.svg)](CHANGELOG.md)
[![Build Status](https://img.shields.io/badge/build-passing-success.svg)]()
[![PowerShell](https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white)](https://learn.microsoft.com/en-us/powershell/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white)](https://www.microsoft.com/en-us/windows)
[![ConfigMgr](https://img.shields.io/badge/MECM-Current%20Branch-0072C6)](https://learn.microsoft.com/en-us/mem/configmgr/)

PowerShell automation scripts for Microsoft Endpoint Configuration Manager (MECM / ConfigMgr / SCCM). Collections, reporting, and client operations — all built on a shared helper and consistent idempotent patterns so scripts are safe to re-run.

## Requirements

- Windows with the **MECM console** installed
- **PowerShell 5.1+**
- Account with appropriate MECM RBAC rights for the scripts you intend to run (collection create, remote WMI, etc.)

## Repository Layout

```
mecm-scripts/
├── Collections/                         Collection builders (data-driven hashtables + WQL)
│   ├── New-OSCollections.ps1            Win 10 / 11 / Server versions
│   ├── New-ClientHealthCollections.ps1  Stale / inactive / obsolete client triage
│   └── New-ServerRoleCollections.ps1    DC, SQL, IIS, DNS, DHCP, Hyper-V, Exchange, WSUS
│
├── Reporting/                           Read-only audit and export
│   ├── Get-EmptyCollections.ps1         Collections with 0 members
│   └── Export-CollectionInventory.ps1   Full collection audit to CSV
│
├── Operations/                          Bulk client actions
│   └── Invoke-ClientActionOnCollection.ps1   Trigger client actions across a collection
│
└── Common/
    └── Connect-CMSite.ps1               Shared module loader + PSDrive helper
```

## Usage

Run any script from a PowerShell session on a machine with the MECM console. Each script prompts for the site code:

```powershell
.\Collections\New-OSCollections.ps1
.\Collections\New-ClientHealthCollections.ps1
.\Collections\New-ServerRoleCollections.ps1

.\Reporting\Get-EmptyCollections.ps1 -OutputPath C:\Temp\empty.csv -IncludeUserCollections
.\Reporting\Export-CollectionInventory.ps1

.\Operations\Invoke-ClientActionOnCollection.ps1 `
    -CollectionName "Patch Ring - Pilot" `
    -Action MachinePolicy
```

## Design Conventions

All scripts follow the same patterns:

- **Fault-tolerant module loading** — resolves `$env:SMS_ADMIN_UI_PATH` first, falls back to `Import-Module ConfigurationManager`
- **PSDrive restore** — saves the original location before switching to `<SiteCode>:\`, restores it on exit (including error paths)
- **Idempotent create-or-skip** — existence checks before every object creation; `Created / Skipped / Errors` counter summary at the end
- **Declarative definitions** — collections and WQL queries live in arrays of hashtables, not imperative code, so the site-shaping decisions are reviewable in one place
- **Shared helper** — scripts dot-source `Common/Connect-CMSite.ps1` and call `Connect-CMSite` / `Disconnect-CMSite` rather than duplicating module loading

## WQL Notes for OS Collections

- `OperatingSystemNameandVersion LIKE '%Workstation 10.0%'` matches **both** Windows 10 and Windows 11 (both report NT 10.0). Combine with `Build LIKE '10.0.<build>%'` to differentiate feature updates.
- `OperatingSystemNameandVersion LIKE '%Server 10.0%'` matches Server 2016 / 2019 / 2022 / 2025 — also requires `Build` to differentiate.
- Server 2012 / 2012 R2 use NT 6.2 / 6.3 and don't need a Build filter.
- Server-role collections (`New-ServerRoleCollections.ps1`) detect roles via `SMS_G_System_SERVICE` joined on service names (NTDS, MSSQL%, W3SVC, etc.) — requires Win32_Service hardware inventory enabled (default).

## Collection Refresh Strategy

Collections are created with **weekly full evaluation** and **no incremental updates**. Enabling incremental updates across many query-based collections scales poorly on larger sites. If you need faster updates for specific collections, change them individually rather than as the default.

## Contributing

Match the existing patterns when adding scripts:

1. Dot-source `Common/Connect-CMSite.ps1` and wrap logic in `try / finally { Disconnect-CMSite }`
2. Use hashtable-driven definitions for anything repetitive
3. Check existence before creating; count Created / Skipped / Errors and print a summary
4. Keep WQL and configuration declarative and visible at the top of the script

## License

MIT — see [LICENSE](LICENSE).
