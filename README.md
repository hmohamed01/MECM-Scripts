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
│   ├── New-ServerRoleCollections.ps1    DC, SQL, IIS, DNS, DHCP, Hyper-V, Exchange, WSUS
│   └── New-Win11ReadinessCollections.ps1  Win11 upgrade readiness (HW checks)
│
├── Reporting/                           Read-only audit and export
│   ├── Get-EmptyCollections.ps1         Collections with 0 members
│   └── Export-CollectionInventory.ps1   Full collection audit to CSV
│
├── Operations/                          Client/server operations and health checks
│   ├── Invoke-ClientActionOnCollection.ps1   Trigger client actions across a collection
│   ├── Invoke-CMClientHealthCheck.ps1        Read-only client health check (CMTrace log)
│   ├── Invoke-CMServerHealthCheck.ps1        Read-only site server health check (CMTrace log)
│   └── Repair-WMISafely.ps1                  Non-destructive WMI repository repair
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
.\Collections\New-Win11ReadinessCollections.ps1

.\Reporting\Get-EmptyCollections.ps1 -OutputPath C:\Temp\empty.csv -IncludeUserCollections
.\Reporting\Export-CollectionInventory.ps1

.\Operations\Invoke-ClientActionOnCollection.ps1 `
    -CollectionName "Patch Ring - Pilot" `
    -Action MachinePolicy

# Health checks — run locally on the target machine as Administrator
.\Operations\Invoke-CMClientHealthCheck.ps1
.\Operations\Invoke-CMClientHealthCheck.ps1 -ManagementPoint "cm01.contoso.com"

.\Operations\Invoke-CMServerHealthCheck.ps1
.\Operations\Invoke-CMServerHealthCheck.ps1 -SiteCode "PS1" -SiteServer "cm01.contoso.com"

# WMI repair — non-destructive, run on machines with WMI issues
.\Operations\Repair-WMISafely.ps1
```

## Design Conventions

All scripts follow the same patterns:

- **Fault-tolerant module loading** — resolves `$env:SMS_ADMIN_UI_PATH` first, falls back to `Import-Module ConfigurationManager`
- **PSDrive restore** — saves the original location before switching to `<SiteCode>:\`, restores it on exit (including error paths)
- **Idempotent create-or-skip** — existence checks before every object creation; `Created / Skipped / Errors` counter summary at the end
- **Declarative definitions** — collections and WQL queries live in arrays of hashtables, not imperative code, so the site-shaping decisions are reviewable in one place
- **Shared helper** — scripts dot-source `Common/Connect-CMSite.ps1` and call `Connect-CMSite` / `Disconnect-CMSite` rather than duplicating module loading

## WQL Notes for OS Collections

- Windows 10 and Windows 11 both report `OperatingSystemNameandVersion LIKE '%Workstation 10.0%'` (both are NT 10.0), so each OS-version collection also filters on `Build LIKE '10.0.<build>%'` to pin the exact feature update — e.g. Windows 10 22H2 uses `10.0.19045%`, Windows 11 24H2 uses `10.0.26100%`.
- Server 2016 / 2019 / 2022 / 2025 all share `OperatingSystemNameandVersion LIKE '%Server 10.0%'`, so each server collection likewise includes a `Build` filter to pin the specific release.
- Server 2012 and 2012 R2 use NT 6.2 and 6.3 respectively, so those collections match on `OperatingSystemNameandVersion` alone with no Build filter needed.
- Server-role collections (`New-ServerRoleCollections.ps1`) detect roles via `SMS_G_System_SERVICE` joined on service names (NTDS, MSSQL%, W3SVC, etc.) — requires Win32_Service hardware inventory enabled (default).

## Hardware Inventory Prerequisites for Win11 Readiness

`New-Win11ReadinessCollections.ps1` queries hardware inventory classes that must be enabled in **Client Settings > Hardware Inventory** for the collections to return results:

| Inventory Class | WMI Class | Used For |
|---|---|---|
| TPM | `Win32_Tpm` | TPM 2.0 detection (`SMS_G_System_TPM.SpecVersion`) |
| Firmware | `Win32_Firmware` | Secure Boot and UEFI detection (`SMS_G_System_FIRMWARE`) |
| Physical Memory | `Win32_PhysicalMemoryArray` | RAM capacity (`SMS_G_System_X86_PC_MEMORY`) |
| Logical Disk | `Win32_LogicalDisk` | System disk size (`SMS_G_System_LOGICAL_DISK`) |

These classes are enabled by default in **MECM 2107+**. On older sites, enable them under **Administration > Client Settings > Default Client Settings > Hardware Inventory > Set Classes**. Devices must complete a hardware inventory cycle after enabling before they appear in the readiness collections.

**Note:** CPU generation compatibility is not checked — there is no reliable inventory field for this. Validate CPU support separately using Microsoft's [PC Health Check](https://aka.ms/GetPCHealthCheckApp) app or the [published processor list](https://learn.microsoft.com/en-us/windows-hardware/design/minimum/supported/windows-11-supported-intel-processors).

## Operations Scripts

The health check and repair scripts in `Operations/` are designed to run **locally on the target machine** as Administrator. They do not use the CM PSDrive or `Connect-CMSite` — they query WMI/CIM directly.

All three produce **CMTrace-compatible logs** that can be opened with [CMTrace](https://learn.microsoft.com/en-us/mem/configmgr/core/support/cmtrace) for color-coded severity filtering:

| Script | Default Log Path | Purpose |
|---|---|---|
| `Invoke-CMClientHealthCheck.ps1` | `C:\Windows\CCM\Logs\ClientHealthCheck.log` | Read-only client validation: services, WMI, policy, inventory, cache, MP connectivity, disk, pending reboot |
| `Invoke-CMServerHealthCheck.ps1` | `C:\Windows\Logs\CMServerHealthCheck.log` | Read-only site server validation: SMS services, SMS Provider, component status, inbox backlogs, disk, SQL connectivity, event log errors |
| `Repair-WMISafely.ps1` | `C:\Windows\Logs\WMIRepair.log` | Non-destructive WMI repair using `winmgmt /salvagerepository` (never `/resetrepository`), DLL re-registration, MOF recompilation |

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
