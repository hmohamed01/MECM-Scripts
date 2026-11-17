# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Purpose

PowerShell automation scripts for Microsoft Endpoint Configuration Manager (MECM/ConfigMgr/SCCM) administration. Scripts run on a machine with the MECM console installed and operate against a site by connecting to the CM PSDrive (`<SiteCode>:\`).

Scripts are organized by purpose:

- `Collections/` — collection builders (OS versions, client health, server roles, hardware, Win11 readiness)
- `Reporting/` — read-only collection audits and CSV exports
- `Operations/` — client actions, health checks, and WMI repair
- `Common/` — `Connect-CMSite.ps1`, the shared module loader and PSDrive helper

The health check and repair scripts in `Operations/` run locally on the target machine and query WMI/CIM directly. They do not use the CM PSDrive or `Connect-CMSite`.

## Running Scripts

These scripts are intended to run on Windows with the MECM console installed. Execute directly from a PowerShell 5.1+ session:

```powershell
.\Collections\New-OSCollections.ps1
```

The collection scripts prompt for the site code interactively. There are no build, lint, or test commands — scripts are executed in place against a live MECM site.

## Script Architecture Conventions

Scripts that connect to a site follow the same pattern; match it when adding new ones. `Collections/New-ServerRoleCollections.ps1` is the reference to copy.

### 1. Connect through the shared helper

Scripts dot-source `Common/Connect-CMSite.ps1`, call `Connect-CMSite -SiteCode $SiteCode`, and wrap all logic in `try { ... } finally { Disconnect-CMSite }`. Do not inline module loading in new scripts. `New-OSCollections.ps1` predates the helper and still inlines it — don't copy that part.

### 2. ConfigurationManager module loading (fault-tolerant)

`Connect-CMSite` loads the ConfigMgr module by first resolving `$env:SMS_ADMIN_UI_PATH` (set by the console install) and transforming it to the `.psd1` path, then falling back to `Import-Module ConfigurationManager`. Preserve this fallback chain — environments with the console installed but the env var unset rely on the fallback.

### 3. PSDrive switching with restore

`Connect-CMSite` saves `Get-Location` before `Set-Location "$($SiteCode):\"` and `Disconnect-CMSite` restores it. The `finally` block is what guarantees the restore on error paths — the CM PSDrive breaks most non-CM cmdlets while active.

### 4. Idempotent object creation

Before creating any MECM object (collection, folder, etc.), check existence first and skip if present. The pattern uses `Get-CMDeviceCollection -ErrorAction SilentlyContinue` + a `Created/Skipped/Errors` counter summary at the end. New scripts should follow this — MECM object creation is not naturally idempotent and re-runs without this check will fail or duplicate.

Create-or-skip keys on the collection name, so a changed WQL never reaches a collection that already exists, and a renamed collection is created alongside the old one. When changing a query or name in an existing script, say so in the changelog entry with the manual step for sites that already ran it.

### 5. Declarative object definitions

Collections are defined as an array of hashtables with `Name`, `Folder`, `Comment`, `WQL` keys, then iterated. When adding new MECM object types, prefer this data-driven style over imperative per-object code — it keeps the WQL/config reviewable in one place. Where many collections share a join, a small query-builder function (e.g. `New-ServiceRoleQuery`, `New-ChassisTypeQuery`) supplies the `WQL` value.

## Documentation

Every script change updates `CHANGELOG.md` (Keep a Changelog, SemVer) and the version badge in `README.md`. New scripts are also added to the README layout tree and usage block.

## WQL Conventions for OS Collections

When extending `New-OSCollections.ps1` or writing similar queries:

- `OperatingSystemNameandVersion LIKE '%Workstation 10.0%'` matches **both** Windows 10 and Windows 11 (both report NT 10.0). You must combine with `Build LIKE '10.0.<buildnumber>%'` to differentiate specific feature updates.
- `OperatingSystemNameandVersion LIKE '%Server 10.0%'` matches Server 2016/2019/2022/2025 — again requires `Build` to differentiate.
- Legacy Server 2012/2012 R2 use NT 6.2 / 6.3 and don't need a Build filter.
- Feature update build number mapping is in the `$Collections` array comments — keep it current when Microsoft ships new versions.

## WQL Conventions for Hardware Collections

`Collections/New-HardwareCollections.ps1` creates collections under `Hardware\Manufacturers`, `Hardware\Processor Architectures`, and `Hardware\Chassis Types`. When extending it or writing similar queries:

- Queries are built by two helpers, `New-ComputerSystemQuery -Filter` (joins `SMS_G_System_COMPUTER_SYSTEM`) and `New-ChassisTypeQuery -ChassisTypeCodes` (joins `SMS_G_System_SYSTEM_ENCLOSURE`). Add a hashtable that calls a helper rather than writing the join inline.
- Collection names carry a group prefix: `Manufacturer - `, `Architecture - `, `Chassis - `.
- `Manufacturer LIKE 'Microsoft%'` matches **both** Surface devices and Hyper-V/Azure guests. Combine with `Model` (`LIKE 'Surface%'` or `= 'Virtual Machine'`) to tell them apart.
- HP hardware reports `HP`, `Hewlett-Packard`, or `HPE`, so the HP collection matches `HP%` OR `Hewlett%`.
- Architecture uses `SystemType` (`x64%`, `x86%`, `ARM64%`), which is the platform Windows runs on. 32-bit Windows on a 64-bit processor reports x86.
- `ChassisTypes` is stored as a string — quote the SMBIOS codes in the `IN` list. The helper does this.
- Chassis queries exclude virtual machines with `SMS_R_System.IsVirtualMachine = 'False'`. VMs report a chassis code of their own (Hyper-V guests report 3, Desktop).
- Each SMBIOS chassis code belongs to exactly one chassis collection. The code table and grouping are in the `$Collections` array comments. Docking codes (12, 18, 21) and 2-in-1 codes (31, 32) are grouped with laptops by decision — don't move them back to tablets.
- These queries have not been validated against a live site. If a query is rejected, the collection is created without a membership rule and skipped on every re-run.

## Collection Refresh Strategy

The current script creates collections with:
- `RefreshType Periodic` + weekly `New-CMSchedule -RecurInterval Days -RecurCount 7`
- **No incremental updates** — intentional, to avoid continuous collection evaluation load across 20+ collections

If modifying this, note that enabling incremental updates on many query-based collections scales poorly on larger sites. Prefer weekly or daily full evaluation unless there is a specific operational reason.
