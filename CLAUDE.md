# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repository Purpose

PowerShell automation scripts for Microsoft Endpoint Configuration Manager (MECM/ConfigMgr/SCCM) administration. Scripts run on a machine with the MECM console installed and operate against a site by connecting to the CM PSDrive (`<SiteCode>:\`).

Scripts are organized by MECM object category. Currently: `Collections/`.

## Running Scripts

These scripts are intended to run on Windows with the MECM console installed. Execute directly from a PowerShell 5.1+ session:

```powershell
.\Collections\New-OSCollections.ps1
```

The script will prompt for the site code interactively. There are no build, lint, or test commands — scripts are executed in place against a live MECM site.

## Script Architecture Conventions

All scripts in this repo follow the same pattern; match it when adding new ones.

### 1. ConfigurationManager module loading (fault-tolerant)

Scripts load the ConfigMgr module by first resolving `$env:SMS_ADMIN_UI_PATH` (set by the console install) and transforming it to the `.psd1` path, then falling back to `Import-Module ConfigurationManager`. Preserve this fallback chain — environments with the console installed but the env var unset rely on the fallback.

### 2. PSDrive switching with restore

Scripts save `Get-Location` before `Set-Location "$($SiteCode):\"` and restore it at the end. Every new script must restore the original location, including on error paths, because the CM PSDrive breaks most non-CM cmdlets while active.

### 3. Idempotent object creation

Before creating any MECM object (collection, folder, etc.), check existence first and skip if present. The pattern in `New-OSCollections.ps1` uses `Get-CMDeviceCollection -ErrorAction SilentlyContinue` + a `Created/Skipped/Errors` counter summary at the end. New scripts should follow this — MECM object creation is not naturally idempotent and re-runs without this check will fail or duplicate.

### 4. Declarative object definitions

Collections are defined as an array of hashtables with `Name`, `Folder`, `Comment`, `WQL` keys, then iterated. When adding new MECM object types, prefer this data-driven style over imperative per-object code — it keeps the WQL/config reviewable in one place.

## WQL Conventions for OS Collections

When extending `New-OSCollections.ps1` or writing similar queries:

- `OperatingSystemNameandVersion LIKE '%Workstation 10.0%'` matches **both** Windows 10 and Windows 11 (both report NT 10.0). You must combine with `Build LIKE '10.0.<buildnumber>%'` to differentiate specific feature updates.
- `OperatingSystemNameandVersion LIKE '%Server 10.0%'` matches Server 2016/2019/2022/2025 — again requires `Build` to differentiate.
- Legacy Server 2012/2012 R2 use NT 6.2 / 6.3 and don't need a Build filter.
- Feature update build number mapping is in the `$Collections` array comments — keep it current when Microsoft ships new versions.

## Collection Refresh Strategy

The current script creates collections with:
- `RefreshType Periodic` + weekly `New-CMSchedule -RecurInterval Days -RecurCount 7`
- **No incremental updates** — intentional, to avoid continuous collection evaluation load across 20+ collections

If modifying this, note that enabling incremental updates on many query-based collections scales poorly on larger sites. Prefer weekly or daily full evaluation unless there is a specific operational reason.
