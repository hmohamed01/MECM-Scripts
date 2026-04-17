# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.3.0] - 2026-04-17

### Added
- `Collections/New-ClientHealthCollections.ps1` — 10 software update health collections under `Client Health\Software Update Health`: WUA service state, BITS service state, CcmEval health results, stale update scan (14/30 days), update enforcement failures, pending reboot, and all-deployments-compliant baseline
- `Operations/Invoke-CMUpdateHealthCheck.ps1` — software update and WUA patching health check script with CMTrace logging; validates both WUA layer (service, version, WSUS registration, SoftwareDistribution store, pending/failed updates via COM API) and ConfigMgr layer (SU client agent, scan status, SUP assignment, deployment compliance, cache usage)

## [1.2.0] - 2026-04-16

### Changed
- `Operations/Invoke-CMServerHealthCheck.ps1` — added ODBC Driver 18 readiness check (step 9) for ConfigMgr 2503 upgrade; detects pre-2503 sites and validates ODBC Driver 18 >= 18.4.1.1

## [1.1.0] - 2026-04-15

### Added
- `Collections/New-Win11ReadinessCollections.ps1` — 13 Windows 11 upgrade readiness collections (TPM 2.0, Secure Boot, UEFI, RAM, disk, composite hardware-ready)
- `Operations/Invoke-CMClientHealthCheck.ps1` — read-only client health check with CMTrace logging (services, WMI, policy, inventory, cache, MP connectivity, disk, pending reboot)
- `Operations/Invoke-CMServerHealthCheck.ps1` — read-only site server health check with CMTrace logging (SMS services, SMS Provider, component status, inbox backlogs, disk, SQL connectivity, event log errors)
- `Operations/Repair-WMISafely.ps1` — non-destructive WMI repository repair (salvage, DLL re-registration, MOF recompilation)

### Removed
- MIT license

## [1.0.0] - 2026-04-15

### Added
- `Common/Connect-CMSite.ps1` — shared helper for module loading and PSDrive switching with location restore
- `Collections/New-OSCollections.ps1` — Windows 10, Windows 11, and Windows Server version collections
- `Collections/New-ClientHealthCollections.ps1` — client health triage collections (no client, obsolete, inactive, stale heartbeat, stale inventory, stale policy)
- `Collections/New-ServerRoleCollections.ps1` — server role collections via SMS_G_System_SERVICE joins (DC, DNS, DHCP, SQL, IIS, Hyper-V, Exchange, WSUS, MECM)
- `Reporting/Get-EmptyCollections.ps1` — report collections with zero members
- `Reporting/Export-CollectionInventory.ps1` — full collection audit export to CSV, including incremental-update usage flag
- `Operations/Invoke-ClientActionOnCollection.ps1` — trigger client actions across collection members via WMI SMS_Client.TriggerSchedule
