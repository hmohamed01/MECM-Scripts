# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.4.2] - 2026-11-09

### Fixed
- `Collections/New-ClientHealthCollections.ps1` — collections queried `SMS_CH_ClientSummary`, which is not a documented SMS Provider class. On sites where the provider rejected the query, the collection was created with no membership rule and skipped on every re-run. `ClientActiveStatus`, `LastHardwareScan`, and `LastPolicyRequest` now come from the documented `SMS_CombinedDeviceResources` class. Affects `Clients - Inactive (per ClientSummary)`, `Clients - No Hardware Inventory 30+ Days`, `Clients - No Policy Request 14+ Days`, `Clients - All Active Healthy`, `SU Health - All Deployments Compliant`, and the two `HW Health - Active Client` collections

### Removed
- `Collections/New-ClientHealthCollections.ps1` — `SU Health - No Update Scan 14+ Days` and `30+ Days`, and the 7-day scan condition in `SU Health - All Deployments Compliant`. They used `LastSWUpdateScanTime`, which has no documented WMI equivalent; update-scan time is only exposed in the `v_UpdateScanStatus` SQL view

## [1.4.1] - 2026-11-09

### Fixed
- `Collections/New-ClientHealthCollections.ps1` — `Clients - No Heartbeat 30/60/90+ Days` filtered on `SMS_R_System.AgentTime`, an array with one timestamp per discovery agent. WQL matches an array if any element matches, so any stale agent entry (e.g. client registration) flagged clients that were heartbeating normally. Now filters on `SMS_CombinedDeviceResources.LastDDR`. Update the membership rules on sites where these collections already exist

## [1.4.0] - 2026-11-09

### Added
- `Collections/New-OSCollections.ps1` — Windows 11 25H2 (26200), 26H1 (28000), and 26H2 (26300) collections. 26H1 ships only on new Snapdragon X2 devices. There is no 25H1 release
- `Collections/New-ClientHealthCollections.ps1` — hardware health collections under `Client Health\Hardware Health`: low C: free space, disk SMART predicted failure, CcmEval WMI check failed/remediated, active clients with stale or missing hardware inventory (WMI suspect), no reboot in 30+ days, and workstation RAM below 8 GB

### Fixed
- `Collections/New-ClientHealthCollections.ps1` — `SU Health - CcmEval Health Check Failed` queried `Result = 7` (Evaluation Succeeded) and so returned healthy clients; now queries `Result IN (3, 4, 5)`. Re-running the script skips existing collections, so update the membership rule on sites where it was already created

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
