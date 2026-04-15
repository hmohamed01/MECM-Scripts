# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.0.0] - 2026-04-15

### Added
- `Common/Connect-CMSite.ps1` — shared helper for module loading and PSDrive switching with location restore
- `Collections/New-OSCollections.ps1` — Windows 10, Windows 11, and Windows Server version collections
- `Collections/New-ClientHealthCollections.ps1` — client health triage collections (no client, obsolete, inactive, stale heartbeat, stale inventory, stale policy)
- `Collections/New-ServerRoleCollections.ps1` — server role collections via SMS_G_System_SERVICE joins (DC, DNS, DHCP, SQL, IIS, Hyper-V, Exchange, WSUS, MECM)
- `Reporting/Get-EmptyCollections.ps1` — report collections with zero members
- `Reporting/Export-CollectionInventory.ps1` — full collection audit export to CSV, including incremental-update usage flag
- `Operations/Invoke-ClientActionOnCollection.ps1` — trigger client actions across collection members via WMI SMS_Client.TriggerSchedule
