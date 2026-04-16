#Requires -Version 5.1
<#
.SYNOPSIS
    Creates device collections for Windows 11 upgrade readiness assessment in MECM.

.DESCRIPTION
    Builds a set of device collections that surface Windows 11 hardware readiness
    across the Windows 10 device population:

        - Overview: All Windows 10 devices and already-upgraded Windows 11 devices
        - TPM 2.0 present vs missing/unknown
        - Secure Boot enabled vs not enabled
        - UEFI firmware vs Legacy BIOS
        - RAM sufficient (4 GB+) vs insufficient
        - System disk sufficient (64 GB+) vs insufficient
        - Composite: all hardware requirements met

    Hardware readiness collections are scoped to Windows 10 devices only
    (the upgrade candidate population).

    Collections are placed under Device Collections\Windows 11 Upgrade Readiness
    and refreshed weekly. Existing collections are skipped (idempotent).

.NOTES
    Run on a machine with the MECM console installed.
    Requires permissions to create collections.

    Hardware inventory classes SMS_G_System_TPM, SMS_G_System_FIRMWARE,
    SMS_G_System_X86_PC_MEMORY, and SMS_G_System_LOGICAL_DISK must be enabled
    in Client Settings > Hardware Inventory for these queries to return results.
    They are enabled by default in MECM 2107+.

    CPU generation compatibility is not checked via WQL — there is no clean
    inventory field for CPU generation. Validate CPU compatibility separately
    using the PC Health Check app or Microsoft's published processor list.
#>

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"

# ─── Connect to CM Site ───────────────────────────────────────────────
. "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
Connect-CMSite -SiteCode $SiteCode

try {
    $FolderRoot = "$($SiteCode.ToUpper()):\DeviceCollection"

    # ─── Create Console Folder Structure ──────────────────────────────
    $Folders = @("Windows 11 Upgrade Readiness")

    foreach ($Folder in $Folders) {
        $FullPath = Join-Path $FolderRoot $Folder
        if (-not (Test-Path $FullPath)) {
            $ParentPath = Split-Path $FullPath -Parent
            $LeafName   = Split-Path $FullPath -Leaf
            New-Item -Path $ParentPath -Name $LeafName -ItemType Folder | Out-Null
            Write-Host "Created folder: Device Collections\$Folder" -ForegroundColor Cyan
        }
        else {
            Write-Host "Folder exists:  Device Collections\$Folder" -ForegroundColor DarkGray
        }
    }

    # ─── Define Collections ───────────────────────────────────────────
    # WQL scoping notes:
    #   Win10 + Win11 both report OperatingSystemNameandVersion = "...Workstation 10.0"
    #   Win10 builds: 10.0.10240 – 10.0.19045  (all match Build LIKE '10.0.1%')
    #   Win11 builds: 10.0.22000+               (all match Build LIKE '10.0.2%')
    #
    # Hardware inventory classes used:
    #   SMS_G_System_TPM              — SpecVersion for TPM version
    #   SMS_G_System_FIRMWARE         — SecureBoot and UEFI booleans
    #   SMS_G_System_X86_PC_MEMORY    — TotalPhysicalMemory in KB
    #   SMS_G_System_LOGICAL_DISK     — Size in MB, DeviceID for drive letter
    #
    # Win11 minimum requirements checked:
    #   TPM 2.0, Secure Boot, UEFI, 4 GB RAM (4194304 KB), 64 GB disk (65536 MB)

    $Collections = @(

        # ── Overview ──────────────────────────────────────────────────
        @{
            Name    = "Win11 Readiness - Already on Windows 11"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Devices already running Windows 11 (any version)"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.2%'"
        },
        @{
            Name    = "Win11 Readiness - All Windows 10 Devices"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "All Windows 10 devices — the upgrade candidate population"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.1%'"
        },

        # ── TPM ───────────────────────────────────────────────────────
        @{
            Name    = "Win11 Readiness - TPM 2.0 Present"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices with TPM 2.0 — meets Win11 TPM requirement"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_TPM ON SMS_G_System_TPM.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_TPM.SpecVersion LIKE '2.0%'
"@
        },
        @{
            Name    = "Win11 Readiness - TPM 2.0 Missing"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices with no TPM 2.0 (absent, disabled, or older version)"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
LEFT JOIN SMS_G_System_TPM ON SMS_G_System_TPM.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND (SMS_G_System_TPM.SpecVersion IS NULL OR SMS_G_System_TPM.SpecVersion NOT LIKE '2.0%')
"@
        },

        # ── Secure Boot ──────────────────────────────────────────────
        @{
            Name    = "Win11 Readiness - Secure Boot Enabled"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices with Secure Boot enabled — meets Win11 requirement"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_FIRMWARE ON SMS_G_System_FIRMWARE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_FIRMWARE.SecureBoot = 1
"@
        },
        @{
            Name    = "Win11 Readiness - Secure Boot Not Enabled"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices without Secure Boot — may need BIOS config or lack capability"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
LEFT JOIN SMS_G_System_FIRMWARE ON SMS_G_System_FIRMWARE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND (SMS_G_System_FIRMWARE.SecureBoot IS NULL OR SMS_G_System_FIRMWARE.SecureBoot = 0)
"@
        },

        # ── UEFI ─────────────────────────────────────────────────────
        @{
            Name    = "Win11 Readiness - UEFI Firmware"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices booting via UEFI — meets Win11 firmware requirement"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_FIRMWARE ON SMS_G_System_FIRMWARE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_FIRMWARE.UEFI = 1
"@
        },
        @{
            Name    = "Win11 Readiness - Legacy BIOS"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices on Legacy BIOS — requires UEFI conversion before Win11 upgrade"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
LEFT JOIN SMS_G_System_FIRMWARE ON SMS_G_System_FIRMWARE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND (SMS_G_System_FIRMWARE.UEFI IS NULL OR SMS_G_System_FIRMWARE.UEFI = 0)
"@
        },

        # ── RAM ──────────────────────────────────────────────────────
        @{
            Name    = "Win11 Readiness - RAM 4GB or More"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices with 4+ GB RAM — meets Win11 memory requirement"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_X86_PC_MEMORY ON SMS_G_System_X86_PC_MEMORY.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_X86_PC_MEMORY.TotalPhysicalMemory >= 4194304
"@
        },
        @{
            Name    = "Win11 Readiness - RAM Below 4GB"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices with less than 4 GB RAM — hardware replacement candidate"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_X86_PC_MEMORY ON SMS_G_System_X86_PC_MEMORY.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_X86_PC_MEMORY.TotalPhysicalMemory < 4194304
"@
        },

        # ── Disk Space ───────────────────────────────────────────────
        @{
            Name    = "Win11 Readiness - Disk 64GB or More"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices with C: drive 64+ GB — meets Win11 storage requirement"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_LOGICAL_DISK ON SMS_G_System_LOGICAL_DISK.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_LOGICAL_DISK.DeviceID = 'C:'
AND SMS_G_System_LOGICAL_DISK.Size >= 65536
"@
        },
        @{
            Name    = "Win11 Readiness - Disk Below 64GB"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices with C: drive under 64 GB — hardware replacement candidate"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_LOGICAL_DISK ON SMS_G_System_LOGICAL_DISK.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_LOGICAL_DISK.DeviceID = 'C:'
AND SMS_G_System_LOGICAL_DISK.Size < 65536
"@
        },

        # ── Composite: All Hardware Requirements Met ─────────────────
        @{
            Name    = "Win11 Readiness - Hardware Ready"
            Folder  = "Windows 11 Upgrade Readiness"
            Comment = "Win10 devices meeting ALL Win11 hardware requirements: TPM 2.0 + Secure Boot + UEFI + 4GB RAM + 64GB disk"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_TPM ON SMS_G_System_TPM.ResourceID = SMS_R_System.ResourceID
INNER JOIN SMS_G_System_FIRMWARE ON SMS_G_System_FIRMWARE.ResourceID = SMS_R_System.ResourceID
INNER JOIN SMS_G_System_X86_PC_MEMORY ON SMS_G_System_X86_PC_MEMORY.ResourceID = SMS_R_System.ResourceID
INNER JOIN SMS_G_System_LOGICAL_DISK ON SMS_G_System_LOGICAL_DISK.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%'
AND SMS_R_System.Build LIKE '10.0.1%'
AND SMS_G_System_TPM.SpecVersion LIKE '2.0%'
AND SMS_G_System_FIRMWARE.SecureBoot = 1
AND SMS_G_System_FIRMWARE.UEFI = 1
AND SMS_G_System_X86_PC_MEMORY.TotalPhysicalMemory >= 4194304
AND SMS_G_System_LOGICAL_DISK.DeviceID = 'C:'
AND SMS_G_System_LOGICAL_DISK.Size >= 65536
"@
        }
    )

    # ─── Weekly Refresh Schedule ──────────────────────────────────────
    $Schedule = New-CMSchedule -RecurInterval Days -RecurCount 7

    # ─── Create Collections ───────────────────────────────────────────
    $Created = 0
    $Skipped = 0
    $Errors  = 0

    foreach ($Col in $Collections) {
        $CollectionName = $Col.Name

        $Existing = Get-CMDeviceCollection -Name $CollectionName -ErrorAction SilentlyContinue
        if ($Existing) {
            Write-Host "SKIP:    $CollectionName (already exists)" -ForegroundColor DarkGray
            $Skipped++
            continue
        }

        try {
            $NewCollection = New-CMDeviceCollection `
                -Name                   $CollectionName `
                -LimitingCollectionName "All Systems" `
                -RefreshType            Periodic `
                -RefreshSchedule        $Schedule `
                -Comment                $Col.Comment

            Add-CMDeviceCollectionQueryMembershipRule `
                -CollectionName  $CollectionName `
                -RuleName        $CollectionName `
                -QueryExpression $Col.WQL

            $TargetFolder = Join-Path $FolderRoot $Col.Folder
            $NewCollection | Move-CMObject -FolderPath $TargetFolder

            Write-Host "CREATED: $CollectionName" -ForegroundColor Green
            $Created++
        }
        catch {
            Write-Host "ERROR:   $CollectionName - $($_.Exception.Message)" -ForegroundColor Red
            $Errors++
        }
    }

    # ─── Summary ──────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "═══════════════════════════════════════════" -ForegroundColor White
    Write-Host "  Collections Created: $Created" -ForegroundColor Green
    Write-Host "  Collections Skipped: $Skipped" -ForegroundColor DarkGray
    Write-Host "  Errors:              $Errors" -ForegroundColor $(if ($Errors -gt 0) { "Red" } else { "DarkGray" })
    Write-Host "═══════════════════════════════════════════" -ForegroundColor White
    Write-Host ""
}
finally {
    Disconnect-CMSite
}
