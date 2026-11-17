#Requires -Version 5.1
<#
.SYNOPSIS
    Creates device collections grouping systems by hardware: manufacturer,
    processor architecture, and chassis type.

.DESCRIPTION
    Builds query-based collections from hardware inventory in three folders
    under Device Collections\Hardware:

        Manufacturers             SMS_G_System_COMPUTER_SYSTEM.Manufacturer / Model
        Processor Architectures   SMS_G_System_COMPUTER_SYSTEM.SystemType
        Chassis Types             SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes

    Chassis collections group the SMBIOS chassis type codes into form factors
    (desktops, laptops, tablets, servers, ...) and exclude virtual machines,
    which report a chassis code of their own (Hyper-V guests report Desktop).

    Collections are created with a weekly full refresh and no incremental
    updates.

.NOTES
    Requires hardware inventory for Win32_ComputerSystem and
    Win32_SystemEnclosure. Devices appear only after reporting a hardware
    inventory cycle.

    SystemType reflects the platform Windows is running on, so 32-bit Windows
    installed on a 64-bit processor lands in the x86 collection.
#>

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"

# ─── Connect to CM Site ───────────────────────────────────────────────
. "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
Connect-CMSite -SiteCode $SiteCode

try {
    $FolderRoot = "$($SiteCode.ToUpper()):\DeviceCollection"

    # ─── Create Console Folder Structure ──────────────────────────────
    $Folders = @(
        "Hardware",
        "Hardware\Manufacturers",
        "Hardware\Processor Architectures",
        "Hardware\Chassis Types"
    )

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

    # ─── WQL Helpers ──────────────────────────────────────────────────
    $SystemColumns = "SMS_R_System.ResourceID, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client"

    function New-ComputerSystemQuery {
        param(
            [Parameter(Mandatory)] [string]$Filter
        )
        @"
SELECT $SystemColumns
FROM SMS_R_System INNER JOIN SMS_G_System_COMPUTER_SYSTEM
ON SMS_G_System_COMPUTER_SYSTEM.ResourceID = SMS_R_System.ResourceID
WHERE $Filter
"@
    }

    # ChassisTypes is stored as a string, so the codes are quoted.
    # A docked laptop can report a second enclosure and match two collections.
    function New-ChassisTypeQuery {
        param(
            [Parameter(Mandatory)] [int[]]$ChassisTypeCodes
        )
        $QuotedCodes = ($ChassisTypeCodes | ForEach-Object { "'$_'" }) -join ", "
        @"
SELECT $SystemColumns
FROM SMS_R_System INNER JOIN SMS_G_System_SYSTEM_ENCLOSURE
ON SMS_G_System_SYSTEM_ENCLOSURE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_G_System_SYSTEM_ENCLOSURE.ChassisTypes IN ($QuotedCodes)
AND SMS_R_System.IsVirtualMachine = 'False'
"@
    }

    # ─── Define Collections ───────────────────────────────────────────
    # SMBIOS chassis type codes used below:
    #    3 Desktop            4 Low Profile Desktop   5 Pizza Box         6 Mini Tower
    #    7 Tower              8 Portable              9 Laptop           10 Notebook
    #   11 Hand Held         12 Docking Station      13 All in One       14 Sub Notebook
    #   15 Space-Saving      16 Lunch Box            17 Main Server      18 Expansion Chassis
    #   21 Peripheral        23 Rack Mount           25 Multi-System     28 Blade
    #   29 Blade Enclosure   30 Tablet               31 Convertible      32 Detachable
    #   35 Mini PC           36 Stick PC
    # 12, 18, and 21 are grouped with laptops: that is what a docked laptop reports.
    # 31 and 32 (2-in-1s) are grouped with laptops so they share laptop targeting.

    $Collections = @(

        # ── Manufacturers ─────────────────────────────────────────────
        @{
            Name    = "Manufacturer - Dell"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by Dell"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Dell%'"
        },
        @{
            Name    = "Manufacturer - HP"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by HP, Hewlett-Packard, or HPE"
            WQL     = New-ComputerSystemQuery -Filter "(SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'HP%' OR SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Hewlett%')"
        },
        @{
            Name    = "Manufacturer - Lenovo"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by Lenovo"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Lenovo%'"
        },
        @{
            Name    = "Manufacturer - Microsoft Surface"
            Folder  = "Hardware\Manufacturers"
            Comment = "Microsoft Surface devices (Microsoft Corporation with a Surface model, which leaves out Hyper-V guests)"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Microsoft%' AND SMS_G_System_COMPUTER_SYSTEM.Model LIKE 'Surface%'"
        },
        @{
            Name    = "Manufacturer - Acer"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by Acer"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Acer%'"
        },
        @{
            Name    = "Manufacturer - ASUS"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by ASUS (ASUSTeK)"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'ASUS%'"
        },
        @{
            Name    = "Manufacturer - Fujitsu"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by Fujitsu"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Fujitsu%'"
        },
        @{
            Name    = "Manufacturer - Panasonic"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by Panasonic"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Panasonic%'"
        },
        @{
            Name    = "Manufacturer - Samsung"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by Samsung"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Samsung%'"
        },
        @{
            Name    = "Manufacturer - Dynabook / Toshiba"
            Folder  = "Hardware\Manufacturers"
            Comment = "Devices manufactured by Toshiba or Dynabook (Toshiba's PC business since 2019)"
            WQL     = New-ComputerSystemQuery -Filter "(SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Toshiba%' OR SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Dynabook%')"
        },
        @{
            Name    = "Manufacturer - VMware Virtual Machines"
            Folder  = "Hardware\Manufacturers"
            Comment = "VMware guests (Manufacturer VMware, Inc.)"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'VMware%'"
        },
        @{
            Name    = "Manufacturer - Microsoft Hyper-V Virtual Machines"
            Folder  = "Hardware\Manufacturers"
            Comment = "Hyper-V and Azure guests (Microsoft Corporation with model Virtual Machine)"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.Manufacturer LIKE 'Microsoft%' AND SMS_G_System_COMPUTER_SYSTEM.Model = 'Virtual Machine'"
        },

        # ── Processor Architectures ───────────────────────────────────
        @{
            Name    = "Architecture - x64"
            Folder  = "Hardware\Processor Architectures"
            Comment = "64-bit Intel/AMD systems (SystemType x64-based PC)"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.SystemType LIKE 'x64%'"
        },
        @{
            Name    = "Architecture - x86 (32-bit)"
            Folder  = "Hardware\Processor Architectures"
            Comment = "32-bit Intel/AMD systems (SystemType x86-based PC), including 32-bit Windows on 64-bit processors"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.SystemType LIKE 'x86%'"
        },
        @{
            Name    = "Architecture - ARM64"
            Folder  = "Hardware\Processor Architectures"
            Comment = "ARM64 systems such as Snapdragon devices (SystemType ARM64-based PC)"
            WQL     = New-ComputerSystemQuery -Filter "SMS_G_System_COMPUTER_SYSTEM.SystemType LIKE 'ARM64%'"
        },

        # ── Chassis Types ─────────────────────────────────────────────
        @{
            Name    = "Chassis - Desktops"
            Folder  = "Hardware\Chassis Types"
            Comment = "Physical desktops and towers (chassis types 3, 4, 5, 6, 7, 15, 16)"
            WQL     = New-ChassisTypeQuery -ChassisTypeCodes 3, 4, 5, 6, 7, 15, 16
        },
        @{
            Name    = "Chassis - All-in-One PCs"
            Folder  = "Hardware\Chassis Types"
            Comment = "All-in-one PCs (chassis type 13)"
            WQL     = New-ChassisTypeQuery -ChassisTypeCodes 13
        },
        @{
            Name    = "Chassis - Mini and Stick PCs"
            Folder  = "Hardware\Chassis Types"
            Comment = "Mini PCs and stick PCs (chassis types 35, 36)"
            WQL     = New-ChassisTypeQuery -ChassisTypeCodes 35, 36
        },
        @{
            Name    = "Chassis - Laptops"
            Folder  = "Hardware\Chassis Types"
            Comment = "Laptops, notebooks, and 2-in-1s (convertibles and detachables), docked or undocked (chassis types 8, 9, 10, 12, 14, 18, 21, 31, 32)"
            WQL     = New-ChassisTypeQuery -ChassisTypeCodes 8, 9, 10, 12, 14, 18, 21, 31, 32
        },
        @{
            Name    = "Chassis - Tablets"
            Folder  = "Hardware\Chassis Types"
            Comment = "Hand-helds and tablets (chassis types 11, 30)"
            WQL     = New-ChassisTypeQuery -ChassisTypeCodes 11, 30
        },
        @{
            Name    = "Chassis - Servers"
            Folder  = "Hardware\Chassis Types"
            Comment = "Physical server chassis: tower server, rack mount, multi-system, blade (chassis types 17, 23, 25, 28, 29)"
            WQL     = New-ChassisTypeQuery -ChassisTypeCodes 17, 23, 25, 28, 29
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
