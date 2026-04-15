#Requires -Version 5.1
<#
.SYNOPSIS
    Creates query-based device collections for every Windows 10, Windows 11,
    and Windows Server version in MECM.

.DESCRIPTION
    Loads the Configuration Manager module, prompts for a site code, creates
    console folder structure under Device Collections, and builds one query-based
    collection per OS version. Each collection uses a WQL query combining
    OperatingSystemNameandVersion and Build to accurately target a specific version.

    Collections are created with a weekly full evaluation schedule and no
    incremental updates to avoid unnecessary collection evaluation load.

.NOTES
    Run this script on a machine with the MECM console installed.
    The script must be run as a user with permissions to create collections.
#>

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"
if ([string]::IsNullOrWhiteSpace($SiteCode)) {
    Write-Host "ERROR: Site code cannot be empty." -ForegroundColor Red
    exit 1
}
$SiteCode = $SiteCode.Trim().ToUpper()

# ─── Load Configuration Manager Module ────────────────────────────────
try {
    $CMModulePath = $env:SMS_ADMIN_UI_PATH
    if ($CMModulePath) {
        $CMModulePath = $CMModulePath.Replace("\bin\i386", "\bin\ConfigurationManager.psd1")
        if (Test-Path $CMModulePath) {
            Import-Module $CMModulePath -Force
            Write-Host "ConfigMgr module loaded via environment variable path." -ForegroundColor Green
        }
        else {
            throw "Module path from environment variable not found"
        }
    }
    else {
        throw "SMS_ADMIN_UI_PATH not set"
    }
}
catch {
    Write-Host "Falling back to Import-Module ConfigurationManager..." -ForegroundColor Yellow
    try {
        Import-Module ConfigurationManager -Force
        Write-Host "ConfigMgr module loaded via Import-Module." -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Failed to load ConfigMgr module. Is the MECM console installed?" -ForegroundColor Red
        exit 1
    }
}

# Switch to the CM drive
$OriginalLocation = Get-Location
Set-Location "$($SiteCode):\"

# ─── Create Console Folder Structure ──────────────────────────────────
$FolderRoot = "$($SiteCode):\DeviceCollection"

$Folders = @(
    "Operating Systems",
    "Operating Systems\Windows 10",
    "Operating Systems\Windows 11",
    "Operating Systems\Windows Server"
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

# ─── Define Collections ───────────────────────────────────────────────
# WQL note:
#   OperatingSystemNameandVersion = "Microsoft Windows NT Workstation 10.0" for Win 10 AND Win 11
#   OperatingSystemNameandVersion = "Microsoft Windows NT Server 10.0" for Server 2016, 2019, 2022, 2025
#   The Build field (e.g., "10.0.19045") is required to differentiate versions.

$Collections = @(

    # ── Windows 10 ────────────────────────────────────────────────────
    @{
        Name    = "Windows 10 - 1507"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1507 (Build 10240)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.10240%'"
    },
    @{
        Name    = "Windows 10 - 1511"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1511 (Build 10586)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.10586%'"
    },
    @{
        Name    = "Windows 10 - 1607"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1607 (Build 14393)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.14393%'"
    },
    @{
        Name    = "Windows 10 - 1703"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1703 (Build 15063)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.15063%'"
    },
    @{
        Name    = "Windows 10 - 1709"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1709 (Build 16299)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.16299%'"
    },
    @{
        Name    = "Windows 10 - 1803"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1803 (Build 17134)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.17134%'"
    },
    @{
        Name    = "Windows 10 - 1809"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1809 (Build 17763)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.17763%'"
    },
    @{
        Name    = "Windows 10 - 1903"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1903 (Build 18362)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.18362%'"
    },
    @{
        Name    = "Windows 10 - 1909"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 1909 (Build 18363)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.18363%'"
    },
    @{
        Name    = "Windows 10 - 2004"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 2004 (Build 19041)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.19041%'"
    },
    @{
        Name    = "Windows 10 - 20H2"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 20H2 (Build 19042)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.19042%'"
    },
    @{
        Name    = "Windows 10 - 21H1"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 21H1 (Build 19043)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.19043%'"
    },
    @{
        Name    = "Windows 10 - 21H2"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 21H2 (Build 19044)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.19044%'"
    },
    @{
        Name    = "Windows 10 - 22H2"
        Folder  = "Operating Systems\Windows 10"
        Comment = "Windows 10 22H2 (Build 19045)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.19045%'"
    },

    # ── Windows 11 ────────────────────────────────────────────────────
    @{
        Name    = "Windows 11 - 21H2"
        Folder  = "Operating Systems\Windows 11"
        Comment = "Windows 11 21H2 (Build 22000)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.22000%'"
    },
    @{
        Name    = "Windows 11 - 22H2"
        Folder  = "Operating Systems\Windows 11"
        Comment = "Windows 11 22H2 (Build 22621)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.22621%'"
    },
    @{
        Name    = "Windows 11 - 23H2"
        Folder  = "Operating Systems\Windows 11"
        Comment = "Windows 11 23H2 (Build 22631)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.22631%'"
    },
    @{
        Name    = "Windows 11 - 24H2"
        Folder  = "Operating Systems\Windows 11"
        Comment = "Windows 11 24H2 (Build 26100)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Workstation 10.0%' AND SMS_R_System.Build LIKE '10.0.26100%'"
    },

    # ── Windows Server ────────────────────────────────────────────────
    @{
        Name    = "Windows Server 2012"
        Folder  = "Operating Systems\Windows Server"
        Comment = "Windows Server 2012 (NT 6.2)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Server 6.2%'"
    },
    @{
        Name    = "Windows Server 2012 R2"
        Folder  = "Operating Systems\Windows Server"
        Comment = "Windows Server 2012 R2 (NT 6.3)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Server 6.3%'"
    },
    @{
        Name    = "Windows Server 2016"
        Folder  = "Operating Systems\Windows Server"
        Comment = "Windows Server 2016 (Build 14393)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Server 10.0%' AND SMS_R_System.Build LIKE '10.0.14393%'"
    },
    @{
        Name    = "Windows Server 2019"
        Folder  = "Operating Systems\Windows Server"
        Comment = "Windows Server 2019 (Build 17763)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Server 10.0%' AND SMS_R_System.Build LIKE '10.0.17763%'"
    },
    @{
        Name    = "Windows Server 2022"
        Folder  = "Operating Systems\Windows Server"
        Comment = "Windows Server 2022 (Build 20348)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Server 10.0%' AND SMS_R_System.Build LIKE '10.0.20348%'"
    },
    @{
        Name    = "Windows Server 2025"
        Folder  = "Operating Systems\Windows Server"
        Comment = "Windows Server 2025 (Build 26100)"
        WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Server 10.0%' AND SMS_R_System.Build LIKE '10.0.26100%'"
    }
)

# ─── Create Refresh Schedule (Weekly) ─────────────────────────────────
$Schedule = New-CMSchedule -RecurInterval Days -RecurCount 7

# ─── Create Collections ───────────────────────────────────────────────
$Created  = 0
$Skipped  = 0
$Errors   = 0

foreach ($Col in $Collections) {
    $CollectionName = $Col.Name

    # Check if collection already exists
    $Existing = Get-CMDeviceCollection -Name $CollectionName -ErrorAction SilentlyContinue
    if ($Existing) {
        Write-Host "SKIP:    $CollectionName (already exists)" -ForegroundColor DarkGray
        $Skipped++
        continue
    }

    try {
        # Create the collection
        $NewCollection = New-CMDeviceCollection `
            -Name             $CollectionName `
            -LimitingCollectionName "All Systems" `
            -RefreshType      Periodic `
            -RefreshSchedule  $Schedule `
            -Comment          $Col.Comment

        # Add the query membership rule
        Add-CMDeviceCollectionQueryMembershipRule `
            -CollectionName   $CollectionName `
            -RuleName         $CollectionName `
            -QueryExpression  $Col.WQL

        # Move to the console folder
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

# ─── Summary ──────────────────────────────────────────────────────────
Write-Host ""
Write-Host "═══════════════════════════════════════════" -ForegroundColor White
Write-Host "  Collections Created: $Created" -ForegroundColor Green
Write-Host "  Collections Skipped: $Skipped" -ForegroundColor DarkGray
Write-Host "  Errors:              $Errors" -ForegroundColor $(if ($Errors -gt 0) { "Red" } else { "DarkGray" })
Write-Host "═══════════════════════════════════════════" -ForegroundColor White
Write-Host ""

# Return to original location
Set-Location $OriginalLocation
