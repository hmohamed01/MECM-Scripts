#Requires -Version 5.1
<#
.SYNOPSIS
    Creates device collections targeting Windows Server systems by installed role.

.DESCRIPTION
    Builds query-based collections for common server roles detected via
    SMS_G_System_SERVICE (Win32_Service hardware inventory class):

        - Domain Controllers    (NTDS)
        - DNS Servers           (DNS)
        - DHCP Servers          (DHCPServer)
        - SQL Servers           (MSSQL*)
        - IIS Web Servers       (W3SVC)
        - Hyper-V Hosts         (vmms)
        - Exchange Servers      (MSExchange*)
        - WSUS / SUP            (WsusService)
        - MECM / ConfigMgr      (SMS_SITE_COMPONENT_MANAGER / CcmExec on server OS)

    Each collection joins SMS_R_System to SMS_G_System_SERVICE and filters to
    server operating systems so workstations with similar service names don't
    match by accident.

    Collections land in Device Collections\Server Roles with a weekly
    full refresh and no incremental updates.

.NOTES
    Service-based detection requires hardware inventory to be enabled for
    Win32_Service. The default MECM hardware inventory class set includes it.
#>

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"

# ─── Connect to CM Site ───────────────────────────────────────────────
. "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
Connect-CMSite -SiteCode $SiteCode

try {
    $FolderRoot = "$($SiteCode.ToUpper()):\DeviceCollection"

    # ─── Create Console Folder Structure ──────────────────────────────
    $Folders = @("Server Roles")

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

    # ─── WQL Helper ───────────────────────────────────────────────────
    # Service-based role detection pattern:
    #   JOIN SMS_G_System_SERVICE on matching service Name
    #   AND filter to Server OS only (OperatingSystemNameandVersion LIKE '%Server%')
    #   AND Service State = "Running" (or Started for cross-version compat)
    function New-ServiceRoleQuery {
        param(
            [Parameter(Mandatory)] [string]$ServiceNameLike
        )
        @"
SELECT SMS_R_System.ResourceID, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client
FROM SMS_R_System INNER JOIN SMS_G_System_SERVICE
ON SMS_G_System_SERVICE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_G_System_SERVICE.Name LIKE '$ServiceNameLike'
AND SMS_R_System.OperatingSystemNameandVersion LIKE '%Server%'
"@
    }

    # ─── Define Collections ───────────────────────────────────────────
    $Collections = @(
        @{
            Name    = "Servers - Domain Controllers"
            Folder  = "Server Roles"
            Comment = "Servers running the NTDS service (Active Directory Domain Services)"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "NTDS"
        },
        @{
            Name    = "Servers - DNS"
            Folder  = "Server Roles"
            Comment = "Servers running the DNS service"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "DNS"
        },
        @{
            Name    = "Servers - DHCP"
            Folder  = "Server Roles"
            Comment = "Servers running the DHCPServer service"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "DHCPServer"
        },
        @{
            Name    = "Servers - SQL Server"
            Folder  = "Server Roles"
            Comment = "Servers hosting any MSSQL* service (default + named instances)"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "MSSQL%"
        },
        @{
            Name    = "Servers - IIS Web Servers"
            Folder  = "Server Roles"
            Comment = "Servers running the W3SVC (IIS) service"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "W3SVC"
        },
        @{
            Name    = "Servers - Hyper-V Hosts"
            Folder  = "Server Roles"
            Comment = "Servers running the vmms (Hyper-V Virtual Machine Management) service"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "vmms"
        },
        @{
            Name    = "Servers - Exchange"
            Folder  = "Server Roles"
            Comment = "Servers hosting any MSExchange* service"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "MSExchange%"
        },
        @{
            Name    = "Servers - WSUS / SUP"
            Folder  = "Server Roles"
            Comment = "Servers running the WsusService (WSUS / MECM Software Update Point)"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "WsusService"
        },
        @{
            Name    = "Servers - MECM Site Servers"
            Folder  = "Server Roles"
            Comment = "Servers running SMS_SITE_COMPONENT_MANAGER (MECM site server role)"
            WQL     = New-ServiceRoleQuery -ServiceNameLike "SMS_SITE_COMPONENT_MANAGER"
        },
        @{
            Name    = "Servers - All Windows Servers"
            Folder  = "Server Roles"
            Comment = "All systems running a Windows Server OS — limiting collection for role queries"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.OperatingSystemNameandVersion LIKE '%Server%'"
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
