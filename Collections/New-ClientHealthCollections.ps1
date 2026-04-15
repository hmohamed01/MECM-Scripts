#Requires -Version 5.1
<#
.SYNOPSIS
    Creates device collections for client health triage in MECM.

.DESCRIPTION
    Builds a standard set of device collections that surface unhealthy
    or stale clients for operational review:

        - No MECM Client Installed
        - Obsolete Records
        - Inactive Clients (per client activity status)
        - No Heartbeat 30/60/90+ Days
        - No Hardware Inventory 30 Days
        - No Policy Request 14 Days
        - All Active Healthy Clients

    Collections are placed under Device Collections\Client Health and
    refreshed weekly. Existing collections are skipped (idempotent).

.NOTES
    Run on a machine with the MECM console installed.
    Requires permissions to create collections.

    Date arithmetic uses DateAdd/GetDate WQL functions supported by the
    ConfigMgr collection query engine.
#>

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"

# ─── Connect to CM Site ───────────────────────────────────────────────
. "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
Connect-CMSite -SiteCode $SiteCode

try {
    $FolderRoot = "$($SiteCode.ToUpper()):\DeviceCollection"

    # ─── Create Console Folder Structure ──────────────────────────────
    $Folders = @("Client Health")

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
    $Collections = @(
        @{
            Name    = "Clients - No MECM Client Installed"
            Folder  = "Client Health"
            Comment = "Discovered systems with Client=0 or NULL — no MECM agent"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.Client = 0 OR SMS_R_System.Client IS NULL"
        },
        @{
            Name    = "Clients - Obsolete Records"
            Folder  = "Client Health"
            Comment = "Obsolete=1 — duplicate/stale records flagged by site maintenance"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.Obsolete = 1"
        },
        @{
            Name    = "Clients - Inactive (per ClientSummary)"
            Folder  = "Client Health"
            Comment = "Clients marked inactive by SMS_CH_ClientSummary.ClientActiveStatus"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client
FROM SMS_R_System INNER JOIN SMS_CH_ClientSummary
ON SMS_CH_ClientSummary.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0 AND SMS_CH_ClientSummary.ClientActiveStatus = 0
"@
        },
        @{
            Name    = "Clients - No Heartbeat 30+ Days"
            Folder  = "Client Health"
            Comment = "AgentTime (heartbeat discovery) older than 30 days"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.AgentTime < DateAdd(dd,-30,GetDate()) AND SMS_R_System.Client = 1"
        },
        @{
            Name    = "Clients - No Heartbeat 60+ Days"
            Folder  = "Client Health"
            Comment = "AgentTime (heartbeat discovery) older than 60 days"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.AgentTime < DateAdd(dd,-60,GetDate()) AND SMS_R_System.Client = 1"
        },
        @{
            Name    = "Clients - No Heartbeat 90+ Days"
            Folder  = "Client Health"
            Comment = "AgentTime (heartbeat discovery) older than 90 days — strong cleanup candidate"
            WQL     = "SELECT * FROM SMS_R_System WHERE SMS_R_System.AgentTime < DateAdd(dd,-90,GetDate()) AND SMS_R_System.Client = 1"
        },
        @{
            Name    = "Clients - No Hardware Inventory 30+ Days"
            Folder  = "Client Health"
            Comment = "LastHardwareScan older than 30 days"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client
FROM SMS_R_System INNER JOIN SMS_CH_ClientSummary
ON SMS_CH_ClientSummary.ResourceID = SMS_R_System.ResourceID
WHERE SMS_CH_ClientSummary.LastHardwareScan < DateAdd(dd,-30,GetDate()) AND SMS_R_System.Client = 1
"@
        },
        @{
            Name    = "Clients - No Policy Request 14+ Days"
            Folder  = "Client Health"
            Comment = "LastPolicyRequest older than 14 days — likely policy or MP communication issue"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client
FROM SMS_R_System INNER JOIN SMS_CH_ClientSummary
ON SMS_CH_ClientSummary.ResourceID = SMS_R_System.ResourceID
WHERE SMS_CH_ClientSummary.LastPolicyRequest < DateAdd(dd,-14,GetDate()) AND SMS_R_System.Client = 1
"@
        },
        @{
            Name    = "Clients - All Active Healthy"
            Folder  = "Client Health"
            Comment = "Client=1, not obsolete, ClientActiveStatus=1 — healthy baseline"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.ResourceType, SMS_R_System.Name, SMS_R_System.SMSUniqueIdentifier, SMS_R_System.ResourceDomainORWorkgroup, SMS_R_System.Client
FROM SMS_R_System INNER JOIN SMS_CH_ClientSummary
ON SMS_CH_ClientSummary.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0 AND SMS_CH_ClientSummary.ClientActiveStatus = 1
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
