#Requires -Version 5.1
<#
.SYNOPSIS
    Creates device collections for client health triage in MECM.

.DESCRIPTION
    Builds a standard set of device collections that surface unhealthy
    or stale clients for operational review:

    Client Health:
        - No MECM Client Installed
        - Obsolete Records
        - Inactive Clients (per client activity status)
        - No Heartbeat 30/60/90+ Days
        - No Hardware Inventory 30 Days
        - No Policy Request 14 Days
        - All Active Healthy Clients

    Software Update Health:
        - WUA Service Not Running / Disabled
        - BITS Service Not Running
        - CcmEval Health Evaluation Failed
        - Last Software Update Scan 14+ Days Old
        - Update Enforcement Failures
        - Pending Reboot (blocking update installs)
        - Update Compliant (all deployments satisfied)

    Collections are placed under Device Collections\Client Health (and
    the Software Update Health subfolder) and refreshed weekly. Existing
    collections are skipped (idempotent).

.NOTES
    Run on a machine with the MECM console installed.
    Requires permissions to create collections.

    Software Update Health collections require Win32_Service enabled in
    hardware inventory (default). Service-state collections reflect the
    state at last hardware inventory, not real-time.

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
    $Folders = @(
        "Client Health",
        "Client Health\Software Update Health"
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
        },

        # ── Software Update Health ───────────────────────────────────
        # Service-state queries use SMS_G_System_SERVICE (hardware inventory).
        # State reflects last HW scan — not real-time.

        @{
            Name    = "SU Health - WUA Service Not Running"
            Folder  = "Client Health\Software Update Health"
            Comment = "Windows Update Agent (wuauserv) not in Running state at last hardware inventory"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_SERVICE ON SMS_G_System_SERVICE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_G_System_SERVICE.Name = 'wuauserv'
AND SMS_G_System_SERVICE.State != 'Running'
"@
        },
        @{
            Name    = "SU Health - WUA Service Disabled"
            Folder  = "Client Health\Software Update Health"
            Comment = "Windows Update Agent start mode set to Disabled — updates cannot run"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_SERVICE ON SMS_G_System_SERVICE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_G_System_SERVICE.Name = 'wuauserv'
AND SMS_G_System_SERVICE.StartMode = 'Disabled'
"@
        },
        @{
            Name    = "SU Health - BITS Service Not Running"
            Folder  = "Client Health\Software Update Health"
            Comment = "Background Intelligent Transfer Service not running — blocks content download"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_G_System_SERVICE ON SMS_G_System_SERVICE.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_G_System_SERVICE.Name = 'BITS'
AND SMS_G_System_SERVICE.State != 'Running'
"@
        },
        @{
            Name    = "SU Health - CcmEval Health Check Failed"
            Folder  = "Client Health\Software Update Health"
            Comment = "CcmEval reported unhealthy and could not auto-remediate (Result 7 = not fixed)"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_CH_EvalResult ON SMS_CH_EvalResult.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_CH_EvalResult.Result = 7
"@
        },
        @{
            Name    = "SU Health - CcmEval Remediated"
            Folder  = "Client Health\Software Update Health"
            Comment = "CcmEval detected a failure and auto-remediated — monitor for recurring issues"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_CH_EvalResult ON SMS_CH_EvalResult.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_CH_EvalResult.Result = 6
"@
        },
        @{
            Name    = "SU Health - No Update Scan 14+ Days"
            Folder  = "Client Health\Software Update Health"
            Comment = "LastSWUpdateScanTime older than 14 days — scan may be failing or SUP unreachable"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_CH_ClientSummary ON SMS_CH_ClientSummary.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_CH_ClientSummary.LastSWUpdateScanTime < DateAdd(dd,-14,GetDate())
"@
        },
        @{
            Name    = "SU Health - No Update Scan 30+ Days"
            Folder  = "Client Health\Software Update Health"
            Comment = "LastSWUpdateScanTime older than 30 days — client is not scanning for updates"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_CH_ClientSummary ON SMS_CH_ClientSummary.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_CH_ClientSummary.LastSWUpdateScanTime < DateAdd(dd,-30,GetDate())
"@
        },
        @{
            Name    = "SU Health - Update Enforcement Failed"
            Folder  = "Client Health\Software Update Health"
            Comment = "At least one update deployment failed enforcement (state 13) — install error on client"
            WQL     = @"
SELECT DISTINCT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_UpdateComplianceStatus ON SMS_UpdateComplianceStatus.MachineID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_UpdateComplianceStatus.LastEnforcementMessageID = 13
"@
        },
        @{
            Name    = "SU Health - Pending Reboot (Update Related)"
            Folder  = "Client Health\Software Update Health"
            Comment = "Update installed successfully but pending reboot (enforcement state 8 or 9) — blocking further installs"
            WQL     = @"
SELECT DISTINCT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_UpdateComplianceStatus ON SMS_UpdateComplianceStatus.MachineID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_UpdateComplianceStatus.LastEnforcementMessageID IN (8, 9)
"@
        },
        @{
            Name    = "SU Health - All Deployments Compliant"
            Folder  = "Client Health\Software Update Health"
            Comment = "All targeted update deployments are satisfied — healthy update baseline"
            WQL     = @"
SELECT SMS_R_System.ResourceID, SMS_R_System.Name
FROM SMS_R_System
INNER JOIN SMS_CH_ClientSummary ON SMS_CH_ClientSummary.ResourceID = SMS_R_System.ResourceID
WHERE SMS_R_System.Client = 1 AND SMS_R_System.Obsolete = 0
AND SMS_CH_ClientSummary.ClientActiveStatus = 1
AND SMS_CH_ClientSummary.LastSWUpdateScanTime > DateAdd(dd,-7,GetDate())
AND SMS_R_System.ResourceID NOT IN (
    SELECT SMS_UpdateComplianceStatus.MachineID
    FROM SMS_UpdateComplianceStatus
    WHERE SMS_UpdateComplianceStatus.LastEnforcementMessageID IN (8, 9, 13)
)
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
