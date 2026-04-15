#Requires -Version 5.1
<#
.SYNOPSIS
    Triggers a MECM client action on every device in a collection.

.DESCRIPTION
    Enumerates members of the specified collection and invokes the chosen
    client action via WMI (SMS_Client.TriggerSchedule) on each machine.

    Standard schedule GUIDs are used — works against any client version.
    Actions run sequentially. Progress and per-machine success/failure are
    reported, with a counter summary at the end.

.PARAMETER CollectionName
    Target device collection name.

.PARAMETER Action
    Client action to trigger. Valid values:
        MachinePolicy       - Retrieve + evaluate machine policy
        HardwareInventory   - Hardware inventory cycle
        SoftwareInventory   - Software inventory cycle
        AppDeployEval       - Application deployment evaluation
        UpdateScan          - Software updates scan cycle
        UpdateDeployEval    - Software update deployment evaluation
        DCMEval             - Configuration baseline (DCM) evaluation
        FileCollection      - File collection cycle
        DiscoveryData       - Discovery data collection

.PARAMETER ThrottleLimit
    Not used in sequential mode. Reserved for future parallelization.

.EXAMPLE
    .\Invoke-ClientActionOnCollection.ps1 -CollectionName "Patch Ring - Pilot" -Action MachinePolicy

.EXAMPLE
    .\Invoke-ClientActionOnCollection.ps1 -CollectionName "All Workstations" -Action HardwareInventory

.NOTES
    Requires the executing account to have remote WMI access (root\ccm)
    on each target machine. Firewall rules for WMI-in and ports 135 + dynamic
    RPC must allow the connection.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CollectionName,

    [Parameter(Mandatory = $true)]
    [ValidateSet(
        'MachinePolicy',
        'HardwareInventory',
        'SoftwareInventory',
        'AppDeployEval',
        'UpdateScan',
        'UpdateDeployEval',
        'DCMEval',
        'FileCollection',
        'DiscoveryData'
    )]
    [string]$Action
)

# ─── Schedule GUID Map ────────────────────────────────────────────────
# https://learn.microsoft.com/en-us/mem/configmgr/develop/reference/core/clients/manage/sms_client-server-wmi-class-methods
$ScheduleMap = @{
    'MachinePolicy'     = '{00000000-0000-0000-0000-000000000022}'  # Machine Policy Evaluation
    'HardwareInventory' = '{00000000-0000-0000-0000-000000000001}'
    'SoftwareInventory' = '{00000000-0000-0000-0000-000000000002}'
    'AppDeployEval'     = '{00000000-0000-0000-0000-000000000121}'
    'UpdateScan'        = '{00000000-0000-0000-0000-000000000113}'  # Scan cycle
    'UpdateDeployEval'  = '{00000000-0000-0000-0000-000000000108}'  # Assignment evaluation
    'DCMEval'           = '{00000000-0000-0000-0000-000000000111}'
    'FileCollection'    = '{00000000-0000-0000-0000-000000000010}'
    'DiscoveryData'     = '{00000000-0000-0000-0000-000000000003}'
}
$ScheduleId = $ScheduleMap[$Action]

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"

# ─── Connect to CM Site ───────────────────────────────────────────────
. "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
Connect-CMSite -SiteCode $SiteCode

try {
    # Validate collection exists
    $Collection = Get-CMDeviceCollection -Name $CollectionName -ErrorAction SilentlyContinue
    if (-not $Collection) {
        throw "Collection '$CollectionName' not found."
    }

    Write-Host "Collection: $CollectionName ($($Collection.CollectionID))" -ForegroundColor Cyan
    Write-Host "Action:     $Action ($ScheduleId)" -ForegroundColor Cyan

    # Enumerate members
    Write-Host "Enumerating members..." -ForegroundColor Cyan
    $Members = Get-CMCollectionMember -CollectionName $CollectionName | Select-Object -ExpandProperty Name
    $TotalCount = $Members.Count
    Write-Host "  $TotalCount members." -ForegroundColor DarkGray

    if ($TotalCount -eq 0) {
        Write-Host "Collection is empty. Nothing to do." -ForegroundColor Yellow
        return
    }

    # ─── Invoke Action Per Member ─────────────────────────────────────
    $Success = 0
    $Failed  = 0
    $i = 0

    # Step out of the CM PSDrive so WMI cmdlets work cleanly
    Push-Location $Script:CMOriginalLocation

    try {
        foreach ($Computer in $Members) {
            $i++
            Write-Progress -Activity "Triggering $Action" `
                           -Status "$i of $TotalCount — $Computer" `
                           -PercentComplete (($i / $TotalCount) * 100)

            try {
                $null = Invoke-WmiMethod -ComputerName $Computer `
                                         -Namespace 'root\ccm' `
                                         -Class 'SMS_Client' `
                                         -Name 'TriggerSchedule' `
                                         -ArgumentList $ScheduleId `
                                         -ErrorAction Stop
                Write-Host "OK:   $Computer" -ForegroundColor Green
                $Success++
            }
            catch {
                Write-Host "FAIL: $Computer — $($_.Exception.Message)" -ForegroundColor Red
                $Failed++
            }
        }
    }
    finally {
        Pop-Location
        Write-Progress -Activity "Triggering $Action" -Completed
    }

    # ─── Summary ──────────────────────────────────────────────────────
    Write-Host ""
    Write-Host "═══════════════════════════════════════════" -ForegroundColor White
    Write-Host "  Action:              $Action"
    Write-Host "  Targeted:            $TotalCount" -ForegroundColor DarkGray
    Write-Host "  Success:             $Success" -ForegroundColor Green
    Write-Host "  Failed:              $Failed" -ForegroundColor $(if ($Failed -gt 0) { "Red" } else { "DarkGray" })
    Write-Host "═══════════════════════════════════════════" -ForegroundColor White
    Write-Host ""
}
finally {
    Disconnect-CMSite
}
