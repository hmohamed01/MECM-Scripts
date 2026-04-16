<#
.SYNOPSIS
    Performs a comprehensive health check of the Configuration Manager client.

.DESCRIPTION
    Validates ConfigMgr client health across services, WMI, policy, inventory,
    cache, certificates, and network connectivity. Writes a CMTrace-compatible
    log and returns a summary object.

.PARAMETER LogPath
    Path to the log file. Defaults to C:\Windows\CCM\Logs\ClientHealthCheck.log.

.PARAMETER ManagementPoint
    FQDN of a management point to test connectivity against. If omitted, the
    currently assigned MP is used.

.EXAMPLE
    .\Invoke-CMClientHealthCheck.ps1

.EXAMPLE
    .\Invoke-CMClientHealthCheck.ps1 -ManagementPoint "cm01.contoso.com" -LogPath "C:\Temp\CHC.log"

.NOTES
    Run as Administrator. Read-only — does not modify client state.
#>

[CmdletBinding()]
param(
    [string]$LogPath = "$env:WinDir\CCM\Logs\ClientHealthCheck.log",
    [string]$ManagementPoint
)

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warning','Error')][string]$Severity = 'Info',
        [string]$Component = 'ClientHealthCheck',
        [string]$LogFile = $script:LogPath
    )
    $typeMap = @{ Info = 1; Warning = 2; Error = 3 }
    $type = $typeMap[$Severity]
    $time = Get-Date -Format 'HH:mm:ss.fff'
    $offset = (Get-Date -Format 'zzz') -replace ':',''
    $date = Get-Date -Format 'MM-dd-yyyy'
    $thread = [System.Threading.Thread]::CurrentThread.ManagedThreadId
    $line = "<![LOG[$Message]LOG]!><time=`"$time$offset`" date=`"$date`" component=`"$Component`" context=`"`" type=`"$type`" thread=`"$thread`" file=`"`">"
    $dir = Split-Path $LogFile -Parent
    if (-not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
    Add-Content -Path $LogFile -Value $line -Encoding UTF8
    switch ($Severity) {
        'Info'    { Write-Host $Message }
        'Warning' { Write-Host $Message -ForegroundColor Yellow }
        'Error'   { Write-Host $Message -ForegroundColor Red }
    }
}

$script:LogPath = $LogPath
$results = [ordered]@{}

Write-CMTraceLog "===== Configuration Manager Client Health Check started on $env:COMPUTERNAME ====="

# 1. Required services
$requiredServices = @('CcmExec','Winmgmt','BITS','wuauserv','lanmanserver','lanmanworkstation')
foreach ($svc in $requiredServices) {
    try {
        $service = Get-Service -Name $svc -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            Write-CMTraceLog "Service [$svc] status: Running" -Severity Info
            $results["Service_$svc"] = 'Running'
        } else {
            Write-CMTraceLog "Service [$svc] status: $($service.Status) — expected Running" -Severity Warning
            $results["Service_$svc"] = $service.Status
        }
    } catch {
        Write-CMTraceLog "Service [$svc] not found: $_" -Severity Error
        $results["Service_$svc"] = 'NotFound'
    }
}

# 2. WMI repository accessible
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    Write-CMTraceLog "WMI query Win32_OperatingSystem succeeded: $($os.Caption)" -Severity Info
    $results['WMI_Query'] = 'OK'
} catch {
    Write-CMTraceLog "WMI query failed: $_" -Severity Error
    $results['WMI_Query'] = 'Failed'
}

# 3. Verify WMI repository integrity
try {
    $verifyResult = & winmgmt /verifyrepository 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-CMTraceLog "WMI repository verification: $verifyResult" -Severity Info
        $results['WMI_Repository'] = 'Consistent'
    } else {
        Write-CMTraceLog "WMI repository inconsistent: $verifyResult" -Severity Error
        $results['WMI_Repository'] = 'Inconsistent'
    }
} catch {
    Write-CMTraceLog "Failed to verify WMI repository: $_" -Severity Error
    $results['WMI_Repository'] = 'VerifyFailed'
}

# 4. ConfigMgr client namespace
try {
    $client = Get-CimInstance -Namespace 'root\ccm' -ClassName SMS_Client -ErrorAction Stop
    Write-CMTraceLog "ConfigMgr client version: $($client.ClientVersion)" -Severity Info
    $results['ClientVersion'] = $client.ClientVersion
} catch {
    Write-CMTraceLog "Cannot access root\ccm namespace: $_" -Severity Error
    $results['ClientVersion'] = 'Inaccessible'
}

# 5. Client GUID registration
try {
    $clientInfo = Get-CimInstance -Namespace 'root\ccm' -ClassName CCM_Client -ErrorAction Stop
    if ($clientInfo.ClientId) {
        Write-CMTraceLog "Client GUID: $($clientInfo.ClientId)" -Severity Info
        $results['ClientGuid'] = $clientInfo.ClientId
    } else {
        Write-CMTraceLog "Client GUID is empty — client may not be registered" -Severity Warning
        $results['ClientGuid'] = 'Missing'
    }
} catch {
    Write-CMTraceLog "Cannot retrieve client GUID: $_" -Severity Warning
    $results['ClientGuid'] = 'Unavailable'
}

# 6. Assigned management point
try {
    $mp = Get-CimInstance -Namespace 'root\ccm' -ClassName SMS_Authority -ErrorAction Stop
    Write-CMTraceLog "Site code: $($mp.Name); Current MP: $($mp.CurrentManagementPoint)" -Severity Info
    $results['SiteCode'] = $mp.Name
    $results['AssignedMP'] = $mp.CurrentManagementPoint
    if (-not $ManagementPoint) { $ManagementPoint = $mp.CurrentManagementPoint }
} catch {
    Write-CMTraceLog "Cannot retrieve assigned site/MP: $_" -Severity Warning
    $results['SiteCode'] = 'Unknown'
}

# 7. Management point connectivity
if ($ManagementPoint) {
    try {
        $mpResponse = Invoke-WebRequest -Uri "http://$ManagementPoint/SMS_MP/.sms_aut?MPLIST" -UseBasicParsing -TimeoutSec 15 -ErrorAction Stop
        if ($mpResponse.StatusCode -eq 200) {
            Write-CMTraceLog "MP [$ManagementPoint] responded HTTP 200" -Severity Info
            $results['MPConnectivity'] = 'OK'
        } else {
            Write-CMTraceLog "MP [$ManagementPoint] responded HTTP $($mpResponse.StatusCode)" -Severity Warning
            $results['MPConnectivity'] = "HTTP_$($mpResponse.StatusCode)"
        }
    } catch {
        Write-CMTraceLog "MP [$ManagementPoint] unreachable: $_" -Severity Error
        $results['MPConnectivity'] = 'Unreachable'
    }
}

# 8. Last hardware inventory
try {
    $hinv = Get-CimInstance -Namespace 'root\ccm\invagt' -ClassName InventoryActionStatus -Filter "InventoryActionID='{00000000-0000-0000-0000-000000000001}'" -ErrorAction Stop
    if ($hinv -and $hinv.LastCycleStartedDate) {
        $lastHinv = [Management.ManagementDateTimeConverter]::ToDateTime($hinv.LastCycleStartedDate)
        $age = (Get-Date) - $lastHinv
        $msg = "Last hardware inventory: $lastHinv ($([int]$age.TotalDays) days ago)"
        if ($age.TotalDays -gt 7) {
            Write-CMTraceLog $msg -Severity Warning
        } else {
            Write-CMTraceLog $msg -Severity Info
        }
        $results['LastHWInventory'] = $lastHinv
    } else {
        Write-CMTraceLog "Hardware inventory has never run" -Severity Warning
        $results['LastHWInventory'] = 'Never'
    }
} catch {
    Write-CMTraceLog "Cannot retrieve hardware inventory status: $_" -Severity Warning
    $results['LastHWInventory'] = 'Unavailable'
}

# 9. Client cache
try {
    $cache = Get-CimInstance -Namespace 'root\ccm\softmgmtagent' -ClassName CacheConfig -ErrorAction Stop
    Write-CMTraceLog "Client cache: Location=$($cache.Location), Size=$($cache.Size) MB" -Severity Info
    $results['CacheSize_MB'] = $cache.Size
} catch {
    Write-CMTraceLog "Cannot retrieve cache config: $_" -Severity Warning
    $results['CacheSize_MB'] = 'Unknown'
}

# 10. Client certificate
try {
    $certInfo = Get-CimInstance -Namespace 'root\ccm' -ClassName CCM_ClientIdentificationInformation -ErrorAction Stop
    if ($certInfo.ClientId) {
        Write-CMTraceLog "Client identification certificate present" -Severity Info
        $results['ClientCert'] = 'Present'
    }
} catch {
    Write-CMTraceLog "Cannot retrieve client certificate info" -Severity Warning
    $results['ClientCert'] = 'Unknown'
}

# 11. Disk space on system drive
$sysDrive = Get-CimInstance -ClassName Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'"
$freeGB = [math]::Round($sysDrive.FreeSpace / 1GB, 2)
$totalGB = [math]::Round($sysDrive.Size / 1GB, 2)
$freePct = [math]::Round(($sysDrive.FreeSpace / $sysDrive.Size) * 100, 1)
$msg = "System drive: $freeGB GB free of $totalGB GB ($freePct%)"
if ($freePct -lt 10) {
    Write-CMTraceLog $msg -Severity Error
} elseif ($freePct -lt 20) {
    Write-CMTraceLog $msg -Severity Warning
} else {
    Write-CMTraceLog $msg -Severity Info
}
$results['SystemDriveFree_GB'] = $freeGB

# 12. Pending reboot check
$rebootKeys = @(
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
    'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
)
$pendingReboot = $false
foreach ($key in $rebootKeys) {
    if (Test-Path $key) { $pendingReboot = $true; break }
}
if ($pendingReboot) {
    Write-CMTraceLog "System has a pending reboot" -Severity Warning
    $results['PendingReboot'] = $true
} else {
    Write-CMTraceLog "No pending reboot detected" -Severity Info
    $results['PendingReboot'] = $false
}

Write-CMTraceLog "===== Configuration Manager Client Health Check complete. Log: $LogPath ====="

[PSCustomObject]$results
