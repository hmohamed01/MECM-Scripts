<#
.SYNOPSIS
    Performs a health check of software update and patching components on a
    Configuration Manager client.

.DESCRIPTION
    Validates the full software update stack on a ConfigMgr-managed device:

    Windows Update Agent (WUA):
        - wuauserv service status
        - WUA version
        - SoftwareDistribution datastore size and age
        - Pending / in-progress / failed updates via COM API
        - WSUS server registration in the registry

    ConfigMgr Software Update Client Agent:
        - CcmExec service status
        - Software Update Client Agent enabled status
        - Last scan time and scan source (SUP assignment)
        - Update deployment compliance state
        - Cached update content vs cache capacity

    General:
        - Pending reboot (CBS, Windows Update, ConfigMgr)
        - BITS service status
        - Disk space on system drive

    Writes a CMTrace-compatible log and returns a summary object.

.PARAMETER LogPath
    Path to the log file. Defaults to C:\Windows\CCM\Logs\UpdateHealthCheck.log.

.EXAMPLE
    .\Invoke-CMUpdateHealthCheck.ps1

.EXAMPLE
    .\Invoke-CMUpdateHealthCheck.ps1 -LogPath "C:\Temp\UpdateHealth.log"

.NOTES
    Run as Administrator. Read-only — does not modify client or WUA state.
#>

[CmdletBinding()]
param(
    [string]$LogPath = "$env:WinDir\CCM\Logs\UpdateHealthCheck.log"
)

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warning','Error')][string]$Severity = 'Info',
        [string]$Component = 'UpdateHealthCheck',
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

Write-CMTraceLog "===== Software Update Health Check started on $env:COMPUTERNAME ====="

# ─── 1. Required services ────────────────────────────────────────────
$updateServices = @('wuauserv','CcmExec','BITS')
foreach ($svc in $updateServices) {
    try {
        $service = Get-Service -Name $svc -ErrorAction Stop
        if ($service.Status -eq 'Running') {
            Write-CMTraceLog "Service [$svc] status: Running" -Severity Info
            $results["Service_$svc"] = 'Running'
        } else {
            Write-CMTraceLog "Service [$svc] status: $($service.Status) — expected Running" -Severity Error
            $results["Service_$svc"] = $service.Status
        }
    } catch {
        Write-CMTraceLog "Service [$svc] not found: $_" -Severity Error
        $results["Service_$svc"] = 'NotFound'
    }
}

# ─── 2. Windows Update Agent version ─────────────────────────────────
try {
    $wuaDll = "$env:WinDir\System32\wuaueng.dll"
    if (Test-Path $wuaDll) {
        $wuaVersion = (Get-Item $wuaDll).VersionInfo.ProductVersion
        Write-CMTraceLog "Windows Update Agent version: $wuaVersion" -Severity Info
        $results['WUAVersion'] = $wuaVersion
    } else {
        Write-CMTraceLog "WUA engine DLL not found at $wuaDll" -Severity Error
        $results['WUAVersion'] = 'NotFound'
    }
} catch {
    Write-CMTraceLog "Cannot determine WUA version: $_" -Severity Warning
    $results['WUAVersion'] = 'Unknown'
}

# ─── 3. WSUS server registration ─────────────────────────────────────
$wuRegPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
try {
    $wuReg = Get-ItemProperty -Path $wuRegPath -ErrorAction Stop
    $wsusServer = $wuReg.WUServer
    $wsusStatus = $wuReg.WUStatusServer
    if ($wsusServer) {
        Write-CMTraceLog "WSUS server: $wsusServer" -Severity Info
        $results['WSUSServer'] = $wsusServer
    } else {
        Write-CMTraceLog "WSUS server not configured in registry — client may be scanning against Microsoft Update" -Severity Warning
        $results['WSUSServer'] = 'NotConfigured'
    }
    if ($wsusStatus) {
        Write-CMTraceLog "WSUS status server: $wsusStatus" -Severity Info
        $results['WSUSStatusServer'] = $wsusStatus
    }
} catch {
    Write-CMTraceLog "Windows Update policy registry key not found — WSUS not configured via policy" -Severity Warning
    $results['WSUSServer'] = 'NoPolicyKey'
}

# ─── 4. SoftwareDistribution datastore ───────────────────────────────
$sdPath = "$env:WinDir\SoftwareDistribution"
if (Test-Path $sdPath) {
    try {
        $sdSize = (Get-ChildItem -Path $sdPath -Recurse -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
        $sdSizeMB = [math]::Round($sdSize / 1MB, 2)
        $dbPath = Join-Path $sdPath 'DataStore\DataStore.edb'
        if (Test-Path $dbPath) {
            $dbAge = (Get-Date) - (Get-Item $dbPath).LastWriteTime
            $msg = "SoftwareDistribution: $sdSizeMB MB total, datastore last modified $([int]$dbAge.TotalDays) days ago"
            if ($sdSizeMB -gt 5120) {
                Write-CMTraceLog $msg -Severity Warning
            } else {
                Write-CMTraceLog $msg -Severity Info
            }
            $results['SoftwareDistribution_MB'] = $sdSizeMB
            $results['DataStoreAge_Days'] = [int]$dbAge.TotalDays
        } else {
            Write-CMTraceLog "SoftwareDistribution: $sdSizeMB MB (datastore file not found)" -Severity Warning
            $results['SoftwareDistribution_MB'] = $sdSizeMB
        }
    } catch {
        Write-CMTraceLog "Cannot measure SoftwareDistribution folder: $_" -Severity Warning
    }
} else {
    Write-CMTraceLog "SoftwareDistribution folder not found at $sdPath" -Severity Error
    $results['SoftwareDistribution_MB'] = 'NotFound'
}

# ─── 5. Pending updates via WUA COM API ──────────────────────────────
try {
    $updateSession = New-Object -ComObject Microsoft.Update.Session
    $searcher = $updateSession.CreateUpdateSearcher()

    # Pending (not installed)
    $pendingResult = $searcher.Search("IsInstalled=0 AND IsHidden=0")
    $pendingCount = $pendingResult.Updates.Count
    Write-CMTraceLog "Pending updates (not installed, not hidden): $pendingCount" -Severity $(if ($pendingCount -gt 0) { 'Warning' } else { 'Info' })
    $results['PendingUpdates'] = $pendingCount

    # Count by severity if available
    $critical = 0
    $important = 0
    foreach ($update in $pendingResult.Updates) {
        $rating = $update.MsrcSeverity
        if ($rating -eq 'Critical') { $critical++ }
        elseif ($rating -eq 'Important') { $important++ }
    }
    if ($critical -gt 0) {
        Write-CMTraceLog "  Critical: $critical" -Severity Error
        $results['PendingCritical'] = $critical
    }
    if ($important -gt 0) {
        Write-CMTraceLog "  Important: $important" -Severity Warning
        $results['PendingImportant'] = $important
    }

    # Failed installs in update history
    $historyCount = $searcher.GetTotalHistoryCount()
    if ($historyCount -gt 0) {
        $history = $searcher.QueryHistory(0, [math]::Min($historyCount, 50))
        $failed = @($history | Where-Object { $_.ResultCode -eq 4 -or $_.ResultCode -eq 5 })
        if ($failed.Count -gt 0) {
            Write-CMTraceLog "Failed updates in recent history: $($failed.Count)" -Severity Error
            foreach ($f in $failed | Select-Object -First 5) {
                Write-CMTraceLog "  FAILED: $($f.Title) (HResult: 0x$($f.HResult.ToString('X8')))" -Severity Error
            }
            $results['FailedUpdates'] = $failed.Count
        } else {
            Write-CMTraceLog "No failed updates in recent history" -Severity Info
            $results['FailedUpdates'] = 0
        }
    }
} catch {
    Write-CMTraceLog "Cannot query WUA COM API: $_" -Severity Error
    $results['PendingUpdates'] = 'ComError'
}

# ─── 6. ConfigMgr Software Update Client Agent ───────────────────────
try {
    $suAgent = Get-CimInstance -Namespace 'root\ccm\policy\machine\actualconfig' `
        -ClassName CCM_SoftwareUpdatesClientConfig -ErrorAction Stop
    if ($suAgent) {
        $enabled = $suAgent.Enabled
        $scanInterval = $suAgent.ScanSchedule
        if ($enabled) {
            Write-CMTraceLog "Software Update Client Agent: Enabled" -Severity Info
            $results['SUClientAgent'] = 'Enabled'
        } else {
            Write-CMTraceLog "Software Update Client Agent: DISABLED — updates will not be managed" -Severity Error
            $results['SUClientAgent'] = 'Disabled'
        }
    }
} catch {
    Write-CMTraceLog "Cannot query Software Update Client Agent config: $_" -Severity Warning
    $results['SUClientAgent'] = 'Unavailable'
}

# ─── 7. Last software update scan ────────────────────────────────────
try {
    $scanStatus = Get-CimInstance -Namespace 'root\ccm\scanagent' `
        -ClassName CCM_ScanStatus -ErrorAction Stop |
        Sort-Object LastScanTime -Descending | Select-Object -First 1
    if ($scanStatus -and $scanStatus.LastScanTime) {
        $lastScan = $scanStatus.LastScanTime
        $scanAge = (Get-Date) - $lastScan
        $msg = "Last software update scan: $lastScan ($([int]$scanAge.TotalHours) hours ago)"
        if ($scanAge.TotalDays -gt 7) {
            Write-CMTraceLog $msg -Severity Error
        } elseif ($scanAge.TotalDays -gt 3) {
            Write-CMTraceLog $msg -Severity Warning
        } else {
            Write-CMTraceLog $msg -Severity Info
        }
        $results['LastScanTime'] = $lastScan
        $results['LastScanAge_Hours'] = [int]$scanAge.TotalHours
    } else {
        Write-CMTraceLog "Software update scan has never completed" -Severity Error
        $results['LastScanTime'] = 'Never'
    }
} catch {
    Write-CMTraceLog "Cannot query scan status from root\ccm\scanagent: $_" -Severity Warning
    $results['LastScanTime'] = 'Unavailable'
}

# ─── 8. SUP assignment (scan source) ─────────────────────────────────
try {
    $supInfo = Get-CimInstance -Namespace 'root\ccm\locationservices' `
        -ClassName SMS_ActiveSoftwareUpdatePoint -ErrorAction Stop
    if ($supInfo) {
        $wsusHost = $supInfo.WSUSServerName
        $wsusPort = $supInfo.WSUSServerPort
        Write-CMTraceLog "Assigned SUP: $wsusHost`:$wsusPort" -Severity Info
        $results['AssignedSUP'] = "$wsusHost`:$wsusPort"
    } else {
        Write-CMTraceLog "No active Software Update Point assigned" -Severity Error
        $results['AssignedSUP'] = 'None'
    }
} catch {
    Write-CMTraceLog "Cannot query SUP assignment: $_" -Severity Warning
    $results['AssignedSUP'] = 'Unavailable'
}

# ─── 9. Update deployment compliance ─────────────────────────────────
try {
    $assignments = @(Get-CimInstance -Namespace 'root\ccm\softwareupdates\deploymentagent' `
        -ClassName CCM_AssignmentCompliance -ErrorAction Stop)
    if ($assignments.Count -gt 0) {
        $compliant     = @($assignments | Where-Object { $_.IsCompliant -eq $true }).Count
        $nonCompliant  = @($assignments | Where-Object { $_.IsCompliant -eq $false }).Count
        Write-CMTraceLog "Update deployments: $($assignments.Count) total, $compliant compliant, $nonCompliant non-compliant"
        if ($nonCompliant -gt 0) {
            Write-CMTraceLog "  $nonCompliant deployment(s) are non-compliant — updates pending install" -Severity Warning
        }
        $results['Deployments_Total'] = $assignments.Count
        $results['Deployments_Compliant'] = $compliant
        $results['Deployments_NonCompliant'] = $nonCompliant
    } else {
        Write-CMTraceLog "No software update deployments targeted to this client" -Severity Info
        $results['Deployments_Total'] = 0
    }
} catch {
    Write-CMTraceLog "Cannot query deployment compliance: $_" -Severity Warning
    $results['Deployments_Total'] = 'Unavailable'
}

# ─── 10. ConfigMgr client cache usage ────────────────────────────────
try {
    $cache = Get-CimInstance -Namespace 'root\ccm\softmgmtagent' -ClassName CacheConfig -ErrorAction Stop
    $cacheSizeMB = $cache.Size
    $cachePath = $cache.Location
    if (Test-Path $cachePath) {
        $cacheUsed = (Get-ChildItem -Path $cachePath -Recurse -File -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum).Sum
        $cacheUsedMB = [math]::Round($cacheUsed / 1MB, 0)
        $usedPct = [math]::Round(($cacheUsedMB / $cacheSizeMB) * 100, 1)
        $msg = "Client cache: $cacheUsedMB MB used of $cacheSizeMB MB ($usedPct%)"
        if ($usedPct -gt 90) {
            Write-CMTraceLog $msg -Severity Warning
        } else {
            Write-CMTraceLog $msg -Severity Info
        }
        $results['CacheSize_MB'] = $cacheSizeMB
        $results['CacheUsed_MB'] = $cacheUsedMB
    } else {
        Write-CMTraceLog "Cache path $cachePath does not exist" -Severity Error
        $results['CacheSize_MB'] = $cacheSizeMB
        $results['CacheUsed_MB'] = 'PathMissing'
    }
} catch {
    Write-CMTraceLog "Cannot retrieve cache config: $_" -Severity Warning
    $results['CacheSize_MB'] = 'Unknown'
}

# ─── 11. Pending reboot (CBS, WU, ConfigMgr) ─────────────────────────
$rebootSources = @()
$rebootKeys = @{
    'CBS'           = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    'WindowsUpdate' = 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
}
foreach ($source in $rebootKeys.GetEnumerator()) {
    if (Test-Path $source.Value) {
        $rebootSources += $source.Key
    }
}
# ConfigMgr reboot pending
try {
    $rebootStatus = Invoke-CimMethod -Namespace 'root\ccm\clientsdk' `
        -ClassName CCM_ClientUtilities -MethodName DetermineIfRebootPending -ErrorAction Stop
    if ($rebootStatus.RebootPending -or $rebootStatus.IsHardRebootPending) {
        $rebootSources += 'ConfigMgr'
    }
} catch { }

if ($rebootSources.Count -gt 0) {
    Write-CMTraceLog "Pending reboot detected — sources: $($rebootSources -join ', ')" -Severity Warning
    $results['PendingReboot'] = $rebootSources -join ', '
} else {
    Write-CMTraceLog "No pending reboot detected" -Severity Info
    $results['PendingReboot'] = $false
}

# ─── 12. System drive free space ──────────────────────────────────────
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

Write-CMTraceLog "===== Software Update Health Check complete. Log: $LogPath ====="

[PSCustomObject]$results
