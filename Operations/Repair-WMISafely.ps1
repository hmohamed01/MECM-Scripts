<#
.SYNOPSIS
    Safely repairs WMI without rebuilding the repository.

.DESCRIPTION
    Performs a non-destructive WMI repair workflow:
      1. Verifies repository consistency
      2. Salvages the repository (repairs in place, no data loss)
      3. Re-registers WMI provider DLLs
      4. Recompiles critical MOF files
      5. Re-verifies repository after repair

    This script deliberately avoids `winmgmt /resetrepository`, which is
    destructive and can break ConfigMgr, SCOM, and other agents.

.PARAMETER LogPath
    Path to the log file. Defaults to C:\Windows\Logs\WMIRepair.log.

.PARAMETER SkipServiceRestart
    Do not restart the Winmgmt service after repair. Use if you want to
    schedule a manual restart.

.EXAMPLE
    .\Repair-WMISafely.ps1

.EXAMPLE
    .\Repair-WMISafely.ps1 -LogPath "C:\Temp\WMIRepair.log"

.NOTES
    Must run as Administrator. Does NOT use /resetrepository (destructive).
    If this script cannot restore WMI health, escalate to a full rebuild
    only as a last resort.
#>

[CmdletBinding()]
param(
    [string]$LogPath = "$env:WinDir\Logs\WMIRepair.log",
    [switch]$SkipServiceRestart
)

#Requires -RunAsAdministrator

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warning','Error')][string]$Severity = 'Info',
        [string]$Component = 'WMIRepair',
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
Write-CMTraceLog "===== Safe WMI Repair started on $env:COMPUTERNAME ====="
Write-CMTraceLog "This script performs NON-DESTRUCTIVE repair. It will NOT rebuild the repository." -Severity Info

# Step 1: Initial verification
Write-CMTraceLog "Step 1: Verifying WMI repository consistency..."
$initialVerify = & winmgmt /verifyrepository 2>&1
$initialStatus = $LASTEXITCODE
Write-CMTraceLog "Initial verification: $initialVerify (ExitCode: $initialStatus)"

if ($initialStatus -eq 0) {
    Write-CMTraceLog "WMI repository is already consistent. No repair required." -Severity Info
    Write-CMTraceLog "Proceeding with provider re-registration as preventive maintenance..."
}

# Step 2: Stop dependent services that can interfere with salvage
Write-CMTraceLog "Step 2: Stopping dependent services..."
$dependentServices = @('CcmExec','ccmsetup','SMS Agent Host','Winmgmt')
$stoppedServices = @()
foreach ($svc in $dependentServices) {
    $service = Get-Service -Name $svc -ErrorAction SilentlyContinue
    if ($service -and $service.Status -eq 'Running') {
        try {
            Stop-Service -Name $svc -Force -ErrorAction Stop
            Write-CMTraceLog "Stopped service: $svc" -Severity Info
            $stoppedServices += $svc
        } catch {
            Write-CMTraceLog "Failed to stop $svc : $_" -Severity Warning
        }
    }
}

# Step 3: Salvage the repository (non-destructive repair)
Write-CMTraceLog "Step 3: Running winmgmt /salvagerepository (non-destructive)..."
try {
    $salvageResult = & winmgmt /salvagerepository 2>&1
    $salvageExit = $LASTEXITCODE
    if ($salvageExit -eq 0) {
        Write-CMTraceLog "Salvage result: $salvageResult" -Severity Info
    } else {
        Write-CMTraceLog "Salvage returned non-zero: $salvageResult (ExitCode: $salvageExit)" -Severity Warning
    }
} catch {
    Write-CMTraceLog "Salvage operation failed: $_" -Severity Error
}

# Step 4: Re-register WMI provider DLLs
Write-CMTraceLog "Step 4: Re-registering WMI provider DLLs..."
$wmiDlls = @(
    "$env:WinDir\System32\wbem\wmiprvsd.dll",
    "$env:WinDir\System32\wbem\wmisvc.dll",
    "$env:WinDir\System32\wbem\wbemcore.dll",
    "$env:WinDir\System32\wbem\wbemprox.dll",
    "$env:WinDir\System32\wbem\wmiutils.dll",
    "$env:WinDir\System32\scrcons.exe"
)
foreach ($dll in $wmiDlls) {
    if (Test-Path $dll) {
        try {
            if ($dll -like '*.exe') {
                $result = & $dll /regserver 2>&1
            } else {
                $result = & regsvr32.exe /s $dll 2>&1
            }
            Write-CMTraceLog "Registered: $dll" -Severity Info
        } catch {
            Write-CMTraceLog "Failed to register $dll : $_" -Severity Warning
        }
    } else {
        Write-CMTraceLog "DLL not found: $dll" -Severity Warning
    }
}

# Step 5: Recompile critical MOF files
Write-CMTraceLog "Step 5: Recompiling critical MOF files..."
$criticalMofs = @(
    'cimwin32.mof',
    'cimwin32.mfl',
    'rsop.mof',
    'rsop.mfl',
    'wmi\wmipcima.mof',
    'wmi\wmipcima.mfl'
)
$wbemPath = "$env:WinDir\System32\wbem"
foreach ($mof in $criticalMofs) {
    $fullPath = Join-Path $wbemPath $mof
    if (Test-Path $fullPath) {
        try {
            $result = & "$wbemPath\mofcomp.exe" $fullPath 2>&1
            Write-CMTraceLog "Compiled: $mof" -Severity Info
        } catch {
            Write-CMTraceLog "Failed to compile $mof : $_" -Severity Warning
        }
    }
}

# Step 6: Restart Winmgmt and dependent services
if (-not $SkipServiceRestart) {
    Write-CMTraceLog "Step 6: Restarting Winmgmt service..."
    try {
        Start-Service -Name Winmgmt -ErrorAction Stop
        Write-CMTraceLog "Winmgmt started" -Severity Info
        Start-Sleep -Seconds 3
        foreach ($svc in $stoppedServices) {
            if ($svc -eq 'Winmgmt') { continue }
            try {
                Start-Service -Name $svc -ErrorAction Stop
                Write-CMTraceLog "Restarted service: $svc" -Severity Info
            } catch {
                Write-CMTraceLog "Failed to restart $svc : $_" -Severity Warning
            }
        }
    } catch {
        Write-CMTraceLog "Failed to start Winmgmt: $_" -Severity Error
    }
} else {
    Write-CMTraceLog "Step 6: Skipping service restart (SkipServiceRestart flag set)" -Severity Info
}

# Step 7: Final verification
Write-CMTraceLog "Step 7: Re-verifying WMI repository consistency..."
Start-Sleep -Seconds 2
$finalVerify = & winmgmt /verifyrepository 2>&1
$finalStatus = $LASTEXITCODE
Write-CMTraceLog "Final verification: $finalVerify (ExitCode: $finalStatus)"

# Step 8: Test a basic WMI query
try {
    $os = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    Write-CMTraceLog "Post-repair WMI query succeeded: $($os.Caption)" -Severity Info
    $wmiHealthy = $true
} catch {
    Write-CMTraceLog "Post-repair WMI query failed: $_" -Severity Error
    $wmiHealthy = $false
}

Write-CMTraceLog "===== Safe WMI Repair complete ====="
if ($finalStatus -eq 0 -and $wmiHealthy) {
    Write-CMTraceLog "Result: WMI is healthy" -Severity Info
    Write-Host "`nWMI repair completed successfully. Log: $LogPath" -ForegroundColor Green
} else {
    Write-CMTraceLog "Result: WMI still unhealthy — manual intervention required" -Severity Error
    Write-Host "`nWMI repair did not fully restore health. Review log: $LogPath" -ForegroundColor Red
    Write-Host "If WMI remains broken, a full repository rebuild may be required." -ForegroundColor Yellow
    Write-Host "WARNING: A rebuild will break ConfigMgr, SCOM, and other WMI-dependent agents." -ForegroundColor Yellow
}
