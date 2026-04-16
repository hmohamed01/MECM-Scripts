<#
.SYNOPSIS
    Performs a comprehensive health check of a Configuration Manager site server.

.DESCRIPTION
    Validates MECM site server health across SMS services, SQL connectivity,
    site component status, inbox backlogs, disk space, WSUS, and SMS Provider.
    Writes a CMTrace-compatible log and returns a summary object.

.PARAMETER SiteCode
    Three-character site code (e.g. "PS1"). Auto-detected if omitted.

.PARAMETER SiteServer
    FQDN of the site server. Defaults to the local computer.

.PARAMETER LogPath
    Path to the log file. Defaults to C:\Windows\Logs\CMServerHealthCheck.log.

.PARAMETER InboxBacklogThreshold
    Number of files in an inbox folder that triggers a warning. Default: 500.

.EXAMPLE
    .\Invoke-CMServerHealthCheck.ps1

.EXAMPLE
    .\Invoke-CMServerHealthCheck.ps1 -SiteCode "PS1" -SiteServer "cm01.contoso.com"

.NOTES
    Run on the site server as an admin with SMS Provider access. Read-only.
#>

[CmdletBinding()]
param(
    [string]$SiteCode,
    [string]$SiteServer = $env:COMPUTERNAME,
    [string]$LogPath = "$env:WinDir\Logs\CMServerHealthCheck.log",
    [int]$InboxBacklogThreshold = 500
)

function Write-CMTraceLog {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('Info','Warning','Error')][string]$Severity = 'Info',
        [string]$Component = 'ServerHealthCheck',
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

Write-CMTraceLog "===== Configuration Manager Server Health Check started on $SiteServer ====="

# 1. Required site server services
$siteServices = @(
    'SMS_SITE_COMPONENT_MANAGER',
    'SMS_EXECUTIVE',
    'SMS_NOTIFICATION_SERVER',
    'CcmExec',
    'IISADMIN',
    'W3SVC',
    'WsusService'
)
foreach ($svc in $siteServices) {
    $service = Get-Service -Name $svc -ComputerName $SiteServer -ErrorAction SilentlyContinue
    if (-not $service) {
        Write-CMTraceLog "Service [$svc] not installed on $SiteServer" -Severity Info
        $results["Service_$svc"] = 'NotInstalled'
        continue
    }
    if ($service.Status -eq 'Running') {
        Write-CMTraceLog "Service [$svc] status: Running" -Severity Info
        $results["Service_$svc"] = 'Running'
    } else {
        Write-CMTraceLog "Service [$svc] status: $($service.Status) — expected Running" -Severity Error
        $results["Service_$svc"] = $service.Status
    }
}

# 2. Auto-detect site code via SMS Provider
if (-not $SiteCode) {
    try {
        $provider = Get-CimInstance -Namespace 'root\sms' -ClassName SMS_ProviderLocation -ComputerName $SiteServer -ErrorAction Stop |
            Where-Object { $_.ProviderForLocalSite } | Select-Object -First 1
        if ($provider) {
            $SiteCode = $provider.SiteCode
            Write-CMTraceLog "Detected site code: $SiteCode" -Severity Info
            $results['SiteCode'] = $SiteCode
        }
    } catch {
        Write-CMTraceLog "Cannot auto-detect site code: $_" -Severity Warning
    }
}

# 3. SMS Provider namespace accessible
if ($SiteCode) {
    try {
        $site = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -ClassName SMS_Site -ComputerName $SiteServer -ErrorAction Stop
        Write-CMTraceLog "SMS Provider accessible. Site: $($site.SiteName), Version: $($site.Version)" -Severity Info
        $results['SiteName'] = $site.SiteName
        $results['SiteVersion'] = $site.Version
    } catch {
        Write-CMTraceLog "Cannot query SMS Provider namespace root\sms\site_$SiteCode : $_" -Severity Error
        $results['SMSProvider'] = 'Inaccessible'
    }
}

# 4. Site component status
if ($SiteCode) {
    try {
        $components = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -ClassName SMS_ComponentSummarizer -ComputerName $SiteServer -ErrorAction Stop
        $critical = $components | Where-Object { $_.Status -eq 2 }
        $warning = $components | Where-Object { $_.Status -eq 1 }
        Write-CMTraceLog "Site components: $($components.Count) total, $($warning.Count) warning, $($critical.Count) critical"
        foreach ($c in $critical) {
            Write-CMTraceLog "CRITICAL component: $($c.ComponentName) on $($c.MachineName)" -Severity Error
        }
        foreach ($c in $warning) {
            Write-CMTraceLog "WARNING component: $($c.ComponentName) on $($c.MachineName)" -Severity Warning
        }
        $results['ComponentsCritical'] = $critical.Count
        $results['ComponentsWarning'] = $warning.Count
    } catch {
        Write-CMTraceLog "Cannot query component summarizer: $_" -Severity Warning
    }
}

# 5. Inbox backlog check
Write-CMTraceLog "Checking inbox backlogs..."
$inboxBase = $null
try {
    $regPath = 'HKLM:\SOFTWARE\Microsoft\SMS\Identification'
    if (Test-Path $regPath) {
        $installDir = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\SMS\Setup' -ErrorAction Stop).'Installation Directory'
        $inboxBase = Join-Path $installDir 'inboxes'
    }
} catch {
    Write-CMTraceLog "Cannot locate MECM installation directory: $_" -Severity Warning
}

if ($inboxBase -and (Test-Path $inboxBase)) {
    $criticalInboxes = @(
        'auth\ddm.box',
        'auth\dataldr.box',
        'auth\statesys.box\incoming',
        'despoolr.box\receive',
        'distmgr.box',
        'inboxmgr.box',
        'schedule.box\outboxes',
        'sinv.box',
        'swmproc.box',
        'replmgr.box'
    )
    foreach ($inbox in $criticalInboxes) {
        $path = Join-Path $inboxBase $inbox
        if (Test-Path $path) {
            $count = (Get-ChildItem -Path $path -File -ErrorAction SilentlyContinue).Count
            if ($count -gt $InboxBacklogThreshold) {
                Write-CMTraceLog "Inbox backlog [$inbox]: $count files (threshold: $InboxBacklogThreshold)" -Severity Warning
            } else {
                Write-CMTraceLog "Inbox [$inbox]: $count files" -Severity Info
            }
            $results["Inbox_$($inbox -replace '[\\.]','_')"] = $count
        }
    }
} else {
    Write-CMTraceLog "Inbox directory not found — skipping backlog checks" -Severity Warning
}

# 6. Disk space on all fixed volumes
Write-CMTraceLog "Checking disk space..."
$disks = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $SiteServer -Filter "DriveType=3" -ErrorAction SilentlyContinue
foreach ($disk in $disks) {
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $totalGB = [math]::Round($disk.Size / 1GB, 2)
    $freePct = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1)
    $msg = "Drive $($disk.DeviceID) $freeGB GB free of $totalGB GB ($freePct%)"
    if ($freePct -lt 10) {
        Write-CMTraceLog $msg -Severity Error
    } elseif ($freePct -lt 20) {
        Write-CMTraceLog $msg -Severity Warning
    } else {
        Write-CMTraceLog $msg -Severity Info
    }
    $results["Disk_$($disk.DeviceID -replace ':','')_FreeGB"] = $freeGB
}

# 7. SQL connectivity (site database)
if ($SiteCode) {
    try {
        $siteDef = Get-CimInstance -Namespace "root\sms\site_$SiteCode" -ClassName SMS_SCI_SiteDefinition -ComputerName $SiteServer -ErrorAction Stop |
            Where-Object { $_.SiteCode -eq $SiteCode } | Select-Object -First 1
        if ($siteDef) {
            $sqlServer = $siteDef.SQLServerName
            $sqlDatabase = "CM_$SiteCode"
            Write-CMTraceLog "Testing SQL connection to $sqlServer\$sqlDatabase..."
            try {
                $conn = New-Object System.Data.SqlClient.SqlConnection
                $conn.ConnectionString = "Server=$sqlServer;Database=$sqlDatabase;Integrated Security=True;Connect Timeout=10"
                $conn.Open()
                $cmd = $conn.CreateCommand()
                $cmd.CommandText = "SELECT TOP 1 SiteCode FROM v_Site"
                $siteResult = $cmd.ExecuteScalar()
                $conn.Close()
                Write-CMTraceLog "SQL connection successful. Queried site: $siteResult" -Severity Info
                $results['SQLConnection'] = 'OK'
            } catch {
                Write-CMTraceLog "SQL connection failed: $_" -Severity Error
                $results['SQLConnection'] = 'Failed'
            }
        }
    } catch {
        Write-CMTraceLog "Cannot retrieve SQL server info from site definition: $_" -Severity Warning
    }
}

# 8. Recent critical events in Application log (SMS source)
try {
    $since = (Get-Date).AddHours(-24)
    $smsEvents = Get-WinEvent -FilterHashtable @{
        LogName = 'Application'
        ProviderName = 'SMS*'
        Level = 1,2
        StartTime = $since
    } -ErrorAction SilentlyContinue
    if ($smsEvents) {
        Write-CMTraceLog "Found $($smsEvents.Count) SMS error/critical events in the last 24 hours" -Severity Warning
        $smsEvents | Select-Object -First 5 | ForEach-Object {
            Write-CMTraceLog "Event $($_.Id) [$($_.ProviderName)]: $($_.Message -replace '\r?\n',' ' | ForEach-Object { $_.Substring(0, [Math]::Min(200, $_.Length)) })" -Severity Warning
        }
        $results['RecentSMSErrors'] = $smsEvents.Count
    } else {
        Write-CMTraceLog "No SMS error/critical events in the last 24 hours" -Severity Info
        $results['RecentSMSErrors'] = 0
    }
} catch {
    Write-CMTraceLog "Cannot query event log: $_" -Severity Warning
}

Write-CMTraceLog "===== Configuration Manager Server Health Check complete. Log: $LogPath ====="

[PSCustomObject]$results
