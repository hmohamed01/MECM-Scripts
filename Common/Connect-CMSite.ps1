#Requires -Version 5.1
<#
.SYNOPSIS
    Shared helper for connecting to a MECM site from PowerShell scripts.

.DESCRIPTION
    Dot-source this file at the top of any script. Exposes two functions:

        Connect-CMSite   Loads the ConfigurationManager module with a fallback
                         chain, switches to the <SiteCode>:\ PSDrive, and stores
                         the original location in a script-scope variable.

        Disconnect-CMSite  Restores the original location. Call at the end of
                           the script (including error paths).

    Usage:
        . "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
        Connect-CMSite -SiteCode $SiteCode
        try {
            # ... script logic ...
        }
        finally {
            Disconnect-CMSite
        }

.NOTES
    Requires the MECM console installed on the machine running the script.
#>

function Connect-CMSite {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteCode
    )

    if ([string]::IsNullOrWhiteSpace($SiteCode)) {
        throw "Site code cannot be empty."
    }

    $Script:CMSiteCode       = $SiteCode.Trim().ToUpper()
    $Script:CMOriginalLocation = Get-Location
    $Script:CMFolderRoot     = "$($Script:CMSiteCode):\"

    # Load module: prefer SMS_ADMIN_UI_PATH, fall back to Import-Module
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
            throw "Failed to load ConfigMgr module. Is the MECM console installed? $($_.Exception.Message)"
        }
    }

    # Switch to the CM drive
    Set-Location "$($Script:CMSiteCode):\"
    Write-Host "Connected to site $($Script:CMSiteCode)." -ForegroundColor Green
}

function Disconnect-CMSite {
    if ($Script:CMOriginalLocation) {
        Set-Location $Script:CMOriginalLocation
    }
}
