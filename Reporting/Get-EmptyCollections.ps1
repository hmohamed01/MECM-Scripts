#Requires -Version 5.1
<#
.SYNOPSIS
    Reports MECM device and user collections with zero members.

.DESCRIPTION
    Scans every device collection and user collection, returns those with
    MemberCount = 0. Prints a summary to the console and optionally exports
    to CSV for cleanup review.

    Collections with MemberCount = 0 are not always dead — they may be
    newly created, feed empty queries temporarily, or exist as limiting
    containers. Review before removing.

.PARAMETER OutputPath
    Optional CSV path. If specified, the result list is exported.
    Default: no export (console only).

.PARAMETER IncludeUserCollections
    Include user collections in addition to device collections. Default: false.

.EXAMPLE
    .\Get-EmptyCollections.ps1

.EXAMPLE
    .\Get-EmptyCollections.ps1 -OutputPath C:\Temp\empty-collections.csv -IncludeUserCollections
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$IncludeUserCollections
)

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"

# ─── Connect to CM Site ───────────────────────────────────────────────
. "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
Connect-CMSite -SiteCode $SiteCode

try {
    Write-Host "Enumerating device collections..." -ForegroundColor Cyan
    $DeviceCollections = Get-CMDeviceCollection |
        Where-Object { $_.MemberCount -eq 0 } |
        Select-Object @{N='Type';E={'Device'}},
                      Name,
                      CollectionID,
                      LimitingCollectionName,
                      MemberCount,
                      @{N='LastRefreshTime';E={$_.LastRefreshTime}},
                      @{N='RefreshType';E={$_.RefreshType}},
                      Comment

    $Results = @($DeviceCollections)

    if ($IncludeUserCollections) {
        Write-Host "Enumerating user collections..." -ForegroundColor Cyan
        $UserCollections = Get-CMUserCollection |
            Where-Object { $_.MemberCount -eq 0 } |
            Select-Object @{N='Type';E={'User'}},
                          Name,
                          CollectionID,
                          LimitingCollectionName,
                          MemberCount,
                          @{N='LastRefreshTime';E={$_.LastRefreshTime}},
                          @{N='RefreshType';E={$_.RefreshType}},
                          Comment
        $Results += @($UserCollections)
    }

    # ─── Output ───────────────────────────────────────────────────────
    if ($Results.Count -eq 0) {
        Write-Host "No empty collections found." -ForegroundColor Green
    }
    else {
        Write-Host ""
        Write-Host "Empty collections found: $($Results.Count)" -ForegroundColor Yellow
        Write-Host ""
        $Results | Format-Table Type, Name, CollectionID, LimitingCollectionName, LastRefreshTime -AutoSize

        if ($OutputPath) {
            $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
            Write-Host "Exported to: $OutputPath" -ForegroundColor Green
        }
    }
}
finally {
    Disconnect-CMSite
}
