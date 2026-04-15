#Requires -Version 5.1
<#
.SYNOPSIS
    Exports a full inventory of MECM device collections to CSV.

.DESCRIPTION
    For each device collection, captures: Name, CollectionID, MemberCount,
    LimitingCollectionName, RefreshType, LastRefreshTime, LastMemberChangeTime,
    IncludeExcludeCount, QueryCount, DirectRuleCount, UseCluster (incremental),
    and Comment.

    Useful for audit, migration planning, and identifying collections that
    use incremental updates at scale (common performance risk).

.PARAMETER OutputPath
    CSV output path. Default: .\collection-inventory-<SiteCode>-<yyyyMMdd>.csv
    in the current directory.

.PARAMETER IncludeUserCollections
    Include user collections in the export. Default: device only.

.EXAMPLE
    .\Export-CollectionInventory.ps1

.EXAMPLE
    .\Export-CollectionInventory.ps1 -OutputPath C:\Audit\collections.csv -IncludeUserCollections
#>
[CmdletBinding()]
param(
    [string]$OutputPath,
    [switch]$IncludeUserCollections
)

# ─── Prompt for Site Code ─────────────────────────────────────────────
$SiteCode = Read-Host -Prompt "Enter your MECM site code (e.g., PS1)"

if (-not $OutputPath) {
    $OutputPath = Join-Path (Get-Location) ("collection-inventory-{0}-{1}.csv" -f $SiteCode.ToUpper(), (Get-Date -Format 'yyyyMMdd'))
}

# ─── Connect to CM Site ───────────────────────────────────────────────
. "$PSScriptRoot\..\Common\Connect-CMSite.ps1"
Connect-CMSite -SiteCode $SiteCode

try {
    function ConvertTo-Inventory {
        param($Collection, [string]$Type)

        $rules = $Collection.CollectionRules
        $queryCount    = ($rules | Where-Object { $_.SmsProviderObjectPath -like '*QueryRule*' }).Count
        $directCount   = ($rules | Where-Object { $_.SmsProviderObjectPath -like '*DirectRule*' }).Count
        $includeExcludeCount = ($rules | Where-Object {
            $_.SmsProviderObjectPath -like '*IncludeCollectionRule*' -or
            $_.SmsProviderObjectPath -like '*ExcludeCollectionRule*'
        }).Count

        [pscustomobject]@{
            Type                   = $Type
            Name                   = $Collection.Name
            CollectionID           = $Collection.CollectionID
            MemberCount            = $Collection.MemberCount
            LimitingCollectionName = $Collection.LimitingCollectionName
            RefreshType            = $Collection.RefreshType
            UseIncremental         = [bool]($Collection.RefreshType -band 4)
            LastRefreshTime        = $Collection.LastRefreshTime
            LastMemberChangeTime   = $Collection.LastMemberChangeTime
            QueryRules             = $queryCount
            DirectRules            = $directCount
            IncludeExcludeRules    = $includeExcludeCount
            Comment                = $Collection.Comment
        }
    }

    $Results = @()

    Write-Host "Enumerating device collections..." -ForegroundColor Cyan
    $device = Get-CMDeviceCollection
    Write-Host "  $($device.Count) device collections found." -ForegroundColor DarkGray
    foreach ($c in $device) { $Results += ConvertTo-Inventory -Collection $c -Type 'Device' }

    if ($IncludeUserCollections) {
        Write-Host "Enumerating user collections..." -ForegroundColor Cyan
        $user = Get-CMUserCollection
        Write-Host "  $($user.Count) user collections found." -ForegroundColor DarkGray
        foreach ($c in $user) { $Results += ConvertTo-Inventory -Collection $c -Type 'User' }
    }

    $Results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "Exported $($Results.Count) collections to: $OutputPath" -ForegroundColor Green

    # Flag collections using incremental updates at scale
    $incremental = $Results | Where-Object { $_.UseIncremental }
    if ($incremental.Count -gt 0) {
        Write-Host ""
        Write-Host "NOTE: $($incremental.Count) collections use incremental updates." -ForegroundColor Yellow
        Write-Host "      Review these if collection evaluation performance becomes an issue." -ForegroundColor Yellow
    }
}
finally {
    Disconnect-CMSite
}
