#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Online operations module for SPOVC version management
.DESCRIPTION
    Handles version retrieval, deletion, and inactive site detection.
    Uses PnP CSOM for URL-safe version operations and batch processing.
.NOTES
    Dependencies: Core.psm1 (Write-Log, Format-FileSize), Authentication.psm1 (Connect-SharePoint)
#>

<#
.FUNCTION Get-FileVersions
.SYNOPSIS
    Retrieves all versions for a list item using CSOM
.DESCRIPTION
    Gets file versions from SharePoint using PnP and CSOM to avoid URL encoding issues.
    Returns versions sorted by Created date (descending - newest first).
.PARAMETER Item
    PnP list item from Get-PnPListItem
.EXAMPLE
    $item = Get-PnPListItem -List "Documents" -Id 1
    $versions = Get-FileVersions -Item $item
    $versions | Select-Object Created, @{N="Size";E={Format-FileSize -Size $_.Size}}
#>
function Get-FileVersions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$Item
    )

    try {
        # Skip if item doesn't have a File (e.g., it's a folder)
        if ($Item.File -eq $null) {
            return @()
        }

        # Get file object from item
        $file = Get-PnPProperty -ClientObject $Item -Property File
        
        # Execute the query to load the File property
        Invoke-PnPQuery
        
        # Check if file is null after query
        if ($file -eq $null) {
            return @()
        }

        # Get versions using CSOM
        $versions = Get-PnPProperty -ClientObject $file -Property Versions
        
        # Execute the query to load versions
        Invoke-PnPQuery

        if ($versions.Count -eq 0) {
            return @()
        }

        # Convert to objects and sort by Created (newest first)
        $versionList = @()
        foreach ($version in $versions) {
            $versionList += @{
                VersionLabel = $version.VersionLabel
                Created      = $version.Created
                Size         = $version.Size
                CheckInComment = $version.CheckInComment
                CreatedBy    = $version.CreatedBy.LoginName
            }
        }

        return $versionList | Sort-Object Created -Descending
    } catch {
        # Silently skip items that can't be processed (folders, special items, etc.)
        return @()
    }
}

<#
.FUNCTION Remove-FileVersions
.SYNOPSIS
    Deletes excess versions from a file using CSOM
.DESCRIPTION
    Keeps only specified number of versions (newest first).
    Uses Invoke-PnPQuery for batch deletion efficiency.
    Can operate in DRYRUN mode without making changes.
.PARAMETER Item
    PnP list item from Get-PnPListItem
.PARAMETER VersionsToKeep
    Number of recent versions to retain
.PARAMETER Mode
    Operation mode: "SCAN" (report only), "DRYRUN" (calculate), "CLEANUP" (delete)
.EXAMPLE
    $deleted = Remove-FileVersions -Item $item -VersionsToKeep 25 -Mode "CLEANUP"
    $deleted | Select-Object VersionLabel, Size
#>
function Remove-FileVersions {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject]$Item,

        [Parameter(Mandatory = $true)]
        [int]$VersionsToKeep,

        [Parameter(Mandatory = $true)]
        [ValidateSet("SCAN", "DRYRUN", "CLEANUP")]
        [string]$Mode
    )

    $deleted = @()
    $totalSaved = 0

    try {
        # Get the file object and its versions using CSOM
        $file = Get-PnPProperty -ClientObject $Item -Property File
        Invoke-PnPQuery
        
        $versions = Get-PnPProperty -ClientObject $file -Property Versions
        Invoke-PnPQuery

        if ($versions.Count -le $VersionsToKeep) {
            return @()
        }

        # Get versions sorted newest first
        $versionArray = @()
        foreach ($version in $versions) {
            $versionArray += $version
        }
        $versionArray = $versionArray | Sort-Object Created -Descending

        # Identify excess versions (beyond VersionsToKeep)
        $versionsToDelete = $versionArray | Select-Object -Skip $VersionsToKeep

        if ($Mode -eq "CLEANUP") {
            # Batch delete using CSOM
            $ctx = Get-PnPContext
            foreach ($version in $versionsToDelete) {
                try {
                    $version.Delete()
                    $deleted += @{
                        VersionLabel = $version.VersionLabel
                        Size = $version.Size
                    }
                    $totalSaved += $version.Size
                } catch {
                    Write-Log "Error deleting version $($version.VersionLabel): $($_.Exception.Message)" "WARN"
                }
            }

            # Batch execute
            if ($deleted.Count -gt 0) {
                Invoke-PnPQuery
                Write-Log "Deleted $($deleted.Count) versions from $($Item.FieldValues.FileLeafRef)" "INFO"
            }
        } else {
            # SCAN or DRYRUN mode - just report
            foreach ($version in $versionsToDelete) {
                $deleted += @{
                    VersionLabel = $version.VersionLabel
                    Size = $version.Size
                }
                $totalSaved += $version.Size
            }

            if ($Mode -eq "DRYRUN") {
                Write-Log "DRYRUN: Would delete $($deleted.Count) versions from $($Item.FieldValues.FileLeafRef), saving $(Format-FileSize -Size $totalSaved)" "INFO"
            }
        }

        # Return results with storage info
        return @{
            Deleted = $deleted
            StorageSaved = $totalSaved
            Count = $deleted.Count
        }

    } catch {
        Write-Log "Error in Remove-FileVersions for item $($Item.Id): $($_.Exception.Message)" "ERROR"
        return @{
            Deleted = @()
            StorageSaved = 0
            Count = 0
        }
    }
}

<#
.FUNCTION Get-InactiveSites
.SYNOPSIS
    Identifies inactive SharePoint sites based on last modified date
.DESCRIPTION
    Queries all sites and filters by last activity timestamp.
    Returns sites with configuration-based inactivity threshold.
.PARAMETER DaysInactive
    Number of days to consider a site inactive (default from config)
.PARAMETER MaxStorageMB
    Optional storage threshold for cleanup candidates (in MB)
.EXAMPLE
    $inactive = Get-InactiveSites -DaysInactive 90 -MaxStorageMB 100
    $inactive | Select-Object Url, Title, StorageMB
#>
function Get-InactiveSites {
    [CmdletBinding()]
    param(
        [Parameter()]
        [int]$DaysInactive = 90,

        [Parameter()]
        [int]$MaxStorageMB = 0
    )

    $inactiveSites = @()
    $inactiveThreshold = (Get-Date).AddDays(-$DaysInactive)

    try {
        Write-Log "Retrieving inactive sites (threshold: $DaysInactive days)" "INFO"

        $sites = Get-PnPTenantSite -Detailed | Where-Object {$_.Status -eq "Active"}

        foreach ($site in $sites) {
            # Get last modified date
            $lastModified = $site.LastContentModifiedDate
            if ($null -eq $lastModified) {
                $lastModified = $site.LastModifiedTime
            }

            # Check inactivity
            if ($lastModified -lt $inactiveThreshold) {
                $siteObj = @{
                    Url            = $site.Url
                    Title          = $site.Title
                    LastModified   = $lastModified
                    StorageMB      = [math]::Round($site.StorageUsageCurrent / 1024, 2)
                    DaysInactive   = ([math]::Round((Get-Date - $lastModified).TotalDays, 0))
                }

                # Filter by storage if specified
                if ($MaxStorageMB -eq 0 -or $siteObj.StorageMB -le $MaxStorageMB) {
                    $inactiveSites += $siteObj
                }
            }
        }

        Write-Log "Found $($inactiveSites.Count) inactive sites" "INFO"
        return $inactiveSites

    } catch {
        Write-Log "Error retrieving inactive sites: $($_.Exception.Message)" "ERROR"
        return @()
    }
}

<#
.FUNCTION Start-InactiveSitesDeletion
.SYNOPSIS
    Deletes versions in inactive sites
.DESCRIPTION
    Processes multiple inactive sites for version cleanup.
    Uses Get-FileVersions and Remove-FileVersions on each site.
.PARAMETER InactiveSites
    Array of site objects from Get-InactiveSites
.PARAMETER VersionsToKeep
    Number of versions to retain per file
.PARAMETER Mode
    Operation mode: "SCAN", "DRYRUN", or "CLEANUP"
.EXAMPLE
    $inactive = Get-InactiveSites -DaysInactive 90
    $results = Start-InactiveSitesDeletion -InactiveSites $inactive -VersionsToKeep 10 -Mode "CLEANUP"
#>
function Start-InactiveSitesDeletion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [PSObject[]]$InactiveSites,

        [Parameter(Mandatory = $true)]
        [int]$VersionsToKeep,

        [Parameter(Mandatory = $true)]
        [ValidateSet("SCAN", "DRYRUN", "CLEANUP")]
        [string]$Mode
    )

    $results = @{
        TotalSites = $InactiveSites.Count
        ProcessedSites = 0
        TotalVersionsDeleted = 0
        TotalStorageSaved = 0
        Errors = @()
    }

    try {
        foreach ($site in $InactiveSites) {
            Write-Log "Processing inactive site: $($site.Url) (inactive for $($site.DaysInactive) days)" "INFO"

            try {
                Connect-SharePoint -SiteUrl $site.Url

                $libraries = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101 -or $_.BaseTemplate -eq 109}

                foreach ($library in $libraries) {
                    $items = Get-PnPListItem -List $library.Id -PageSize 5000

                    foreach ($item in $items) {
                        $versionResult = Remove-FileVersions -Item $item -VersionsToKeep $VersionsToKeep -Mode $Mode
                        $results.TotalVersionsDeleted += $versionResult.Count
                        $results.TotalStorageSaved += $versionResult.StorageSaved
                    }
                }

                $results.ProcessedSites++

            } catch {
                $results.Errors += @{
                    SiteUrl = $site.Url
                    Error   = $_.Exception.Message
                }
                Write-Log "Error processing site $($site.Url): $($_.Exception.Message)" "ERROR"
            }
        }

    } catch {
        Write-Log "Error in Start-InactiveSitesDeletion: $($_.Exception.Message)" "ERROR"
    }

    Write-Log "Inactive site cleanup complete: $($results.ProcessedSites)/$($results.TotalSites) sites processed, $($results.TotalVersionsDeleted) versions deleted, $(Format-FileSize -Size $results.TotalStorageSaved) saved" "INFO"

    return $results
}

# Export public functions
Export-ModuleMember -Function @(
    'Get-FileVersions',
    'Remove-FileVersions',
    'Get-InactiveSites',
    'Start-InactiveSitesDeletion'
)

