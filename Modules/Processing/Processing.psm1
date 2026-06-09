#Requires -Version 7.0
<#
.SYNOPSIS
    Background job orchestration and processing module for SPOVC
.DESCRIPTION
    Manages job creation, monitoring, and thread-safe UI updates for version cleanup operations.
    Handles progress tracking and communication between background jobs and main UI thread.
.NOTES
    Dependencies: Core.psm1 (Write-Log), Authentication.psm1 (Connect-SharePoint), Reporting.psm1 (Add-SessionOperation)
#>

# Import dependencies (in main script context)
# Import-Module "Modules/Core/Core.psm1" -Force
# Import-Module "Modules/Reporting/Reporting.psm1" -Force
# Import-Module "Modules/Authentication/Authentication.psm1" -Force

<#
.FUNCTION Start-ProcessingWithProgress
.SYNOPSIS
    Creates progress dialog and launches background job for version cleanup
.DESCRIPTION
    Creates a modal progress form with progress bar, details textbox, and cancel button.
    Launches background job scriptblock that processes sites/libraries/versions.
    Monitors job output and updates UI with progress.
.PARAMETER Mode
    Operation mode: "SCAN", "DRYRUN", or "CLEANUP"
.PARAMETER Scope
    "AllSites" or "SingleSite"
.PARAMETER SiteUrl
    (Optional) Site URL if Scope="SingleSite"
.PARAMETER LibraryName
    (Optional) Library name if Scope="SingleSite"
.EXAMPLE
    Start-ProcessingWithProgress -Mode "SCAN" -Scope "AllSites"
#>
function Start-ProcessingWithProgress {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("SCAN", "ESTIMATE", "DRYRUN", "CLEANUP")]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [ValidateSet("AllSites", "SingleSite", "SelectedSites")]
        [string]$Scope,

        [Parameter()]
        [string]$SiteUrl,

        [Parameter()]
        [string]$LibraryName,

        [Parameter()]
        [object[]]$SelectedSites = @()
    )

    try {
        Write-Log "Starting ProcessingWithProgress: Mode=$Mode, Scope=$Scope" "INFO"

        # Create progress form
        $progressForm = New-Object System.Windows.Forms.Form
        $progressForm.Text = "SPOVC - Processing ($Mode)"
        $progressForm.Width = 500
        $progressForm.Height = 300
        $progressForm.StartPosition = "CenterScreen"
        $progressForm.TopMost = $true
        $progressForm.FormBorderStyle = "FixedDialog"
        $progressForm.MaximizeBox = $false
        $progressForm.MinimizeBox = $false

        # Progress bar
        $progressBar = New-Object System.Windows.Forms.ProgressBar
        $progressBar.Left = 10
        $progressBar.Top = 10
        $progressBar.Width = 470
        $progressBar.Height = 30
        $progressBar.Style = "Marquee"
        $progressForm.Controls.Add($progressBar)

        # Status label
        $statusLabel = New-Object System.Windows.Forms.Label
        $statusLabel.Left = 10
        $statusLabel.Top = 50
        $statusLabel.Width = 470
        $statusLabel.Height = 20
        $statusLabel.Text = "Initializing..."
        $progressForm.Controls.Add($statusLabel)

        # Details textbox
        $detailsBox = New-Object System.Windows.Forms.TextBox
        $detailsBox.Left = 10
        $detailsBox.Top = 80
        $detailsBox.Width = 470
        $detailsBox.Height = 150
        $detailsBox.Multiline = $true
        $detailsBox.ScrollBars = "Vertical"
        $detailsBox.ReadOnly = $true
        $progressForm.Controls.Add($detailsBox)

        # Cancel button
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Left = 410
        $cancelButton.Top = 240
        $cancelButton.Width = 70
        $cancelButton.Height = 25
        $cancelButton.Text = "Cancel"
        $cancelButton.Enabled = $true
        $progressForm.Controls.Add($cancelButton)

        $global:cancelRequested = $false
        $cancelButton.Add_Click({
            if ($cancelButton.Text -eq "Done" -or $cancelButton.Text -eq "Close") {
                # Close the form when Done/Close is clicked
                $progressForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
                $progressForm.Close()
            } else {
                # Request cancellation during processing
                $global:cancelRequested = $true
                $cancelButton.Enabled = $false
                $statusLabel.Text = "Cancelling..."
            }
        })

        # Show form (non-blocking with DoEvents loop)
        $progressForm.Show()
        [System.Windows.Forms.Application]::DoEvents()

        # Build job scriptblock
        $jobScriptBlock = {
            param($Mode, $Scope, $SiteUrl, $LibraryName, $GlobalAuthToken, $GlobalSPToken, $AdminUrl, $KeepVersions, $ModuleBasePath)

            # Import modules in job context using absolute paths
            $corePath = Join-Path $ModuleBasePath "Core/Core.psm1"
            $authPath = Join-Path $ModuleBasePath "Authentication/Authentication.psm1"
            $reportingPath = Join-Path $ModuleBasePath "Reporting/Reporting.psm1"
            $opsPath = Join-Path $ModuleBasePath "Processing/SharePointOps.psm1"
            $procPath = Join-Path $ModuleBasePath "Processing/Processing.psm1"
            
            Write-Host "Loading modules from: $ModuleBasePath"
            Write-Host "  Core: $corePath"
            
            Import-Module $corePath -Force -ErrorAction Stop
            Import-Module $authPath -Force -ErrorAction Stop
            Import-Module $reportingPath -Force -ErrorAction Stop
            Import-Module $opsPath -Force -ErrorAction Stop
            Import-Module $procPath -Force -ErrorAction Stop
            
            Write-Host "Modules loaded successfully in background job"

            # Restore auth tokens and configuration
            $global:globalAccessToken = $GlobalAuthToken
            $global:globalSharePointToken = $GlobalSPToken
            $global:adminUrl = $AdminUrl
            $global:keepVersions = $KeepVersions
            $global:isGloballyAuthenticated = $true

            # Authenticate to admin URL to enable tenant operations
            try {
                Write-Host "Connecting to admin site for authentication..."
                
                # Check if we have a real token or connection-based auth
                if ($GlobalSPToken -eq "CONNECTION_BASED_AUTH") {
                    Write-Host "Using connection-based authentication (no token available)"
                    # For background jobs with CONNECTION_BASED_AUTH, we need to use interactive auth
                    # or fall back to single-site operations
                    Write-Host "WARNING: Connection-based auth in background job. Single-site operations only."
                } else {
                    Connect-PnPOnline -Url $AdminUrl -AccessToken $GlobalSPToken -ErrorAction Stop
                    Write-Host "Successfully authenticated to admin site with token"
                }
            } catch {
                throw "Background job authentication failed: $($_.Exception.Message)"
            }

            $results = @{
                Mode      = $Mode
                Scope     = $Scope
                Sites     = @()
                TotalVersionsDeleted = 0
                TotalStorageSaved    = 0
                Errors    = @()
            }

            try {
                if ($Scope -eq "AllSites") {
                    # Check if we have token-based auth
                    if ($GlobalSPToken -eq "CONNECTION_BASED_AUTH") {
                        Write-Host "ERROR: Cannot list all sites without token-based authentication."
                        Write-Host "Please use 'Scan Single Site' mode instead."
                        throw "AllSites mode requires token-based authentication. This account cannot retrieve tokens. Use 'Scan Single Site' mode."
                    }
                    
                    # Process all sites
                    Write-Host "Retrieving all sites..."
                    try {
                        $sites = Get-PnPTenantSite -Detailed -ErrorAction Stop | Where-Object {$_.Status -eq "Active"}
                    } catch {
                        # Check if it's a permission error
                        if ($_.Exception.Message -match "Unauthorized|permission|Access Denied") {
                            throw "PERMISSION ERROR: Your account or Entra app lacks permissions to list all sites.`n`nSolution: Use 'Scan Single Site' mode instead, or ensure your Entra app has 'Sites.Read.All' permission.`n`nTechnical details: $_"
                        } else {
                            throw $_
                        }
                    }

                    foreach ($site in $sites) {
                        Write-Host "Processing site: $($site.Url)"
                        $siteResult = Invoke-SiteVersionCleanup -SiteUrl $site.Url -Mode $Mode
                        $results.Sites += $siteResult
                        $results.TotalVersionsDeleted += $siteResult.VersionsDeleted
                        $results.TotalStorageSaved += $siteResult.StorageSaved
                    }
                } else {
                    # Process single site
                    Write-Host "Processing single site: $SiteUrl"
                    $siteResult = Invoke-SiteVersionCleanup -SiteUrl $SiteUrl -LibraryName $LibraryName -Mode $Mode
                    $results.Sites += $siteResult
                    $results.TotalVersionsDeleted = $siteResult.VersionsDeleted
                    $results.TotalStorageSaved = $siteResult.StorageSaved
                }
            } catch {
                $results.Errors += @{
                    Site  = $SiteUrl
                    Error = $_.Exception.Message
                }
                Write-Host "Error: $($_.Exception.Message)"
            }

            return $results
        }

        # For AllSites and SelectedSites operations, run on main thread (already authenticated)
        # For SingleSite operations, background job is fine
        if ($Scope -eq "AllSites" -or $Scope -eq "SelectedSites") {
            Write-Log "Running $Scope operation on main thread for proper authentication" "INFO"
            $results = Start-MainThreadProcessing -Mode $Mode -Scope $Scope -SiteUrl $SiteUrl -LibraryName $LibraryName -SelectedSites $SelectedSites -ProgressForm $progressForm -StatusLabel $statusLabel -DetailsBox $detailsBox -ProgressBar $progressBar -CancelButton $cancelButton
            
            # Add to session statistics  
            if (Get-Command Add-SessionOperation -ErrorAction SilentlyContinue) {
                Add-SessionOperation -Operation $Mode -SitesProcessed $results.Sites.Count -VersionsDeleted $results.TotalVersionsDeleted -StorageSaved $results.TotalStorageSaved
            }
            
            return $results
        }

        # For single-site operations, use background job
        $moduleBasePath = Split-Path -Parent $PSScriptRoot
        
        Write-Log "Module base path for job: $moduleBasePath" "INFO"
        
        $jobParams = @{
            ScriptBlock = $jobScriptBlock
            ArgumentList = @(
                $Mode,
                $Scope,
                $SiteUrl,
                $LibraryName,
                $global:globalAccessToken,
                $global:globalSharePointToken,
                $global:adminUrl,
                $global:keepVersions,
                $moduleBasePath
            )
        }

        $job = Start-Job @jobParams
        Write-Log "Background job started: $($job.Id)" "INFO"

        # Monitor job output and update UI
        $lastStatusUpdate = Get-Date
        while ($job.State -eq "Running" -and -not $global:cancelRequested) {
            [System.Windows.Forms.Application]::DoEvents()

            # Get job output
            $output = Receive-Job -Job $job -ErrorVariable jobError -ErrorAction SilentlyContinue
            if ($output) {
                foreach ($line in @($output)) {
                    $detailsBox.AppendText("$line`r`n")
                    $detailsBox.ScrollToCaret()
                    $statusLabel.Text = $line
                }
            }

            # Update progress bar
            if ((Get-Date) - $lastStatusUpdate | Select-Object -ExpandProperty TotalSeconds | Where-Object {$_ -gt 0.5}) {
                $progressBar.Value = ($progressBar.Value + 10) % 100
                $lastStatusUpdate = Get-Date
            }

            Start-Sleep -Milliseconds 500
        }

        # Cancel job if requested
        if ($global:cancelRequested) {
            Stop-Job -Job $job
            Remove-Job -Job $job
            Write-Log "Job cancelled by user" "WARN"
            $detailsBox.AppendText("`r`nOperation cancelled.`r`n")
            $progressForm.Close()
            return
        }

        # Wait for job completion
        Wait-Job -Job $job | Out-Null

        # Get final results
        $finalResults = Receive-Job -Job $job -Wait -ErrorVariable jobError
        Remove-Job -Job $job

        # Update UI with final results
        $detailsBox.AppendText("`r`n--- OPERATION COMPLETE ---`r`n")
        $detailsBox.AppendText("Total versions deleted: $($finalResults.TotalVersionsDeleted)`r`n")
        $detailsBox.AppendText("Total storage saved: $(Format-FileSize -Size $finalResults.TotalStorageSaved)`r`n")

        if ($finalResults.Errors.Count -gt 0) {
            $detailsBox.AppendText("`r`n⚠️  ERRORS ENCOUNTERED ($($finalResults.Errors.Count)):`r`n")
            foreach ($error in $finalResults.Errors) {
                $detailsBox.AppendText("`r`n[ERROR]`r`n")
                if ($error.Site) { $detailsBox.AppendText("Site: $($error.Site)`r`n") }
                $detailsBox.AppendText("Message: $($error.Error)`r`n")
            }
        }

        # Add to session statistics
        Add-SessionOperation -Operation $Mode -SitesProcessed $finalResults.Sites.Count -VersionsDeleted $finalResults.TotalVersionsDeleted -StorageSaved $finalResults.TotalStorageSaved

        $statusLabel.Text = "Complete. Close window to continue."
        $progressBar.Style = "Continuous"
        $progressBar.Value = $progressBar.Maximum
        $cancelButton.Enabled = $true
        $cancelButton.Text = "Close"

        Write-Log "Processing complete: $($finalResults.Sites.Count) sites processed" "INFO"

    } catch {
        Write-Log "Error in Start-ProcessingWithProgress: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Processing Error", "OK", "Error")
    } finally {
        # Wait for form to close and keep message loop alive
        while ($progressForm.Visible) {
            [System.Windows.Forms.Application]::DoEvents()
            Start-Sleep -Milliseconds 100
        }
        
        # Dispose form resources
        if ($progressForm) {
            $progressForm.Dispose()
        }
    }
}

<#
.FUNCTION Start-MainThreadProcessing
.SYNOPSIS
    Handles UI updates from background job output on main thread
.DESCRIPTION
    Processes structured job output and safely updates UI controls.
    Ensures all UI modifications happen on main thread (checked via InvokeRequired).
.PARAMETER Control
    UI control to update
.PARAMETER JobOutput
    Output from background job
#>
function Start-MainThreadProcessing {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("SCAN", "DRYRUN", "CLEANUP")]
        [string]$Mode,

        [Parameter(Mandatory = $true)]
        [ValidateSet("AllSites", "SingleSite", "SelectedSites")]
        [string]$Scope,

        [Parameter()]
        [string]$SiteUrl,

        [Parameter()]
        [string]$LibraryName,

        [Parameter()]
        [object[]]$SelectedSites = @(),

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Form]$ProgressForm,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Label]$StatusLabel,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.TextBox]$DetailsBox,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.ProgressBar]$ProgressBar,

        [Parameter(Mandatory = $true)]
        [System.Windows.Forms.Button]$CancelButton
    )

    $results = @{
        Mode = $Mode
        Scope = $Scope
        Sites = @()
        TotalVersionsDeleted = 0
        TotalStorageSaved = 0
        Errors = @()
    }

    try {
        if ($Scope -eq "AllSites") {
            Write-Log "Starting All Sites operation on main thread" "INFO"
            $StatusLabel.Text = "Retrieving all sites..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Get all active sites with error handling
            try {
                $sites = Get-PnPTenantSite -Detailed | Where-Object { $_.Status -eq "Active" }
            } catch {
                $errorMessage = $_.Exception.Message
                Write-Log "Error retrieving tenant sites: $errorMessage" "ERROR"
                
                # Check if it's an authorization error
                if ($errorMessage -match "Unauthorized|unauthorized|permission|Access Denied") {
                    $DetailsBox.AppendText("⚠️  PERMISSION ERROR: Cannot list all sites`r`n`r`n")
                    $DetailsBox.AppendText("Your account or Entra app lacks permission to retrieve all tenant sites.`r`n`r`n")
                    $DetailsBox.AppendText("SOLUTIONS:`r`n")
                    $DetailsBox.AppendText("1. Use 'Scan Single Site' mode instead`r`n")
                    $DetailsBox.AppendText("2. Ask your SharePoint Administrator to grant 'Sites.Read.All' permission`r`n")
                    $DetailsBox.AppendText("3. Ensure your Entra app has Application-level permissions (not delegated)`r`n`r`n")
                    $DetailsBox.AppendText("Technical details:`r`n$errorMessage`r`n")
                    $DetailsBox.ScrollToCaret()
                    
                    $results.Errors += @{
                        General = "Authorization failed. Use Single Site mode instead."
                    }
                    
                    Write-Log "Falling back to single-site guidance" "INFO"
                } else {
                    # Other errors
                    $DetailsBox.AppendText("⚠️  ERROR: Failed to retrieve sites`r`n`r`n")
                    $DetailsBox.AppendText("Error: $errorMessage`r`n")
                    $DetailsBox.ScrollToCaret()
                    
                    $results.Errors += @{
                        General = $errorMessage
                    }
                }
                
                $StatusLabel.Text = "Error - Use Single Site mode"
                $CancelButton.Text = "Done"
                $CancelButton.Enabled = $true
                return $results
            }
            
            if (-not $sites) {
                $DetailsBox.AppendText("No active sites found.`r`n")
                return $results
            }

            $totalSites = @($sites).Count
            $DetailsBox.AppendText("Found $totalSites sites to process`r`n`r`n")
            $ProgressBar.Maximum = $totalSites
            $ProgressBar.Value = 0
            $ProgressBar.Style = "Continuous"

            $siteIndex = 0
            foreach ($site in $sites) {
                if ($global:cancelRequested) {
                    $DetailsBox.AppendText("Operation cancelled by user`r`n")
                    break
                }

                $siteIndex++
                $StatusLabel.Text = "Processing: $($site.Title) ($siteIndex/$totalSites)"
                $ProgressBar.Value = $siteIndex
                [System.Windows.Forms.Application]::DoEvents()

                try {
                    $DetailsBox.AppendText("`r`nProcessing site: $($site.Title)`r`n")
                    $DetailsBox.AppendText("  URL: $($site.Url)`r`n")
                    [System.Windows.Forms.Application]::DoEvents()

                    $siteResult = Invoke-SiteVersionCleanup -SiteUrl $site.Url -Mode $Mode -DetailsBox $DetailsBox
                    $results.Sites += $siteResult
                    $results.TotalVersionsDeleted += $siteResult.VersionsDeleted
                    $results.TotalStorageSaved += $siteResult.StorageSaved

                    $DetailsBox.AppendText("  ✓ Versions deleted: $($siteResult.VersionsDeleted), Storage saved: $(Format-FileSize -Size $siteResult.StorageSaved)`r`n")
                } catch {
                    $results.Errors += @{
                        Site = $site.Url
                        Error = $_.Exception.Message
                    }
                    $DetailsBox.AppendText("  ✗ ERROR: $($_.Exception.Message)`r`n")
                    Write-Log "Error processing site $($site.Url): $($_.Exception.Message)" "ERROR"
                }
            }
        } elseif ($Scope -eq "SelectedSites") {
            Write-Log "Starting Selected Sites operation on main thread" "INFO"
            
            if ($SelectedSites.Count -eq 0) {
                $DetailsBox.AppendText("No sites selected for processing.`r`n")
                $results.Errors += @{ General = "No sites selected" }
                $StatusLabel.Text = "No sites selected"
                $CancelButton.Text = "Done"
                $CancelButton.Enabled = $true
                return $results
            }

            $totalSites = @($SelectedSites).Count
            $DetailsBox.AppendText("Processing $totalSites selected site(s)`r`n`r`n")
            $ProgressBar.Maximum = $totalSites
            $ProgressBar.Value = 0
            $ProgressBar.Style = "Continuous"

            $siteIndex = 0
            foreach ($site in $SelectedSites) {
                if ($global:cancelRequested) {
                    $DetailsBox.AppendText("Operation cancelled by user`r`n")
                    break
                }

                $siteIndex++
                $StatusLabel.Text = "Processing: $($site.Title) ($siteIndex/$totalSites)"
                $ProgressBar.Value = $siteIndex
                [System.Windows.Forms.Application]::DoEvents()

                try {
                    $DetailsBox.AppendText("`r`nProcessing site: $($site.Title)`r`n")
                    $DetailsBox.AppendText("  URL: $($site.Url)`r`n")
                    [System.Windows.Forms.Application]::DoEvents()

                    $siteResult = Invoke-SiteVersionCleanup -SiteUrl $site.Url -Mode $Mode -DetailsBox $DetailsBox
                    $results.Sites += $siteResult
                    $results.TotalVersionsDeleted += $siteResult.VersionsDeleted
                    $results.TotalStorageSaved += $siteResult.StorageSaved

                    $DetailsBox.AppendText("  ✓ Versions deleted: $($siteResult.VersionsDeleted), Storage saved: $(Format-FileSize -Size $siteResult.StorageSaved)`r`n")
                } catch {
                    $results.Errors += @{
                        Site = $site.Url
                        Error = $_.Exception.Message
                    }
                    $DetailsBox.AppendText("  ✗ ERROR: $($_.Exception.Message)`r`n")
                    Write-Log "Error processing site $($site.Url): $($_.Exception.Message)" "ERROR"
                }
            }
        } else {
            Write-Log "Starting Single Site operation on main thread" "INFO"
            $StatusLabel.Text = "Processing single site..."
            [System.Windows.Forms.Application]::DoEvents()

            $siteResult = Invoke-SiteVersionCleanup -SiteUrl $SiteUrl -LibraryName $LibraryName -Mode $Mode
            $results.Sites += $siteResult
            $results.TotalVersionsDeleted = $siteResult.VersionsDeleted
            $results.TotalStorageSaved = $siteResult.StorageSaved

            $DetailsBox.AppendText("`r`n✓ Site processed successfully`r`n")
            $DetailsBox.AppendText("  Versions deleted: $($siteResult.VersionsDeleted)`r`n")
            $DetailsBox.AppendText("  Storage saved: $(Format-FileSize -Size $siteResult.StorageSaved)`r`n")
        }
    } catch {
        $results.Errors += @{
            General = $_.Exception.Message
        }
        Write-Log "Error in main thread processing: $($_.Exception.Message)" "ERROR"
        $DetailsBox.AppendText("`r`n✗ CRITICAL ERROR: $($_.Exception.Message)`r`n")
    } finally {
        $DetailsBox.AppendText("`r`n--- OPERATION COMPLETE ---`r`n")
        $DetailsBox.AppendText("Total sites processed: $($results.Sites.Count)`r`n")
        
        # Calculate totals for files scanned and with excess versions
        $totalFilesScanned = ($results.Sites | Measure-Object -Property TotalFilesScanned -Sum).Sum
        $totalFilesWithExcess = ($results.Sites | Measure-Object -Property FilesWithExcess -Sum).Sum
        
        $DetailsBox.AppendText("Total files scanned: $totalFilesScanned`r`n")
        $DetailsBox.AppendText("Files with excess versions: $totalFilesWithExcess`r`n")
        $DetailsBox.AppendText("Total versions deleted: $($results.TotalVersionsDeleted)`r`n")
        $DetailsBox.AppendText("Total storage saved: $(Format-FileSize -Size $results.TotalStorageSaved)`r`n")

        if ($results.Errors.Count -gt 0) {
            $DetailsBox.AppendText("`r`n⚠️  Errors encountered: $($results.Errors.Count)`r`n")
        }

        $StatusLabel.Text = "Complete. Click Done to continue."
        $ProgressBar.Style = "Continuous"
        $ProgressBar.Value = $ProgressBar.Maximum
        $CancelButton.Enabled = $true
        $CancelButton.Text = "Done"
        [System.Windows.Forms.Application]::DoEvents()
    }

    return $results
}

# Private helper function
function Invoke-SiteVersionCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SiteUrl,

        [Parameter()]
        [string]$LibraryName,

        [Parameter(Mandatory = $true)]
        [ValidateSet("SCAN", "DRYRUN", "CLEANUP")]
        [string]$Mode,

        [Parameter()]
        [System.Windows.Forms.TextBox]$DetailsBox = $null
    )

    $siteResult = @{
        SiteUrl = $SiteUrl
        VersionsDeleted = 0
        StorageSaved = 0
        LibrariesProcessed = 0
        FilesProcessed = 0
        TotalFilesScanned = 0
        FilesWithExcess = 0
    }

    try {
        Connect-SharePoint -Url $SiteUrl

        if ($LibraryName) {
            # Single library
            $libraries = Get-PnPList -Identity $LibraryName | Where-Object {$_.BaseTemplate -eq 101 -or $_.BaseTemplate -eq 109}
        } else {
            # All document libraries
            $libraries = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101 -or $_.BaseTemplate -eq 109}
        }

        foreach ($library in $libraries) {
            $libraryOutput = "  Library: $($library.Title)"
            Write-Host $libraryOutput
            if ($DetailsBox) {
                $DetailsBox.AppendText("$libraryOutput`r`n")
                [System.Windows.Forms.Application]::DoEvents()
            }
            
            $items = Get-PnPListItem -List $library.Id -PageSize 5000

            foreach ($item in $items) {
                # Check if item is a file (not a folder)
                $isFile = ![string]::IsNullOrEmpty($item.FieldValues.FileLeafRef)
                
                if ($isFile) {
                    $versions = Get-FileVersions -Item $item
                    
                    # For SCAN and DRYRUN: Show ALL files (including those with no excess versions)
                    # For CLEANUP: Only show files with excess versions that will be deleted
                    if ($Mode -in @("SCAN", "DRYRUN")) {
                        # Count ALL files scanned (even those with no version history)
                        $siteResult.TotalFilesScanned++
                        
                        # Show all files with version counts
                        if ($versions.Count -gt 0) {
                            $excess = [Math]::Max(0, $versions.Count - $global:keepVersions)
                            if ($excess -gt 0) {
                                $siteResult.FilesWithExcess++
                                $fileOutput = "    File: $($item.FieldValues.FileLeafRef) - $($versions.Count) versions (excess: $excess)"
                                Write-Host $fileOutput
                                if ($DetailsBox) {
                                    $DetailsBox.AppendText("$fileOutput`r`n")
                                    [System.Windows.Forms.Application]::DoEvents()
                                }
                                $siteResult.FilesProcessed++
                            } else {
                                $fileOutput = "    File: $($item.FieldValues.FileLeafRef) - $($versions.Count) versions (no excess)"
                                Write-Host $fileOutput
                                if ($DetailsBox) {
                                    $DetailsBox.AppendText("$fileOutput`r`n")
                                    [System.Windows.Forms.Application]::DoEvents()
                                }
                            }
                            
                            # For DRYRUN: Show what would be deleted
                            if ($Mode -eq "DRYRUN" -and $excess -gt 0) {
                                $deleted = Remove-FileVersions -Item $item -VersionsToKeep $global:keepVersions -Mode $Mode
                                $siteResult.VersionsDeleted += $deleted.Count
                                $siteResult.StorageSaved += $deleted.StorageSaved
                            }
                        } else {
                            # File with no version history
                            $fileOutput = "    File: $($item.FieldValues.FileLeafRef) - (no version history)"
                            Write-Host $fileOutput
                            if ($DetailsBox) {
                                $DetailsBox.AppendText("$fileOutput`r`n")
                                [System.Windows.Forms.Application]::DoEvents()
                            }
                        }
                    } elseif ($Mode -eq "CLEANUP" -and $versions.Count -gt $global:keepVersions) {
                        # CLEANUP: Only process files with excess versions
                        $excess = $versions.Count - $global:keepVersions
                        $fileOutput = "    File: $($item.FieldValues.FileLeafRef) - $excess excess versions to DELETE"
                        Write-Host $fileOutput
                        if ($DetailsBox) {
                            $DetailsBox.AppendText("$fileOutput`r`n")
                            [System.Windows.Forms.Application]::DoEvents()
                        }

                        $deleted = Remove-FileVersions -Item $item -VersionsToKeep $global:keepVersions -Mode $Mode
                        $siteResult.VersionsDeleted += $deleted.Count
                        $siteResult.StorageSaved += $deleted.StorageSaved
                        $siteResult.FilesProcessed++
                    }
                }
            }

            $siteResult.LibrariesProcessed++
        }
    } catch {
        Write-Host "Error processing site $SiteUrl : $($_.Exception.Message)"
    }

    return $siteResult
}

# Export public functions
Export-ModuleMember -Function @(
    'Start-ProcessingWithProgress',
    'Start-MainThreadProcessing',
    'Invoke-SiteVersionCleanup'
)

