<# 
.SYNOPSIS
    UI module for SharePoint Version Cleaner
    Contains all Windows Forms dialog functions
#>

# Note: Large UI functions organized below

function Show-SiteSelectorDialog {
    $siteForm = New-Object System.Windows.Forms.Form
    $siteForm.Text = "SharePoint Site Selector"
    $siteForm.Size = New-Object System.Drawing.Size(900, 700)
    $siteForm.StartPosition = "CenterScreen"
    $siteForm.FormBorderStyle = "FixedDialog"
    $siteForm.MaximizeBox = $false
    $siteForm.MinimizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(860, 30)
    $titleLabel.Text = "Select SharePoint Site"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $siteForm.Controls.Add($titleLabel)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 60)
    $statusLabel.Size = New-Object System.Drawing.Size(860, 30)
    $statusLabel.Text = "Ready to load sites..."
    $siteForm.Controls.Add($statusLabel)

    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Location = New-Object System.Drawing.Point(20, 100)
    $searchLabel.Size = New-Object System.Drawing.Size(100, 20)
    $searchLabel.Text = "Search sites:"
    $siteForm.Controls.Add($searchLabel)

    $searchTextBox = New-Object System.Windows.Forms.TextBox
    $searchTextBox.Location = New-Object System.Drawing.Point(130, 98)
    $searchTextBox.Size = New-Object System.Drawing.Size(500, 20)
    $searchTextBox.Enabled = $false
    $siteForm.Controls.Add($searchTextBox)

    $loadButton = New-Object System.Windows.Forms.Button
    $loadButton.Location = New-Object System.Drawing.Point(640, 97)
    $loadButton.Size = New-Object System.Drawing.Size(80, 22)
    $loadButton.Text = "Load Sites"
    $loadButton.BackColor = [System.Drawing.Color]::LightGreen
    $siteForm.Controls.Add($loadButton)

    $sitesListView = New-Object System.Windows.Forms.ListView
    $sitesListView.Location = New-Object System.Drawing.Point(20, 130)
    $sitesListView.Size = New-Object System.Drawing.Size(860, 480)
    $sitesListView.View = [System.Windows.Forms.View]::Details
    $sitesListView.FullRowSelect = $true
    $sitesListView.GridLines = $true
    $sitesListView.MultiSelect = $false

    $sitesListView.Columns.Add("Title", 300) | Out-Null
    $sitesListView.Columns.Add("URL", 450) | Out-Null
    $sitesListView.Columns.Add("Template", 100) | Out-Null

    $siteForm.Controls.Add($sitesListView)

    $selectButton = New-Object System.Windows.Forms.Button
    $selectButton.Location = New-Object System.Drawing.Point(650, 620)
    $selectButton.Size = New-Object System.Drawing.Size(80, 30)
    $selectButton.Text = "Select"
    $selectButton.Enabled = $false
    $selectButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $siteForm.Controls.Add($selectButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(740, 620)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $siteForm.Controls.Add($cancelButton)

    $script:allSites = @()
    $script:selectedSiteUrl = $null

    $sitesListView.Add_SelectedIndexChanged({
        $selectButton.Enabled = ($sitesListView.SelectedItems.Count -gt 0)
    })

    $sitesListView.Add_DoubleClick({
        if ($sitesListView.SelectedItems.Count -gt 0) {
            $script:selectedSiteUrl = $sitesListView.SelectedItems[0].SubItems[1].Text
            $siteForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $siteForm.Close()
        }
    })

    $selectButton.Add_Click({
        if ($sitesListView.SelectedItems.Count -gt 0) {
            $script:selectedSiteUrl = $sitesListView.SelectedItems[0].SubItems[1].Text
            $siteForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $siteForm.Close()
        }
    })

    $searchTextBox.Add_TextChanged({
        if ($searchTextBox.Enabled -and $script:allSites.Count -gt 0) {
            $searchText = $searchTextBox.Text.ToLower()
            $sitesListView.Items.Clear()
            
            # Batch UI updates
            $sitesListView.BeginUpdate()
            
            $matchCount = 0
            foreach ($site in $script:allSites) {
                $siteTitle = if ($site.Title) { $site.Title } else { "Unknown" }
                $siteUrl = if ($site.Url) { $site.Url } else { "" }
                
                if ([string]::IsNullOrWhiteSpace($searchText) -or
                    $siteTitle.ToLower().Contains($searchText) -or
                    $siteUrl.ToLower().Contains($searchText)) {
                    
                    $item = New-Object System.Windows.Forms.ListViewItem($siteTitle)
                    $item.SubItems.Add($siteUrl)
                    $item.SubItems.Add($site.Template)
                    $sitesListView.Items.Add($item)
                    $matchCount++
                }
            }
            
            $sitesListView.EndUpdate()
        }
    })

    $loadButton.Add_Click({
        try {
            $statusLabel.Text = "Loading sites..."
            $loadButton.Enabled = $false
            $loadButton.Text = "Loading..."
            $searchTextBox.Enabled = $false
            $selectButton.Enabled = $false
            $sitesListView.Items.Clear()
            $script:allSites = @()
            
            [System.Windows.Forms.Application]::DoEvents()

            # Check if global cache exists and is fresh (from main script)
            if ($global:cachedAllSites -and $global:cachedAllSites.Count -gt 0) {
                $statusLabel.Text = "Using cached site list..."
                [System.Windows.Forms.Application]::DoEvents()
                $sites = $global:cachedAllSites
                Write-Log "Site selector using global cache: $($sites.Count) sites" "INFO"
            } else {
                Connect-SharePoint -Url $script:adminUrl -UseCache -SuppressSuccessMessage
                $statusLabel.Text = "Fetching site list from SharePoint..."
                [System.Windows.Forms.Application]::DoEvents()
                
                $sites = Get-PnPTenantSite | Where-Object { $_.Template -ne "SPSPERS" } | Sort-Object Title
                Write-Log "Site selector fetched fresh list: $($sites.Count) sites" "INFO"
            }
                
            $statusLabel.Text = "Processing $($sites.Count) sites..."
            [System.Windows.Forms.Application]::DoEvents()
            
            # Use BeginUpdate/EndUpdate for batch UI updates
            $sitesListView.BeginUpdate()
            
            $processedCount = 0
            foreach ($site in $sites) {
                $siteObj = @{
                    Title = if ($site.Title) { $site.Title } else { "Unknown Site" }
                    Url = if ($site.Url) { $site.Url } else { "" }
                    Template = if ($site.Template) { $site.Template } else { "Unknown" }
                }
                
                $script:allSites += $siteObj
                
                $item = New-Object System.Windows.Forms.ListViewItem($siteObj.Title)
                $item.SubItems.Add($siteObj.Url)
                $item.SubItems.Add($siteObj.Template)
                $sitesListView.Items.Add($item)
                
                $processedCount++
                # Show progress every 200 items without blocking UI
                if (($processedCount % 200) -eq 0) {
                    $statusLabel.Text = "Processing $processedCount of $($sites.Count) sites..."
                    [System.Windows.Forms.Application]::DoEvents()
                }
            }
            
            # End batch update - applies all changes at once
            $sitesListView.EndUpdate()
                
        } catch {
            $statusLabel.Text = "Connection error"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to connect to SharePoint: $($_.Exception.Message)",
                "Connection Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        } finally {
            $loadButton.Enabled = $true
            $loadButton.Text = "Load Sites"
        }
    })

    $siteForm.AcceptButton = $selectButton
    $siteForm.CancelButton = $cancelButton

    $result = $siteForm.ShowDialog()
    $returnUrl = if ($result -eq [System.Windows.Forms.DialogResult]::OK) { $script:selectedSiteUrl } else { $null }
    
    $siteForm.Dispose()
    return $returnUrl
}

function Show-InitialSetupDialog {
    $setupForm = New-Object System.Windows.Forms.Form
    $setupForm.Text = "SPOVC - Initial Setup"
    $setupForm.Size = New-Object System.Drawing.Size(600, 350)
    $setupForm.StartPosition = "CenterScreen"
    $setupForm.FormBorderStyle = "FixedDialog"
    $setupForm.MaximizeBox = $false
    $setupForm.MinimizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(560, 30)
    $titleLabel.Text = "Configure SharePoint Version Cleaner"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $setupForm.Controls.Add($titleLabel)

    # Client ID
    $clientIdLabel = New-Object System.Windows.Forms.Label
    $clientIdLabel.Location = New-Object System.Drawing.Point(20, 70)
    $clientIdLabel.Size = New-Object System.Drawing.Size(150, 20)
    $clientIdLabel.Text = "Entra App ID:"
    $setupForm.Controls.Add($clientIdLabel)

    $clientIdBox = New-Object System.Windows.Forms.TextBox
    $clientIdBox.Location = New-Object System.Drawing.Point(180, 68)
    $clientIdBox.Size = New-Object System.Drawing.Size(380, 20)
    $clientIdBox.Text = $script:clientId
    $setupForm.Controls.Add($clientIdBox)

    # Admin URL
    $adminUrlLabel = New-Object System.Windows.Forms.Label
    $adminUrlLabel.Location = New-Object System.Drawing.Point(20, 110)
    $adminUrlLabel.Size = New-Object System.Drawing.Size(150, 20)
    $adminUrlLabel.Text = "Admin URL:"
    $setupForm.Controls.Add($adminUrlLabel)

    $adminUrlBox = New-Object System.Windows.Forms.TextBox
    $adminUrlBox.Location = New-Object System.Drawing.Point(180, 108)
    $adminUrlBox.Size = New-Object System.Drawing.Size(380, 20)
    $adminUrlBox.Text = $script:adminUrl
    $setupForm.Controls.Add($adminUrlBox)

    # Keep Versions
    $keepVersionsLabel = New-Object System.Windows.Forms.Label
    $keepVersionsLabel.Location = New-Object System.Drawing.Point(20, 150)
    $keepVersionsLabel.Size = New-Object System.Drawing.Size(150, 20)
    $keepVersionsLabel.Text = "Keep Versions:"
    $setupForm.Controls.Add($keepVersionsLabel)

    $keepVersionsSpinner = New-Object System.Windows.Forms.NumericUpDown
    $keepVersionsSpinner.Location = New-Object System.Drawing.Point(180, 148)
    $keepVersionsSpinner.Size = New-Object System.Drawing.Size(100, 20)
    $keepVersionsSpinner.Minimum = 1
    $keepVersionsSpinner.Maximum = 500
    $keepVersionsSpinner.Value = $script:keepVersions
    $setupForm.Controls.Add($keepVersionsSpinner)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 200)
    $infoLabel.Size = New-Object System.Drawing.Size(560, 60)
    $infoLabel.Text = "These settings are stored in config.json. You can also edit the file directly." + "`r`n`r`n" + "For Entra app setup, see the README.md file."
    $infoLabel.AutoSize = $false
    $infoLabel.ForeColor = [System.Drawing.Color]::Gray
    $setupForm.Controls.Add($infoLabel)

    $saveButton = New-Object System.Windows.Forms.Button
    $saveButton.Location = New-Object System.Drawing.Point(400, 290)
    $saveButton.Size = New-Object System.Drawing.Size(80, 30)
    $saveButton.Text = "Save"
    $saveButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $setupForm.Controls.Add($saveButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(490, 290)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $setupForm.Controls.Add($cancelButton)

    $result = $setupForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $config = @{
            clientId = $clientIdBox.Text
            adminUrl = $adminUrlBox.Text
            keepVersions = [int]$keepVersionsSpinner.Value
        }
        $setupForm.Dispose()
        return $config
    }
    
    $setupForm.Dispose()
    return $null
}

function Show-MainMenu {
    $script:selectedMode = $null
    
    $menuForm = New-Object System.Windows.Forms.Form
    $menuForm.Text = "SharePoint Version Cleaner - Main Menu"
    $menuForm.Size = New-Object System.Drawing.Size(500, 570)
    $menuForm.StartPosition = "CenterScreen"
    $menuForm.FormBorderStyle = "FixedDialog"
    $menuForm.MaximizeBox = $false
    $menuForm.MinimizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 45)
    $titleLabel.Text = "SharePoint Version Cleaner - Select an Operation"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $titleLabel.AutoSize = $false
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::TopLeft
    $menuForm.Controls.Add($titleLabel)

    # Status label showing selected sites
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 62)
    $statusLabel.Size = New-Object System.Drawing.Size(460, 15)
    if ($global:selectedSitesForProcessing -and $global:selectedSitesForProcessing.Count -gt 0) {
        $siteNames = $global:selectedSitesForProcessing | ForEach-Object { $_.Title } | Join-String -Separator ", "
        $statusLabel.Text = "Selected: $siteNames"
    } else {
        $statusLabel.Text = "No sites selected"
    }
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Italic)
    $statusLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $menuForm.Controls.Add($statusLabel)

    $selectSitesButton = New-Object System.Windows.Forms.Button
    $selectSitesButton.Location = New-Object System.Drawing.Point(50, 80)
    $selectSitesButton.Size = New-Object System.Drawing.Size(400, 35)
    $selectSitesButton.Text = "[T] Select Sites to Process (Required)"
    $selectSitesButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $selectSitesButton.BackColor = [System.Drawing.Color]::LightGreen
    $menuForm.Controls.Add($selectSitesButton)

    $scanButton = New-Object System.Windows.Forms.Button
    $scanButton.Location = New-Object System.Drawing.Point(50, 115)
    $scanButton.Size = New-Object System.Drawing.Size(400, 50)
    $scanButton.Text = "[S] Scan - Analyze all sites and show storage saved"
    $scanButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $scanButton.BackColor = [System.Drawing.Color]::LightBlue
    $menuForm.Controls.Add($scanButton)

    $dryrunButton = New-Object System.Windows.Forms.Button
    $dryrunButton.Location = New-Object System.Drawing.Point(50, 180)
    $dryrunButton.Size = New-Object System.Drawing.Size(400, 50)
    $dryrunButton.Text = "[!] Dry Run - Simulate deletion without making changes"
    $dryrunButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $dryrunButton.BackColor = [System.Drawing.Color]::LightYellow
    $menuForm.Controls.Add($dryrunButton)

    $cleanupButton = New-Object System.Windows.Forms.Button
    $cleanupButton.Location = New-Object System.Drawing.Point(50, 245)
    $cleanupButton.Size = New-Object System.Drawing.Size(400, 50)
    $cleanupButton.Text = "[X] Cleanup - Permanently delete excess versions (DOUBLE CONFIRMATION)"
    $cleanupButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $cleanupButton.BackColor = [System.Drawing.Color]::LightCoral
    $menuForm.Controls.Add($cleanupButton)

    $reportButton = New-Object System.Windows.Forms.Button
    $reportButton.Location = New-Object System.Drawing.Point(50, 310)
    $reportButton.Size = New-Object System.Drawing.Size(400, 35)
    $reportButton.Text = "[R] View Last Session Report"
    $reportButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $menuForm.Controls.Add($reportButton)

    $recycleBinButton = New-Object System.Windows.Forms.Button
    $recycleBinButton.Location = New-Object System.Drawing.Point(50, 350)
    $recycleBinButton.Size = New-Object System.Drawing.Size(400, 35)
    $recycleBinButton.Text = "[B] Recycle Bin Browser - Search and restore deleted files"
    $recycleBinButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $recycleBinButton.BackColor = [System.Drawing.Color]::LightCyan
    $menuForm.Controls.Add($recycleBinButton)

    $disconnectButton = New-Object System.Windows.Forms.Button
    $disconnectButton.Location = New-Object System.Drawing.Point(50, 395)
    $disconnectButton.Size = New-Object System.Drawing.Size(400, 35)
    $disconnectButton.Text = "[-] Disconnect - Clear cached authentication tokens"
    $disconnectButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $disconnectButton.BackColor = [System.Drawing.Color]::LightGray
    $menuForm.Controls.Add($disconnectButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Location = New-Object System.Drawing.Point(50, 440)
    $exitButton.Size = New-Object System.Drawing.Size(400, 35)
    $exitButton.Text = "[Q] Exit"
    $exitButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $menuForm.Controls.Add($exitButton)

    $selectSitesButton.Add_Click({
        $script:selectedMode = "SELECT_SITES"
        $menuForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $menuForm.Close()
    })

    $scanButton.Add_Click({
        $script:selectedMode = "SCAN"
        $menuForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $menuForm.Close()
    })

    $dryrunButton.Add_Click({
        $script:selectedMode = "DRYRUN"
        $menuForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $menuForm.Close()
    })

    $cleanupButton.Add_Click({
        $script:selectedMode = "CLEANUP"
        $menuForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $menuForm.Close()
    })

    $reportButton.Add_Click({
        # Generate a new session report from current session statistics
        Generate-SessionReport
    })

    $recycleBinButton.Add_Click({
        Show-RecycleBinBrowserDialog
    })

    $disconnectButton.Add_Click({
        $script:selectedMode = "DISCONNECT"
        $menuForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $menuForm.Close()
    })

    $exitButton.Add_Click({
        $script:selectedMode = "EXIT"
        $menuForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $menuForm.Close()
    })

    $menuForm.AcceptButton = $scanButton
    $menuForm.CancelButton = $exitButton

    $menuForm.ShowDialog() | Out-Null
    $menuForm.Dispose()
    return $script:selectedMode
}

function Show-ScanOptions {
    $optionsForm = New-Object System.Windows.Forms.Form
    $optionsForm.Text = "SPOVC - Scan Options"
    $optionsForm.Size = New-Object System.Drawing.Size(500, 320)
    $optionsForm.StartPosition = "CenterScreen"
    $optionsForm.FormBorderStyle = "FixedDialog"
    $optionsForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Text = "Select scan scope:"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $optionsForm.Controls.Add($titleLabel)

    $allSitesRadio = New-Object System.Windows.Forms.RadioButton
    $allSitesRadio.Location = New-Object System.Drawing.Point(40, 70)
    $allSitesRadio.Size = New-Object System.Drawing.Size(420, 50)
    $allSitesRadio.Text = "Scan All Sites - Analyze versions across all SharePoint sites"
    $allSitesRadio.AutoSize = $false
    $allSitesRadio.Checked = $true
    $optionsForm.Controls.Add($allSitesRadio)

    $multiRadio = New-Object System.Windows.Forms.RadioButton
    $multiRadio.Location = New-Object System.Drawing.Point(40, 125)
    $multiRadio.Size = New-Object System.Drawing.Size(420, 50)
    $multiRadio.Text = "Select Multiple Sites - Choose specific sites to analyze"
    $multiRadio.AutoSize = $false
    $optionsForm.Controls.Add($multiRadio)

    $singleRadio = New-Object System.Windows.Forms.RadioButton
    $singleRadio.Location = New-Object System.Drawing.Point(40, 180)
    $singleRadio.Size = New-Object System.Drawing.Size(420, 50)
    $singleRadio.Text = "Scan Single Site/Library - Analyze versions in specific library"
    $singleRadio.AutoSize = $false
    $optionsForm.Controls.Add($singleRadio)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(320, 250)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Text = "Next"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $optionsForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(410, 250)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $optionsForm.Controls.Add($cancelButton)

    $optionsForm.AcceptButton = $okButton
    $optionsForm.CancelButton = $cancelButton

    $result = $optionsForm.ShowDialog()
    $scope = if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($singleRadio.Checked) { 
            "SingleSite" 
        } elseif ($multiRadio.Checked) {
            "SelectedSites"
        } else { 
            "AllSites" 
        }
    } else {
        $null
    }

    $optionsForm.Dispose()
    return $scope
}

function Show-CleanupOptions {
    $optionsForm = New-Object System.Windows.Forms.Form
    $optionsForm.Text = "SPOVC - Cleanup Options"
    $optionsForm.Size = New-Object System.Drawing.Size(500, 300)
    $optionsForm.StartPosition = "CenterScreen"
    $optionsForm.FormBorderStyle = "FixedDialog"
    $optionsForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Text = "Cleanup scope and intensity:"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $optionsForm.Controls.Add($titleLabel)

    $allDryrunRadio = New-Object System.Windows.Forms.RadioButton
    $allDryrunRadio.Location = New-Object System.Drawing.Point(40, 70)
    $allDryrunRadio.Size = New-Object System.Drawing.Size(420, 50)
    $allDryrunRadio.Text = "Dry-Run All Sites - Simulate cleanup without deleting"
    $allDryrunRadio.AutoSize = $false
    $optionsForm.Controls.Add($allDryrunRadio)

    $allCleanupRadio = New-Object System.Windows.Forms.RadioButton
    $allCleanupRadio.Location = New-Object System.Drawing.Point(40, 125)
    $allCleanupRadio.Size = New-Object System.Drawing.Size(420, 50)
    $allCleanupRadio.Text = "Cleanup All Sites (WARNING: PERMANENT) - Delete excess versions across all sites"
    $allCleanupRadio.AutoSize = $false
    $allCleanupRadio.Checked = $true
    $optionsForm.Controls.Add($allCleanupRadio)

    $singleRadio = New-Object System.Windows.Forms.RadioButton
    $singleRadio.Location = New-Object System.Drawing.Point(40, 180)
    $singleRadio.Size = New-Object System.Drawing.Size(420, 50)
    $singleRadio.Text = "Cleanup Single Site - Delete excess versions in one site only"
    $singleRadio.AutoSize = $false
    $optionsForm.Controls.Add($singleRadio)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(320, 260)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Text = "Next"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $optionsForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(410, 260)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $optionsForm.Controls.Add($cancelButton)

    $optionsForm.AcceptButton = $okButton
    $optionsForm.CancelButton = $cancelButton

    $result = $optionsForm.ShowDialog()
    
    $option = $null
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ($allDryrunRadio.Checked) { $option = "DryRunAll" }
        elseif ($allCleanupRadio.Checked) { $option = "CleanupAll" }
        else { $option = "CleanupSingle" }
    }

    $optionsForm.Dispose()
    return $option
}

function Show-SingleSiteDialog {
    param(
        [Parameter()]
        [string]$DefaultSiteUrl = "https://tenant.sharepoint.com/sites/sitename"
    )

    $siteForm = New-Object System.Windows.Forms.Form
    $siteForm.Text = "SPOVC - Enter Site Details"
    $siteForm.Size = New-Object System.Drawing.Size(600, 250)
    $siteForm.StartPosition = "CenterScreen"
    $siteForm.FormBorderStyle = "FixedDialog"
    $siteForm.MaximizeBox = $false

    $siteUrlLabel = New-Object System.Windows.Forms.Label
    $siteUrlLabel.Location = New-Object System.Drawing.Point(20, 30)
    $siteUrlLabel.Size = New-Object System.Drawing.Size(150, 20)
    $siteUrlLabel.Text = "Site URL:"
    $siteForm.Controls.Add($siteUrlLabel)

    $siteUrlBox = New-Object System.Windows.Forms.TextBox
    $siteUrlBox.Location = New-Object System.Drawing.Point(180, 28)
    $siteUrlBox.Size = New-Object System.Drawing.Size(380, 20)
    $siteUrlBox.Text = $DefaultSiteUrl
    $siteForm.Controls.Add($siteUrlBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(490, 28)
    $browseButton.Size = New-Object System.Drawing.Size(70, 20)
    $browseButton.Text = "Browse"
    $siteForm.Controls.Add($browseButton)

    $browseButton.Add_Click({
        $selected = Show-SiteSelectorDialog
        if ($selected) {
            $siteUrlBox.Text = $selected
        }
    })

    $libraryLabel = New-Object System.Windows.Forms.Label
    $libraryLabel.Location = New-Object System.Drawing.Point(20, 70)
    $libraryLabel.Size = New-Object System.Drawing.Size(150, 20)
    $libraryLabel.Text = "Library Name (optional):"
    $siteForm.Controls.Add($libraryLabel)

    $libraryBox = New-Object System.Windows.Forms.TextBox
    $libraryBox.Location = New-Object System.Drawing.Point(180, 68)
    $libraryBox.Size = New-Object System.Drawing.Size(380, 20)
    $libraryBox.Text = "Documents"
    $siteForm.Controls.Add($libraryBox)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Location = New-Object System.Drawing.Point(20, 110)
    $infoLabel.Size = New-Object System.Drawing.Size(560, 70)
    $infoLabel.Text = "Enter the SharePoint site URL and optionally a specific library name.`r`nIf no library is specified, all document libraries will be processed.`r`n`r`nExample: https://contoso.sharepoint.com/sites/sales"
    $infoLabel.AutoSize = $false
    $infoLabel.ForeColor = [System.Drawing.Color]::Gray
    $siteForm.Controls.Add($infoLabel)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(420, 190)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $siteForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(510, 190)
    $cancelButton.Size = New-Object System.Drawing.Size(70, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $siteForm.Controls.Add($cancelButton)

    $siteForm.AcceptButton = $okButton
    $siteForm.CancelButton = $cancelButton

    $result = $siteForm.ShowDialog()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $siteForm.Dispose()
        return @{
            SiteUrl = $siteUrlBox.Text
            LibraryName = if ($libraryBox.Text) { $libraryBox.Text } else { $null }
        }
    }
    
    $siteForm.Dispose()
    return $null
}

function Show-InactiveSitesDialog {
    param([PSObject[]]$InactiveSites)

    $inactiveForm = New-Object System.Windows.Forms.Form
    $inactiveForm.Text = "SPOVC - Inactive Sites"
    $inactiveForm.Size = New-Object System.Drawing.Size(900, 600)
    $inactiveForm.StartPosition = "CenterScreen"

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(860, 30)
    $titleLabel.Text = "Found $($InactiveSites.Count) inactive sites. Select sites for version cleanup:"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 11, [System.Drawing.FontStyle]::Bold)
    $inactiveForm.Controls.Add($titleLabel)

    $sitesListView = New-Object System.Windows.Forms.ListView
    $sitesListView.Location = New-Object System.Drawing.Point(20, 60)
    $sitesListView.Size = New-Object System.Drawing.Size(860, 450)
    $sitesListView.View = [System.Windows.Forms.View]::Details
    $sitesListView.CheckBoxes = $true
    $sitesListView.GridLines = $true
    $sitesListView.MultiSelect = $true

    $sitesListView.Columns.Add("Title", 250) | Out-Null
    $sitesListView.Columns.Add("Last Modified", 120) | Out-Null
    $sitesListView.Columns.Add("Days Inactive", 80) | Out-Null
    $sitesListView.Columns.Add("Storage (MB)", 100) | Out-Null
    $sitesListView.Columns.Add("URL", 310) | Out-Null

    foreach ($site in $InactiveSites) {
        $item = New-Object System.Windows.Forms.ListViewItem
        $item.Text = $site.Title
        $item.SubItems.Add($site.LastModified.ToString("yyyy-MM-dd")) | Out-Null
        $item.SubItems.Add($site.DaysInactive) | Out-Null
        $item.SubItems.Add([math]::Round($site.StorageMB, 2)) | Out-Null
        $item.SubItems.Add($site.Url) | Out-Null
        $sitesListView.Items.Add($item) | Out-Null
    }

    $inactiveForm.Controls.Add($sitesListView)

    $cleanupButton = New-Object System.Windows.Forms.Button
    $cleanupButton.Location = New-Object System.Drawing.Point(660, 530)
    $cleanupButton.Size = New-Object System.Drawing.Size(100, 30)
    $cleanupButton.Text = "Cleanup Selected"
    $cleanupButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $inactiveForm.Controls.Add($cleanupButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(780, 530)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $inactiveForm.Controls.Add($cancelButton)

    $result = $inactiveForm.ShowDialog()
    
    $selectedSites = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($item in $sitesListView.CheckedItems) {
            $siteUrl = $item.SubItems[4].Text
            $selectedSites += $InactiveSites | Where-Object {$_.Url -eq $siteUrl}
        }
    }

    $inactiveForm.Dispose()
    return $selectedSites
}

# Export public functions

function Show-MultiSiteSelectorDialog {
    param(
        [Parameter()]
        [string]$DefaultSiteUrl = ""
    )

    $multiForm = New-Object System.Windows.Forms.Form
    $multiForm.Text = "SPOVC - Select SharePoint Sites"
    $multiForm.Size = New-Object System.Drawing.Size(700, 600)
    $multiForm.StartPosition = "CenterScreen"
    $multiForm.FormBorderStyle = "FixedDialog"
    $multiForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(660, 25)
    $titleLabel.Text = "Select Sites to Process (check multiple sites)"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $multiForm.Controls.Add($titleLabel)

    $loadingLabel = New-Object System.Windows.Forms.Label
    $loadingLabel.Location = New-Object System.Drawing.Point(20, 55)
    $loadingLabel.Size = New-Object System.Drawing.Size(660, 20)
    $loadingLabel.Text = "Loading sites..."
    $loadingLabel.ForeColor = [System.Drawing.Color]::Blue
    $multiForm.Controls.Add($loadingLabel)

    $sitesListView = New-Object System.Windows.Forms.ListView
    $sitesListView.Location = New-Object System.Drawing.Point(20, 85)
    $sitesListView.Size = New-Object System.Drawing.Size(660, 400)
    $sitesListView.View = [System.Windows.Forms.View]::Details
    $sitesListView.FullRowSelect = $true
    $sitesListView.GridLines = $true
    $sitesListView.CheckBoxes = $true
    $sitesListView.Columns.Add("Title", 200) | Out-Null
    $sitesListView.Columns.Add("URL", 350) | Out-Null
    $sitesListView.Columns.Add("Status", 80) | Out-Null
    $multiForm.Controls.Add($sitesListView)

    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Location = New-Object System.Drawing.Point(20, 500)
    $selectAllButton.Size = New-Object System.Drawing.Size(100, 25)
    $selectAllButton.Text = "Select All"
    $selectAllButton.Add_Click({
        foreach ($item in $sitesListView.Items) {
            $item.Checked = $true
        }
    })
    $multiForm.Controls.Add($selectAllButton)

    $selectNoneButton = New-Object System.Windows.Forms.Button
    $selectNoneButton.Location = New-Object System.Drawing.Point(130, 500)
    $selectNoneButton.Size = New-Object System.Drawing.Size(100, 25)
    $selectNoneButton.Text = "Select None"
    $selectNoneButton.Add_Click({
        foreach ($item in $sitesListView.Items) {
            $item.Checked = $false
        }
    })
    $multiForm.Controls.Add($selectNoneButton)

    $processButton = New-Object System.Windows.Forms.Button
    $processButton.Location = New-Object System.Drawing.Point(520, 500)
    $processButton.Size = New-Object System.Drawing.Size(80, 30)
    $processButton.Text = "Process"
    $processButton.Enabled = $false
    $processButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $multiForm.Controls.Add($processButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(610, 500)
    $cancelButton.Size = New-Object System.Drawing.Size(70, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $multiForm.Controls.Add($cancelButton)

    $multiForm.AcceptButton = $processButton
    $multiForm.CancelButton = $cancelButton

    # Enable Process button when any item is checked
    $sitesListView.Add_ItemChecked({
        $processButton.Enabled = ($sitesListView.CheckedItems.Count -gt 0)
    })

    $multiForm.Show()
    [System.Windows.Forms.Application]::DoEvents()

    # Load sites in background to avoid blocking UI
    $loadingSites = $true
    try {
        $loadingLabel.Text = "Retrieving SharePoint sites..."
        [System.Windows.Forms.Application]::DoEvents()

        # Try to get all sites
        try {
            $sites = Get-PnPTenantSite -Detailed | Where-Object { $_.Status -eq "Active" } | Sort-Object Title
            $loadingLabel.Text = "Found $($sites.Count) active sites"
        } catch {
            $loadingLabel.Text = "Error loading sites: $_"
            $loadingLabel.ForeColor = [System.Drawing.Color]::Red
            $sites = @()
        }

        # Populate list
        foreach ($site in @($sites)) {
            $item = New-Object System.Windows.Forms.ListViewItem($site.Title)
            $item.SubItems.Add($site.Url) | Out-Null
            $item.SubItems.Add($site.Status) | Out-Null
            $sitesListView.Items.Add($item) | Out-Null
        }

        if ($sites.Count -eq 0) {
            $loadingLabel.Text = "No active sites found. Enter site URL manually or check permissions."
            $loadingLabel.ForeColor = [System.Drawing.Color]::Orange
        } else {
            $loadingLabel.Text = "$($sites.Count) active sites available"
            $loadingLabel.ForeColor = [System.Drawing.Color]::Green
        }

        $loadingSites = $false
    } catch {
        $loadingLabel.Text = "Error: $_"
        $loadingLabel.ForeColor = [System.Drawing.Color]::Red
        $loadingSites = $false
    }

    # Show dialog and get result
    $result = $multiForm.ShowDialog()

    $selectedSites = @()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($item in $sitesListView.CheckedItems) {
            $selectedSites += @{
                Title = $item.Text
                Url = $item.SubItems[1].Text
                Status = $item.SubItems[2].Text
            }
        }
    }

    $multiForm.Dispose()
    return $selectedSites
}

<#
.FUNCTION Show-SitesBrowserDialog
.SYNOPSIS
    Displays all SharePoint sites with caching and multi-site selection
.DESCRIPTION
    Retrieves all active SharePoint sites, caches them, displays in ListView with sizes,
    allows users to select multiple sites for processing, includes Select All/None buttons
.EXAMPLE
    $sites = Show-SitesBrowserDialog
    Returns array of selected site objects with Title and Url properties
#>
function Show-SitesBrowserDialog {
    $script:sitesCache = @()
    $selectedSites = @()
    
    $browserForm = New-Object System.Windows.Forms.Form
    $browserForm.Text = "SharePoint Sites Browser - Select Sites to Process"
    $browserForm.Size = New-Object System.Drawing.Size(800, 600)
    $browserForm.StartPosition = "CenterScreen"
    $browserForm.FormBorderStyle = "FixedDialog"
    $browserForm.MaximizeBox = $false

    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 20)
    $statusLabel.Size = New-Object System.Drawing.Size(750, 25)
    $statusLabel.Text = "Loading all SharePoint sites..."
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $browserForm.Controls.Add($statusLabel)

    # ListView for sites
    $listView = New-Object System.Windows.Forms.ListView
    $listView.Location = New-Object System.Drawing.Point(20, 55)
    $listView.Size = New-Object System.Drawing.Size(750, 420)
    $listView.View = "Details"
    $listView.CheckBoxes = $true
    $listView.FullRowSelect = $false
    $listView.MultiSelect = $true
    
    # Add columns
    $listView.Columns.Add("Site Name", 250) | Out-Null
    $listView.Columns.Add("URL", 350) | Out-Null
    $listView.Columns.Add("Status", 100) | Out-Null
    
    $browserForm.Controls.Add($listView)

    # Select All button
    $selectAllButton = New-Object System.Windows.Forms.Button
    $selectAllButton.Location = New-Object System.Drawing.Point(20, 490)
    $selectAllButton.Size = New-Object System.Drawing.Size(100, 30)
    $selectAllButton.Text = "[+] Select All"
    $browserForm.Controls.Add($selectAllButton)

    # Select None button
    $selectNoneButton = New-Object System.Windows.Forms.Button
    $selectNoneButton.Location = New-Object System.Drawing.Point(130, 490)
    $selectNoneButton.Size = New-Object System.Drawing.Size(100, 30)
    $selectNoneButton.Text = "[-] Select None"
    $browserForm.Controls.Add($selectNoneButton)

    # OK button
    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(620, 490)
    $okButton.Size = New-Object System.Drawing.Size(75, 30)
    $okButton.Text = "OK"
    $browserForm.Controls.Add($okButton)

    # Cancel button
    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(705, 490)
    $cancelButton.Size = New-Object System.Drawing.Size(75, 30)
    $cancelButton.Text = "Cancel"
    $browserForm.Controls.Add($cancelButton)

    # Load sites BEFORE showing dialog (synchronously on main thread)
    try {
        # Use cached sites from startup instead of retrieving again
        if ($global:cachedAllSites -and $global:cachedAllSites.Count -gt 0) {
            Write-Host "Using cached site list ($($global:cachedAllSites.Count) sites)..." -ForegroundColor Green
            $sites = $global:cachedAllSites
        } else {
            Write-Host "Retrieving all active SharePoint sites..." -ForegroundColor Cyan
            $sites = Get-PnPTenantSite -Detailed | Where-Object { $_.Status -eq "Active" } | Sort-Object Title
        }
        
        foreach ($site in $sites) {
            $item = $listView.Items.Add($site.Title)
            $item.SubItems.Add($site.Url) | Out-Null
            $item.SubItems.Add($site.Status) | Out-Null
            $item.Tag = $site  # Store full site object
        }
        
        $statusLabel.Text = "Found $($sites.Count) active sites - Select sites to process"
    } catch {
        $statusLabel.Text = "Error loading sites: $_"
        Write-Host "Error loading sites: $_" -ForegroundColor Red
    }

    # Select All click handler
    $selectAllButton.Add_Click({
        foreach ($item in $listView.Items) {
            $item.Checked = $true
        }
    })

    # Select None click handler
    $selectNoneButton.Add_Click({
        foreach ($item in $listView.Items) {
            $item.Checked = $false
        }
    })

    # OK click handler
    $okButton.Add_Click({
        $selectedSites = @()
        foreach ($item in $listView.CheckedItems) {
            $selectedSites += $item.Tag
        }
        
        $script:selectedSitesCache = $selectedSites
        $browserForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $browserForm.Close()
    })

    # Cancel click handler
    $cancelButton.Add_Click({
        $browserForm.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $browserForm.Close()
    })

    $browserForm.AcceptButton = $okButton
    $browserForm.CancelButton = $cancelButton

    $result = $browserForm.ShowDialog()
    $browserForm.Dispose()
    
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $script:selectedSitesCache
    }
    
    return $null
}

function Show-RecycleBinBrowserDialog {
    param(
        [Parameter()]
        [string]$SiteUrl = ""
    )

    $recycleForm = New-Object System.Windows.Forms.Form
    $recycleForm.Text = "SharePoint Recycle Bin Browser"
    $recycleForm.Size = New-Object System.Drawing.Size(1000, 750)
    $recycleForm.StartPosition = "CenterScreen"
    $recycleForm.FormBorderStyle = "FixedDialog"
    $recycleForm.MaximizeBox = $false

    # Title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 15)
    $titleLabel.Size = New-Object System.Drawing.Size(960, 30)
    $titleLabel.Text = "Recycle Bin Browser - Search for deleted files"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 12, [System.Drawing.FontStyle]::Bold)
    $recycleForm.Controls.Add($titleLabel)

    # Site URL label and input
    $siteLabel = New-Object System.Windows.Forms.Label
    $siteLabel.Location = New-Object System.Drawing.Point(20, 50)
    $siteLabel.Size = New-Object System.Drawing.Size(100, 20)
    $siteLabel.Text = "Site URL:"
    $recycleForm.Controls.Add($siteLabel)

    $siteUrlBox = New-Object System.Windows.Forms.TextBox
    $siteUrlBox.Location = New-Object System.Drawing.Point(130, 48)
    $siteUrlBox.Size = New-Object System.Drawing.Size(650, 20)
    $siteUrlBox.Text = $SiteUrl
    $siteUrlBox.BackColor = [System.Drawing.Color]::White
    $recycleForm.Controls.Add($siteUrlBox)

    # Copy URL button
    $copyButton = New-Object System.Windows.Forms.Button
    $copyButton.Location = New-Object System.Drawing.Point(790, 47)
    $copyButton.Size = New-Object System.Drawing.Size(50, 22)
    $copyButton.Text = "Copy"
    $copyButton.Add_Click({
        if ($siteUrlBox.Text) {
            [System.Windows.Forms.Clipboard]::SetText($siteUrlBox.Text)
            [System.Windows.Forms.MessageBox]::Show("URL copied to clipboard", "Copied", "OK", "Information")
        } else {
            [System.Windows.Forms.MessageBox]::Show("No URL to copy", "Empty", "OK", "Warning")
        }
    })
    $recycleForm.Controls.Add($copyButton)

    # Instructions label
    $instructionsLabel = New-Object System.Windows.Forms.Label
    $instructionsLabel.Location = New-Object System.Drawing.Point(20, 72)
    $instructionsLabel.Size = New-Object System.Drawing.Size(820, 30)
    $instructionsLabel.Text = "INSTRUCTIONS: Click [Browse] to select a site, or manually enter a complete site URL (e.g., https://tenant.sharepoint.com/sites/mysite). Then click [Load Items]."
    $instructionsLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $instructionsLabel.Font = New-Object System.Drawing.Font("Arial", 8, [System.Drawing.FontStyle]::Italic)
    $instructionsLabel.AutoSize = $false
    $recycleForm.Controls.Add($instructionsLabel)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(840, 48)
    $browseButton.Size = New-Object System.Drawing.Size(60, 22)
    $browseButton.Text = "Browse"
    $recycleForm.Controls.Add($browseButton)

    # Search box
    $searchLabel = New-Object System.Windows.Forms.Label
    $searchLabel.Location = New-Object System.Drawing.Point(20, 108)
    $searchLabel.Size = New-Object System.Drawing.Size(100, 20)
    $searchLabel.Text = "Search:"
    $recycleForm.Controls.Add($searchLabel)

    $searchBox = New-Object System.Windows.Forms.TextBox
    $searchBox.Location = New-Object System.Drawing.Point(130, 106)
    $searchBox.Size = New-Object System.Drawing.Size(700, 20)
    $recycleForm.Controls.Add($searchBox)

    # Status label
    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 136)
    $statusLabel.Size = New-Object System.Drawing.Size(960, 20)
    $statusLabel.Text = "Ready. Enter site URL and click Load to retrieve recycle bin items."
    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 9, [System.Drawing.FontStyle]::Italic)
    $recycleForm.Controls.Add($statusLabel)

    # ListView for recycle bin items
    $recycleBinListView = New-Object System.Windows.Forms.ListView
    $recycleBinListView.Location = New-Object System.Drawing.Point(20, 163)
    $recycleBinListView.Size = New-Object System.Drawing.Size(960, 380)
    $recycleBinListView.View = "Details"
    $recycleBinListView.CheckBoxes = $true
    $recycleBinListView.FullRowSelect = $true
    $recycleBinListView.GridLines = $true
    $recycleBinListView.MultiSelect = $true

    $recycleBinListView.Columns.Add("Full Path", 450) | Out-Null
    $recycleBinListView.Columns.Add("Type", 70) | Out-Null
    $recycleBinListView.Columns.Add("Size (KB)", 80) | Out-Null
    $recycleBinListView.Columns.Add("Deleted Date", 130) | Out-Null
    $recycleBinListView.Columns.Add("Deleted By", 120) | Out-Null
    $recycleBinListView.Columns.Add("Stage", 60) | Out-Null

    $recycleForm.Controls.Add($recycleBinListView)

    # Load button
    $loadButton = New-Object System.Windows.Forms.Button
    $loadButton.Location = New-Object System.Drawing.Point(20, 655)
    $loadButton.Size = New-Object System.Drawing.Size(80, 30)
    $loadButton.Text = "Load Items"
    $loadButton.BackColor = [System.Drawing.Color]::LightBlue
    $recycleForm.Controls.Add($loadButton)

    # Refresh button (clear cache and reload)
    $refreshButton = New-Object System.Windows.Forms.Button
    $refreshButton.Location = New-Object System.Drawing.Point(105, 655)
    $refreshButton.Size = New-Object System.Drawing.Size(80, 30)
    $refreshButton.Text = "Refresh"
    $refreshButton.BackColor = [System.Drawing.Color]::LightCyan
    $recycleForm.Controls.Add($refreshButton)

    # Restore button
    $restoreButton = New-Object System.Windows.Forms.Button
    $restoreButton.Location = New-Object System.Drawing.Point(190, 655)
    $restoreButton.Size = New-Object System.Drawing.Size(100, 30)
    $restoreButton.Text = "[+] Restore Selected"
    $restoreButton.Enabled = $false
    $restoreButton.BackColor = [System.Drawing.Color]::LightGreen
    $recycleForm.Controls.Add($restoreButton)

    # Purge button
    $purgeButton = New-Object System.Windows.Forms.Button
    $purgeButton.Location = New-Object System.Drawing.Point(295, 655)
    $purgeButton.Size = New-Object System.Drawing.Size(100, 30)
    $purgeButton.Text = "[X] Purge Selected"
    $purgeButton.Enabled = $false
    $purgeButton.BackColor = [System.Drawing.Color]::LightCoral
    $recycleForm.Controls.Add($purgeButton)

    # Export to CSV button
    $exportCsvButton = New-Object System.Windows.Forms.Button
    $exportCsvButton.Location = New-Object System.Drawing.Point(400, 655)
    $exportCsvButton.Size = New-Object System.Drawing.Size(110, 30)
    $exportCsvButton.Text = "[CSV] Export to CSV"
    $exportCsvButton.BackColor = [System.Drawing.Color]::LightYellow
    $recycleForm.Controls.Add($exportCsvButton)

    # Export to XLSX button
    $exportXlsxButton = New-Object System.Windows.Forms.Button
    $exportXlsxButton.Location = New-Object System.Drawing.Point(515, 655)
    $exportXlsxButton.Size = New-Object System.Drawing.Size(110, 30)
    $exportXlsxButton.Text = "[XLSX] Export"
    $exportXlsxButton.BackColor = [System.Drawing.Color]::LightGreen
    $recycleForm.Controls.Add($exportXlsxButton)

    # Close button
    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(900, 655)
    $closeButton.Size = New-Object System.Drawing.Size(80, 30)
    $closeButton.Text = "Close"
    $closeButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $recycleForm.Controls.Add($closeButton)

    # Browse click handler
    $browseButton.Add_Click({
        try {
            Write-Log "Browse button clicked" "INFO"
            $statusLabel.Text = "Opening site selector... Please wait."
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            [System.Windows.Forms.Application]::DoEvents()
            
            $selected = Show-SiteSelectorDialog
            
            if ($selected -and -not [string]::IsNullOrWhiteSpace($selected)) {
                Write-Log "Site selected from browser: $selected" "INFO"
                $siteUrlBox.Text = $selected
                $statusLabel.Text = "✓ Site URL: $selected - Click Load Items to fetch recycle bin contents."
                $statusLabel.ForeColor = [System.Drawing.Color]::DarkGreen
            } else {
                Write-Log "No site selected from browser" "INFO"
                $statusLabel.Text = "No site selected. Either click Browse and select a site, or manually enter the site URL."
                $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            }
        } catch {
            Write-Log "Browse button error: $($_.Exception.Message)" "ERROR"
            $statusLabel.Text = "Error selecting site: $($_.Exception.Message)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("Browse Error: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    # Refresh button - clear cache and force re-fetch from SharePoint
    $refreshButton.Add_Click({
        try {
            $siteUrl = $siteUrlBox.Text.Trim()
            
            if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                [System.Windows.Forms.MessageBox]::Show("Please enter a site URL first", "No URL", "OK", "Warning")
                return
            }
            
            Write-Log "Refresh button clicked for: $siteUrl" "INFO"
            $statusLabel.Text = "Clearing cache and fetching fresh data..."
            $statusLabel.ForeColor = [System.Drawing.Color]::Orange
            [System.Windows.Forms.Application]::DoEvents()
            
            # Clear cache for this site
            $cacheKey = $siteUrl
            if ($global:recycleBinItemsCache.ContainsKey($cacheKey)) {
                $global:recycleBinItemsCache.Remove($cacheKey)
                $global:recycleBinCacheTimestamp.Remove($cacheKey)
                Write-Log "Cache cleared for: $siteUrl" "INFO"
            }
            
            # Trigger load by clicking Load button
            $loadButton.PerformClick()
            
        } catch {
            Write-Log "Refresh button error: $($_.Exception.Message)" "ERROR"
            $statusLabel.Text = "Error refreshing: $($_.Exception.Message)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            [System.Windows.Forms.MessageBox]::Show("Refresh Error: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    # Load click handler with global caching
    $script:recycleBinCache = @()
    
    # Initialize global recycle bin cache if not exists
    if (-not $global:recycleBinItemsCache) {
        $global:recycleBinItemsCache = @{}
        $global:recycleBinCacheTimestamp = @{}
    }
    
    $loadButton.Add_Click({
        try {
            $siteUrl = $siteUrlBox.Text.Trim()
            
            Write-Log "RecycleBin Load: Raw textbox value = '$($siteUrlBox.Text)'" "INFO"
            Write-Log "RecycleBin Load: After trim = '$siteUrl'" "INFO"
            
            # Enhanced validation
            if ([string]::IsNullOrWhiteSpace($siteUrl)) {
                $msg = "Site URL is empty.`n`nHow to fix:`n1. Click [Browse] to select a site, OR`n2. Manually copy-paste a site URL`n3. Then click Load Items`n`nExample URL: https://tenant.sharepoint.com/sites/mysiteorlistname"
                [System.Windows.Forms.MessageBox]::Show($msg, "Missing Site URL", "OK", "Warning")
                $statusLabel.Text = "ERROR: Site URL is required"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
                Write-Log "RecycleBin Load: Empty URL detected" "WARN"
                return
            }
            
            # Validate URL format
            if (-not ($siteUrl -like "https://*")) {
                $msg = "Invalid URL format.`n`nProvided: $siteUrl`n`nMust start with: https://`n`nExample: https://tenant.sharepoint.com/sites/mysite"
                [System.Windows.Forms.MessageBox]::Show($msg, "Invalid URL Format", "OK", "Warning")
                $statusLabel.Text = "ERROR: Invalid URL format"
                $statusLabel.ForeColor = [System.Drawing.Color]::Red
                Write-Log "RecycleBin Load: Invalid URL format - $siteUrl" "WARN"
                return
            }

            # Check if we have cached items for this site (cache expires after 10 minutes)
            $cacheKey = $siteUrl
            $useCache = $false
            
            if ($global:recycleBinItemsCache.ContainsKey($cacheKey)) {
                $cacheAge = (Get-Date) - $global:recycleBinCacheTimestamp[$cacheKey]
                if ($cacheAge.TotalMinutes -lt 10) {
                    Write-Log "Using cached recycle bin items for $siteUrl (age: $([math]::Round($cacheAge.TotalSeconds))s)" "INFO"
                    $statusLabel.Text = "Using cached recycle bin items (from cache)..."
                    $statusLabel.ForeColor = [System.Drawing.Color]::DarkBlue
                    $items = $global:recycleBinItemsCache[$cacheKey]
                    $useCache = $true
                } else {
                    Write-Log "Cache expired for $siteUrl - fetching fresh data" "INFO"
                    $statusLabel.Text = "Cache expired - fetching fresh data from SharePoint..."
                    $statusLabel.ForeColor = [System.Drawing.Color]::Blue
                }
            } else {
                $statusLabel.Text = "Fetching recycle bin items from SharePoint..."
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
            }
            
            [System.Windows.Forms.Application]::DoEvents()

            # If not using cache, fetch from SharePoint and cache
            if (-not $useCache) {
                Write-Log "RecycleBin Load: Calling Get-RecycleBinItems with URL: $siteUrl" "INFO"
                $items = Get-RecycleBinItems -SiteUrl $siteUrl
                
                # Store in global cache
                $global:recycleBinItemsCache[$cacheKey] = $items
                $global:recycleBinCacheTimestamp[$cacheKey] = Get-Date
                Write-Log "Cached $($items.Count) items for $siteUrl" "INFO"
            }
            
            $script:recycleBinCache = $items

            $recycleBinListView.Items.Clear()

            if ($items.Count -eq 0) {
                $statusLabel.Text = "No items found in recycle bin"
                $statusLabel.ForeColor = [System.Drawing.Color]::Gray
            } else {
                $statusLabel.Text = "Loading $($items.Count) items into display... Please wait."
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
                [System.Windows.Forms.Application]::DoEvents()
                
                # Use BeginUpdate/EndUpdate to batch UI updates - dramatically improves performance
                $recycleBinListView.BeginUpdate()
                
                $itemCount = 0
                foreach ($item in $items) {
                    # Build full path and add type indicator
                    $fullPath = if ($item.DirName) { "$($item.DirName)/$($item.LeafName)" } else { "/$($item.LeafName)" }
                    $itemTypeStr = [string]$item.ItemType
                    $typeIndicator = if ($itemTypeStr -eq "Folder") { "[F]" } else { "[D]" }
                    $displayPath = "$typeIndicator $fullPath"
                    
                    $listItem = New-Object System.Windows.Forms.ListViewItem($displayPath)
                    $listItem.SubItems.Add($itemTypeStr) | Out-Null
                    $sizeKB = if ($item.Size) { [math]::Round($item.Size / 1024, 2) } else { 0 }
                    $listItem.SubItems.Add([string]$sizeKB) | Out-Null
                    $listItem.SubItems.Add($item.DeletedDate.ToString("yyyy-MM-dd HH:mm")) | Out-Null
                    $listItem.SubItems.Add($item.DeletedByName) | Out-Null
                    $stage = if ($item.Stage -eq 1) { "First" } else { "Second" }
                    $listItem.SubItems.Add([string]$stage) | Out-Null
                    $listItem.Tag = $item
                    $recycleBinListView.Items.Add($listItem) | Out-Null
                    
                    $itemCount++
                    # Show progress every 1000 items
                    if (($itemCount % 1000) -eq 0) {
                        $statusLabel.Text = "Loading items... $itemCount / $($items.Count) processed"
                        [System.Windows.Forms.Application]::DoEvents()
                    }
                }
                
                # End batch update - applies all changes at once
                $recycleBinListView.EndUpdate()

                $statusLabel.Text = "Found $($items.Count) items in recycle bin"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
            }

            $restoreButton.Enabled = ($recycleBinListView.Items.Count -gt 0)
            $purgeButton.Enabled = ($recycleBinListView.Items.Count -gt 0)
            $exportCsvButton.Enabled = ($recycleBinListView.Items.Count -gt 0)
            $exportXlsxButton.Enabled = ($recycleBinListView.Items.Count -gt 0)

        } catch {
            $errorMsg = $_.Exception.Message
            $errorStack = $_.ScriptStackTrace
            
            Write-Log "RecycleBin Load Error: $errorMsg" "ERROR"
            Write-Log "Stack trace: $errorStack" "ERROR"
            
            $statusLabel.Text = "ERROR: Failed to load recycle bin"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            
            $fullErrorMsg = @"
Failed to load recycle bin items:

Error: $errorMsg

Troubleshooting:
1. Check that the site URL is correct
2. Verify you have permission to access the site
3. Check network connectivity
4. Review logs in Logs/ folder for details

URL attempted: $(if ($siteUrl) { $siteUrl } else { "[empty]" })
"@
            
            [System.Windows.Forms.MessageBox]::Show($fullErrorMsg, "Recycle Bin Load Failed", "OK", "Error")
        }
    })

    # Search text changed handler - simplified, no timer needed with BeginUpdate
    $script:checkedItemIds = @()
    
    $searchBox.Add_TextChanged({
        try {
            # Store currently checked item IDs before clearing
            $script:checkedItemIds = @()
            foreach ($item in $recycleBinListView.CheckedItems) {
                $script:checkedItemIds += $item.Tag.Id
            }

            $searchText = $searchBox.Text.ToLower()
            $recycleBinListView.Items.Clear()
            
            # Batch UI updates
            $recycleBinListView.BeginUpdate()
            
            $matchCount = 0
            foreach ($item in $script:recycleBinCache) {
                $fullPath = if ($item.DirName) { "$($item.DirName)/$($item.LeafName)" } else { "/$($item.LeafName)" }
                
                if ([string]::IsNullOrWhiteSpace($searchText) -or
                    $item.LeafName.ToLower().Contains($searchText) -or
                    $fullPath.ToLower().Contains($searchText) -or
                    $item.DirName.ToLower().Contains($searchText)) {

                    $itemTypeStr = [string]$item.ItemType
                    $typeIndicator = if ($itemTypeStr -eq "Folder") { "[F]" } else { "[D]" }
                    $displayPath = "$typeIndicator $fullPath"
                    
                    $listItem = New-Object System.Windows.Forms.ListViewItem($displayPath)
                    $listItem.SubItems.Add($itemTypeStr) | Out-Null
                    $sizeKB = if ($item.Size) { [math]::Round($item.Size / 1024, 2) } else { 0 }
                    $listItem.SubItems.Add([string]$sizeKB) | Out-Null
                    $listItem.SubItems.Add($item.DeletedDate.ToString("yyyy-MM-dd HH:mm")) | Out-Null
                    $listItem.SubItems.Add($item.DeletedByName) | Out-Null
                    $stage = if ($item.Stage -eq 1) { "First" } else { "Second" }
                    $listItem.SubItems.Add([string]$stage) | Out-Null
                    $listItem.Tag = $item
                    $listItem.Checked = ($script:checkedItemIds -contains $item.Id)
                    $recycleBinListView.Items.Add($listItem) | Out-Null
                    $matchCount++
                }
            }
            
            # End batch update - applies all changes at once
            $recycleBinListView.EndUpdate()
            
            $statusLabel.Text = "Found $matchCount items matching '$searchText'"
            $statusLabel.ForeColor = [System.Drawing.Color]::Blue
        } catch {
            Write-Log "Search error: $($_.Exception.Message)" "ERROR"
            $statusLabel.Text = "Search error: $($_.Exception.Message)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
        }
    })

    # Restore click handler - process all checked items
    $restoreButton.Add_Click({
        $checkedItems = @($recycleBinListView.CheckedItems)
        if ($checkedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select items to restore.", "No Selection", "OK", "Information")
            return
        }

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Restore $($checkedItems.Count) selected item(s) to their original locations?",
            "Confirm Restore",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $siteUrl = $siteUrlBox.Text.Trim()
            $successCount = 0
            $failCount = 0

            foreach ($listItem in $checkedItems) {
                $item = $listItem.Tag
                $statusLabel.Text = "Restoring $($item.LeafName)..."
                [System.Windows.Forms.Application]::DoEvents()

                $success = Restore-RecycleBinItem -SiteUrl $siteUrl -ItemId $item.Id

                if ($success) {
                    $successCount++
                    $recycleBinListView.Items.Remove($listItem)
                } else {
                    $failCount++
                }
            }

            $statusLabel.Text = "Restore complete: $successCount restored, $failCount failed"
            $statusLabel.ForeColor = if ($failCount -eq 0) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Orange }
            
            if ($successCount -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Successfully restored $successCount item(s)", "Restore Complete", "OK", "Information")
            }
            if ($failCount -gt 0) {
                [System.Windows.Forms.MessageBox]::Show("Failed to restore $failCount item(s)", "Partial Failure", "OK", "Warning")
            }
        }
    })

    # Purge click handler - process all checked items
    $purgeButton.Add_Click({
        $checkedItems = @($recycleBinListView.CheckedItems)
        if ($checkedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select items to purge.", "No Selection", "OK", "Information")
            return
        }

        $result = [System.Windows.Forms.MessageBox]::Show(
            "Permanently delete $($checkedItems.Count) selected item(s) from recycle bin? This cannot be undone!",
            "Confirm Permanent Deletion",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $result2 = [System.Windows.Forms.MessageBox]::Show(
                "Are you ABSOLUTELY sure? This is permanent!",
                "Final Confirmation",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )

            if ($result2 -eq [System.Windows.Forms.DialogResult]::Yes) {
                $siteUrl = $siteUrlBox.Text.Trim()
                $successCount = 0
                $failCount = 0

                foreach ($listItem in $checkedItems) {
                    $item = $listItem.Tag
                    $statusLabel.Text = "Purging $($item.LeafName)..."
                    [System.Windows.Forms.Application]::DoEvents()

                    $success = Remove-RecycleBinItem -SiteUrl $siteUrl -ItemId $item.Id

                    if ($success) {
                        $successCount++
                        $recycleBinListView.Items.Remove($listItem)
                    } else {
                        $failCount++
                    }
                }

                $statusLabel.Text = "Purge complete: $successCount purged, $failCount failed"
                $statusLabel.ForeColor = if ($failCount -eq 0) { [System.Drawing.Color]::Green } else { [System.Drawing.Color]::Orange }
                
                if ($successCount -gt 0) {
                    [System.Windows.Forms.MessageBox]::Show("Successfully purged $successCount item(s)", "Purge Complete", "OK", "Information")
                }
                if ($failCount -gt 0) {
                    [System.Windows.Forms.MessageBox]::Show("Failed to purge $failCount item(s)", "Partial Failure", "OK", "Warning")
                }
            }
        }
    })

    # Export to CSV button
    $exportCsvButton.Add_Click({
        try {
            if ($recycleBinListView.Items.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No items to export", "Empty List", "OK", "Information")
                return
            }

            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "CSV files (*.csv)|*.csv"
            $saveDialog.FileName = "RecycleBin_Export_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').csv"
            $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $exportPath = $saveDialog.FileName
                
                $statusLabel.Text = "Exporting $($recycleBinListView.Items.Count) items to CSV..."
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
                [System.Windows.Forms.Application]::DoEvents()

                # Build CSV content
                $csvData = @()
                $csvData += "Full Path,Type,Size (KB),Deleted Date,Deleted By,Stage"
                
                foreach ($item in $recycleBinListView.Items) {
                    $fullPath = $item.Text
                    $type = $item.SubItems[1].Text
                    $size = $item.SubItems[2].Text
                    $deletedDate = $item.SubItems[3].Text
                    $deletedBy = $item.SubItems[4].Text
                    $stage = $item.SubItems[5].Text
                    
                    # Escape quotes in CSV
                    $fullPath = $fullPath -replace '"', '""'
                    $type = $type -replace '"', '""'
                    $deletedBy = $deletedBy -replace '"', '""'
                    
                    $csvData += "`"$fullPath`",`"$type`",$size,`"$deletedDate`",`"$deletedBy`",$stage"
                }
                
                $csvData | Set-Content -Path $exportPath -Encoding UTF8
                
                $statusLabel.Text = "✓ Exported $($recycleBinListView.Items.Count) items to: $exportPath"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                Write-Log "Exported $($recycleBinListView.Items.Count) items to CSV: $exportPath" "INFO"
                
                [System.Windows.Forms.MessageBox]::Show("Successfully exported $($recycleBinListView.Items.Count) items to CSV`n`nLocation: $exportPath", "Export Complete", "OK", "Information")
            }
        } catch {
            $statusLabel.Text = "Error exporting to CSV: $($_.Exception.Message)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            Write-Log "CSV Export Error: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show("Export Error: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    # Export to XLSX button
    $exportXlsxButton.Add_Click({
        try {
            if ($recycleBinListView.Items.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No items to export", "Empty List", "OK", "Information")
                return
            }

            # Check if ImportExcel module is available, otherwise use CSV approach
            $moduleAvailable = Get-Module -ListAvailable -Name ImportExcel
            
            if (-not $moduleAvailable) {
                [System.Windows.Forms.MessageBox]::Show("ImportExcel module not available. Using CSV export instead.`n`nTo use XLSX export, install: Install-Module -Name ImportExcel", "XLSX Export", "OK", "Information")
                $exportCsvButton.PerformClick()
                return
            }

            $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
            $saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx"
            $saveDialog.FileName = "RecycleBin_Export_$(Get-Date -Format 'yyyy-MM-dd_HHmmss').xlsx"
            $saveDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")

            if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
                $exportPath = $saveDialog.FileName
                
                $statusLabel.Text = "Exporting $($recycleBinListView.Items.Count) items to XLSX..."
                $statusLabel.ForeColor = [System.Drawing.Color]::Blue
                [System.Windows.Forms.Application]::DoEvents()

                # Build array of objects for Excel export
                $exportData = @()
                foreach ($item in $recycleBinListView.Items) {
                    $exportData += [PSCustomObject]@{
                        "Full Path" = $item.Text
                        "Type" = $item.SubItems[1].Text
                        "Size (KB)" = $item.SubItems[2].Text
                        "Deleted Date" = $item.SubItems[3].Text
                        "Deleted By" = $item.SubItems[4].Text
                        "Stage" = $item.SubItems[5].Text
                    }
                }
                
                # Export to XLSX
                $exportData | Export-Excel -Path $exportPath -WorksheetName "Recycle Bin" -AutoSize -FreezeTopRow
                
                $statusLabel.Text = "✓ Exported $($recycleBinListView.Items.Count) items to: $exportPath"
                $statusLabel.ForeColor = [System.Drawing.Color]::Green
                Write-Log "Exported $($recycleBinListView.Items.Count) items to XLSX: $exportPath" "INFO"
                
                [System.Windows.Forms.MessageBox]::Show("Successfully exported $($recycleBinListView.Items.Count) items to XLSX`n`nLocation: $exportPath", "Export Complete", "OK", "Information")
            }
        } catch {
            $statusLabel.Text = "Error exporting to XLSX: $($_.Exception.Message)"
            $statusLabel.ForeColor = [System.Drawing.Color]::Red
            Write-Log "XLSX Export Error: $($_.Exception.Message)" "ERROR"
            [System.Windows.Forms.MessageBox]::Show("Export Error: $($_.Exception.Message)", "Error", "OK", "Error")
        }
    })

    # Enable/disable buttons based on content
    $recycleBinListView.Add_ItemChecked({
        $restoreButton.Enabled = ($recycleBinListView.CheckedItems.Count -gt 0)
        $purgeButton.Enabled = ($recycleBinListView.CheckedItems.Count -gt 0)
    })
    
    # Also enable export buttons when items exist (not just when checked)
    $recycleBinListView.Add_ItemSelectionChanged({
        $exportCsvButton.Enabled = ($recycleBinListView.Items.Count -gt 0)
        $exportXlsxButton.Enabled = ($recycleBinListView.Items.Count -gt 0)
    })

    $recycleForm.CancelButton = $closeButton
    $recycleForm.ShowDialog() | Out-Null
    $recycleForm.Dispose()
}

Export-ModuleMember -Function @(
    'Show-SiteSelectorDialog',
    'Show-InitialSetupDialog',
    'Show-MainMenu',
    'Show-ScanOptions',
    'Show-CleanupOptions',
    'Show-SingleSiteDialog',
    'Show-InactiveSitesDialog',
    'Show-MultiSiteSelectorDialog',
    'Show-SitesBrowserDialog',
    'Show-RecycleBinBrowserDialog'
)
