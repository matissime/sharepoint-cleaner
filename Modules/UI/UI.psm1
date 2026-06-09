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
                }
            }
        }
    })

    $loadButton.Add_Click({
        try {
            $statusLabel.Text = "Connecting to SharePoint..."
            $loadButton.Enabled = $false
            $loadButton.Text = "Loading..."
            $searchTextBox.Enabled = $false
            $selectButton.Enabled = $false
            $sitesListView.Items.Clear()
            $script:allSites = @()
            
            [System.Windows.Forms.Application]::DoEvents()

            if ($script:cachedSites.Count -gt 0 -and $script:lastSitesFetch -and ((Get-Date) - $script:lastSitesFetch).TotalMinutes -lt 15) {
                $statusLabel.Text = "Using cached site list..."
                [System.Windows.Forms.Application]::DoEvents()
                $sites = $script:cachedSites
            } else {
                Connect-SharePoint -Url $script:adminUrl -UseCache -SuppressSuccessMessage
                $statusLabel.Text = "Fetching site list..."
                [System.Windows.Forms.Application]::DoEvents()
                
                $sites = Get-PnPTenantSite | Where-Object { $_.Template -ne "SPSPERS" } | Sort-Object Title
                $script:cachedSites = $sites
                $script:lastSitesFetch = Get-Date
            }
                
            $statusLabel.Text = "Processing $($sites.Count) sites..."
            [System.Windows.Forms.Application]::DoEvents()
            
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
                if (($processedCount % 50) -eq 0) {
                    $statusLabel.Text = "Processed $processedCount of $($sites.Count) sites..."
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 10
                }
            }
            
            $statusLabel.Text = "Found $($sites.Count) SharePoint sites"
            $searchTextBox.Enabled = $true
                
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
    $menuForm.Size = New-Object System.Drawing.Size(500, 530)
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

    $disconnectButton = New-Object System.Windows.Forms.Button
    $disconnectButton.Location = New-Object System.Drawing.Point(50, 355)
    $disconnectButton.Size = New-Object System.Drawing.Size(400, 35)
    $disconnectButton.Text = "[-] Disconnect - Clear cached authentication tokens"
    $disconnectButton.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $disconnectButton.BackColor = [System.Drawing.Color]::LightGray
    $menuForm.Controls.Add($disconnectButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Location = New-Object System.Drawing.Point(50, 400)
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

Export-ModuleMember -Function @(
    'Show-SiteSelectorDialog',
    'Show-InitialSetupDialog',
    'Show-MainMenu',
    'Show-ScanOptions',
    'Show-CleanupOptions',
    'Show-SingleSiteDialog',
    'Show-InactiveSitesDialog',
    'Show-MultiSiteSelectorDialog',
    'Show-SitesBrowserDialog'
)
