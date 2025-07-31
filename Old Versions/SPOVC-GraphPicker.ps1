
if (-not (Get-Module -ListAvailable -Name "Microsoft.Graph")) {
    Write-Host "Installing Microsoft.Graph module..." -ForegroundColor Yellow
    Install-Module Microsoft.Graph -Force
}
Import-Module Microsoft.Graph.Files

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:clientId = ""
$script:adminUrl = ""
$script:keepVersions = 25
$script:logDirectory = "$PSScriptRoot\Logs"
$script:dateTime = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$script:logFile = "$script:logDirectory\SPVersionCleaner_$script:dateTime.log"
$script:cancelRequested = $false

if (!(Test-Path -Path $script:logDirectory)) {
    New-Item -ItemType Directory -Path $script:logDirectory -Force | Out-Null
}

function Ensure-PnPModule {
    param (
        [string]$ModuleName = "PnP.PowerShell"
    )

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "$ModuleName is not installed. Would you like to install it now?",
            "Module Required",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Install-Module $ModuleName -AllowClobber -Force -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show(
                    "$ModuleName has been installed successfully.",
                    "Installation Complete",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            } catch {
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to install $ModuleName. Error: $_",
                    "Installation Failed",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                exit 1
            }
        } else {
            [System.Windows.Forms.MessageBox]::Show(
                "The script requires $ModuleName to function. Exiting.",
                "Module Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            exit 1
        }
    }
}

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    Add-Content -Path $script:logFile -Value $logMessage
    
    if ($script:progressForm -and $script:logTextBox) {
        $script:logTextBox.Invoke([Action]{
            $script:logTextBox.AppendText("$logMessage`r`n")
            $script:logTextBox.ScrollToCaret()
        })
    }
}

function Format-FileSize {
    param(
        [Parameter(Mandatory=$true)]
        [long]$Size
    )
    
    if ($Size -lt 1KB) {
        return "$Size B"
    }
    elseif ($Size -lt 1MB) {
        return "{0:N2} KB" -f ($Size / 1KB)
    }
    elseif ($Size -lt 1GB) {
        return "{0:N2} MB" -f ($Size / 1MB)
    }
    else {
        return "{0:N2} GB" -f ($Size / 1GB)
    }
}

function Show-InitialSetupDialog {
    $setupForm = New-Object System.Windows.Forms.Form
    $setupForm.Text = "SharePoint Version Cleaner - Initial Setup"
    $setupForm.Size = New-Object System.Drawing.Size(500, 250)
    $setupForm.StartPosition = "CenterScreen"
    $setupForm.FormBorderStyle = "FixedDialog"
    $setupForm.MaximizeBox = $false
    $setupForm.MinimizeBox = $false

    $clientIdLabel = New-Object System.Windows.Forms.Label
    $clientIdLabel.Location = New-Object System.Drawing.Point(20, 20)
    $clientIdLabel.Size = New-Object System.Drawing.Size(100, 20)
    $clientIdLabel.Text = "Client ID:"
    $setupForm.Controls.Add($clientIdLabel)

    $clientIdTextBox = New-Object System.Windows.Forms.TextBox
    $clientIdTextBox.Location = New-Object System.Drawing.Point(130, 18)
    $clientIdTextBox.Size = New-Object System.Drawing.Size(320, 20)
    $setupForm.Controls.Add($clientIdTextBox)

    $adminUrlLabel = New-Object System.Windows.Forms.Label
    $adminUrlLabel.Location = New-Object System.Drawing.Point(20, 60)
    $adminUrlLabel.Size = New-Object System.Drawing.Size(100, 20)
    $adminUrlLabel.Text = "Admin URL:"
    $setupForm.Controls.Add($adminUrlLabel)

    $adminUrlTextBox = New-Object System.Windows.Forms.TextBox
    $adminUrlTextBox.Location = New-Object System.Drawing.Point(130, 58)
    $adminUrlTextBox.Size = New-Object System.Drawing.Size(320, 20)
    $adminUrlTextBox.Text = "https://yourtenant-admin.sharepoint.com"
    $setupForm.Controls.Add($adminUrlTextBox)

    $versionLabel = New-Object System.Windows.Forms.Label
    $versionLabel.Location = New-Object System.Drawing.Point(20, 100)
    $versionLabel.Size = New-Object System.Drawing.Size(150, 20)
    $versionLabel.Text = "Versions to keep:"
    $setupForm.Controls.Add($versionLabel)

    $versionNumeric = New-Object System.Windows.Forms.NumericUpDown
    $versionNumeric.Location = New-Object System.Drawing.Point(170, 98)
    $versionNumeric.Size = New-Object System.Drawing.Size(80, 20)
    $versionNumeric.Minimum = 1
    $versionNumeric.Maximum = 999
    $versionNumeric.Value = 25
    $setupForm.Controls.Add($versionNumeric)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(280, 150)
    $okButton.Size = New-Object System.Drawing.Size(80, 30)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $setupForm.Controls.Add($okButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(370, 150)
    $cancelButton.Size = New-Object System.Drawing.Size(80, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $setupForm.Controls.Add($cancelButton)

    $setupForm.AcceptButton = $okButton
    $setupForm.CancelButton = $cancelButton

    $result = $setupForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        if ([string]::IsNullOrWhiteSpace($clientIdTextBox.Text) -or [string]::IsNullOrWhiteSpace($adminUrlTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all required fields.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            $setupForm.Dispose()
            return Show-InitialSetupDialog
        }
        
        $script:clientId = $clientIdTextBox.Text.Trim()
        $script:adminUrl = $adminUrlTextBox.Text.Trim()
        $script:keepVersions = $versionNumeric.Value
        
        $setupForm.Dispose()
        return $true
    } else {
        $setupForm.Dispose()
        return $false
    }
}

function Show-MainMenu {
    $mainForm = New-Object System.Windows.Forms.Form
    $mainForm.Text = "SharePoint Version Cleaner"
    $mainForm.Size = New-Object System.Drawing.Size(600, 450)
    $mainForm.StartPosition = "CenterScreen"
    $mainForm.FormBorderStyle = "FixedDialog"
    $mainForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(560, 30)
    $titleLabel.Text = "SharePoint Version Cleaner"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 16, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $mainForm.Controls.Add($titleLabel)

    $settingsLabel = New-Object System.Windows.Forms.Label
    $settingsLabel.Location = New-Object System.Drawing.Point(20, 60)
    $settingsLabel.Size = New-Object System.Drawing.Size(560, 40)
    $settingsLabel.Text = "Current Settings:`nVersions to keep: $script:keepVersions | Admin URL: $script:adminUrl"
    $settingsLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $mainForm.Controls.Add($settingsLabel)

    $menuGroupBox = New-Object System.Windows.Forms.GroupBox
    $menuGroupBox.Location = New-Object System.Drawing.Point(20, 110)
    $menuGroupBox.Size = New-Object System.Drawing.Size(560, 250)
    $menuGroupBox.Text = "Select Operation"
    $mainForm.Controls.Add($menuGroupBox)

    $configButton = New-Object System.Windows.Forms.Button
    $configButton.Location = New-Object System.Drawing.Point(20, 30)
    $configButton.Size = New-Object System.Drawing.Size(520, 35)
    $configButton.Text = "1. Configure Settings (Versions to keep, Client ID, Admin URL)"
    $configButton.Add_Click({
        $mainForm.Hide()
        if (Show-InitialSetupDialog) {
            $settingsLabel.Text = "Current Settings:`nVersions to keep: $script:keepVersions | Admin URL: $script:adminUrl"
        }
        $mainForm.Show()
    })
    $menuGroupBox.Controls.Add($configButton)

    $scanButton = New-Object System.Windows.Forms.Button
    $scanButton.Location = New-Object System.Drawing.Point(20, 75)
    $scanButton.Size = New-Object System.Drawing.Size(520, 35)
    $scanButton.Text = "2. Scan Libraries (Calculate storage used by excess versions)"
    $scanButton.Add_Click({
        $mainForm.Hide()
        Show-ScanOptions
        $mainForm.Show()
    })
    $menuGroupBox.Controls.Add($scanButton)

    $cleanupButton = New-Object System.Windows.Forms.Button
    $cleanupButton.Location = New-Object System.Drawing.Point(20, 120)
    $cleanupButton.Size = New-Object System.Drawing.Size(520, 35)
    $cleanupButton.Text = "3. Perform Cleanup (Delete excess versions)"
    $cleanupButton.BackColor = [System.Drawing.Color]::LightCoral
    $cleanupButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This will DELETE old file versions. Are you sure you want to continue?",
            "Confirm Cleanup",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $mainForm.Hide()
            Show-CleanupOptions
            $mainForm.Show()
        }
    })
    $menuGroupBox.Controls.Add($cleanupButton)

    $logsButton = New-Object System.Windows.Forms.Button
    $logsButton.Location = New-Object System.Drawing.Point(20, 165)
    $logsButton.Size = New-Object System.Drawing.Size(520, 35)
    $logsButton.Text = "4. View Logs and Reports"
    $logsButton.Add_Click({
        if (Test-Path $script:logDirectory) {
            Start-Process explorer.exe -ArgumentList $script:logDirectory
        } else {
            [System.Windows.Forms.MessageBox]::Show("No logs directory found.", "Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
        }
    })
    $menuGroupBox.Controls.Add($logsButton)

    $exitButton = New-Object System.Windows.Forms.Button
    $exitButton.Location = New-Object System.Drawing.Point(250, 370)
    $exitButton.Size = New-Object System.Drawing.Size(100, 35)
    $exitButton.Text = "Exit"
    $exitButton.Add_Click({
        $mainForm.Close()
    })
    $mainForm.Controls.Add($exitButton)

    $mainForm.ShowDialog()
}

function Show-ScanOptions {
    $scanForm = New-Object System.Windows.Forms.Form
    $scanForm.Text = "Scan Options"
    $scanForm.Size = New-Object System.Drawing.Size(500, 300)
    $scanForm.StartPosition = "CenterScreen"
    $scanForm.FormBorderStyle = "FixedDialog"
    $scanForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Text = "Select Scan Scope"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $scanForm.Controls.Add($titleLabel)

    $scanAllButton = New-Object System.Windows.Forms.Button
    $scanAllButton.Location = New-Object System.Drawing.Point(50, 70)
    $scanAllButton.Size = New-Object System.Drawing.Size(400, 40)
    $scanAllButton.Text = "Scan All SharePoint Sites and Libraries"
    $scanAllButton.Add_Click({
        $scanForm.Hide()
        Start-ProcessingWithProgress -Mode "calculate" -ProcessAll $true
        $scanForm.Close()
    })
    $scanForm.Controls.Add($scanAllButton)

    $scanSingleButton = New-Object System.Windows.Forms.Button
    $scanSingleButton.Location = New-Object System.Drawing.Point(50, 120)
    $scanSingleButton.Size = New-Object System.Drawing.Size(400, 40)
    $scanSingleButton.Text = "Scan Single Site/Library"
    $scanSingleButton.Add_Click({
        $scanForm.Hide()
        Show-SingleSiteDialog -Mode "calculate"
        $scanForm.Close()
    })
    $scanForm.Controls.Add($scanSingleButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(200, 210)
    $backButton.Size = New-Object System.Drawing.Size(100, 30)
    $backButton.Text = "Back"
    $backButton.Add_Click({
        $scanForm.Close()
    })
    $scanForm.Controls.Add($backButton)

    $scanForm.ShowDialog()
}

function Show-CleanupOptions {
    $cleanupForm = New-Object System.Windows.Forms.Form
    $cleanupForm.Text = "Cleanup Options"
    $cleanupForm.Size = New-Object System.Drawing.Size(500, 350)
    $cleanupForm.StartPosition = "CenterScreen"
    $cleanupForm.FormBorderStyle = "FixedDialog"
    $cleanupForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Text = "Select Cleanup Scope"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $cleanupForm.Controls.Add($titleLabel)

    $dryRunAllButton = New-Object System.Windows.Forms.Button
    $dryRunAllButton.Location = New-Object System.Drawing.Point(50, 70)
    $dryRunAllButton.Size = New-Object System.Drawing.Size(400, 40)
    $dryRunAllButton.Text = "Dry Run - All Sites (Simulate cleanup)"
    $dryRunAllButton.BackColor = [System.Drawing.Color]::LightYellow
    $dryRunAllButton.Add_Click({
        $cleanupForm.Hide()
        Start-ProcessingWithProgress -Mode "dry" -ProcessAll $true
        $cleanupForm.Close()
    })
    $cleanupForm.Controls.Add($dryRunAllButton)

    $dryRunSingleButton = New-Object System.Windows.Forms.Button
    $dryRunSingleButton.Location = New-Object System.Drawing.Point(50, 120)
    $dryRunSingleButton.Size = New-Object System.Drawing.Size(400, 40)
    $dryRunSingleButton.Text = "Dry Run - Single Site/Library"
    $dryRunSingleButton.BackColor = [System.Drawing.Color]::LightYellow
    $dryRunSingleButton.Add_Click({
        $cleanupForm.Hide()
        Show-SingleSiteDialog -Mode "dry"
        $cleanupForm.Close()
    })
    $cleanupForm.Controls.Add($dryRunSingleButton)

    $cleanupAllButton = New-Object System.Windows.Forms.Button
    $cleanupAllButton.Location = New-Object System.Drawing.Point(50, 170)
    $cleanupAllButton.Size = New-Object System.Drawing.Size(400, 40)
    $cleanupAllButton.Text = "CLEANUP - All Sites (DELETE VERSIONS)"
    $cleanupAllButton.BackColor = [System.Drawing.Color]::Red
    $cleanupAllButton.ForeColor = [System.Drawing.Color]::White
    $cleanupAllButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This will PERMANENTLY DELETE old file versions from ALL SharePoint sites. This action cannot be undone. Are you absolutely sure?",
            "Final Warning",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $cleanupForm.Hide()
            Start-ProcessingWithProgress -Mode "run" -ProcessAll $true
            $cleanupForm.Close()
        }
    })
    $cleanupForm.Controls.Add($cleanupAllButton)

    $cleanupSingleButton = New-Object System.Windows.Forms.Button
    $cleanupSingleButton.Location = New-Object System.Drawing.Point(50, 220)
    $cleanupSingleButton.Size = New-Object System.Drawing.Size(400, 40)
    $cleanupSingleButton.Text = "CLEANUP - Single Site/Library (DELETE VERSIONS)"
    $cleanupSingleButton.BackColor = [System.Drawing.Color]::Red
    $cleanupSingleButton.ForeColor = [System.Drawing.Color]::White
    $cleanupSingleButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This will PERMANENTLY DELETE old file versions. This action cannot be undone. Are you sure?",
            "Confirm Cleanup",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $cleanupForm.Hide()
            Show-SingleSiteDialog -Mode "run"
            $cleanupForm.Close()
        }
    })
    $cleanupForm.Controls.Add($cleanupSingleButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(200, 280)
    $backButton.Size = New-Object System.Drawing.Size(100, 30)
    $backButton.Text = "Back"
    $backButton.Add_Click({
        $cleanupForm.Close()
    })
    $cleanupForm.Controls.Add($backButton)

    $cleanupForm.ShowDialog()
}

function Show-SingleSiteDialog {
    param(
        [string]$Mode
    )

    $singleForm = New-Object System.Windows.Forms.Form
    $singleForm.Text = "Single Site/Library Selection"
    $singleForm.Size = New-Object System.Drawing.Size(600, 350)
    $singleForm.StartPosition = "CenterScreen"
    $singleForm.FormBorderStyle = "FixedDialog"
    $singleForm.MaximizeBox = $false

    $siteUrlLabel = New-Object System.Windows.Forms.Label
    $siteUrlLabel.Location = New-Object System.Drawing.Point(20, 30)
    $siteUrlLabel.Size = New-Object System.Drawing.Size(100, 20)
    $siteUrlLabel.Text = "Site URL:"
    $singleForm.Controls.Add($siteUrlLabel)

    $siteUrlTextBox = New-Object System.Windows.Forms.TextBox
    $siteUrlTextBox.Location = New-Object System.Drawing.Point(130, 28)
    $siteUrlTextBox.Size = New-Object System.Drawing.Size(420, 20)
    $siteUrlTextBox.Text = "https://yourtenant.sharepoint.com/sites/sitename"
    $singleForm.Controls.Add($siteUrlTextBox)

    $libraryLabel = New-Object System.Windows.Forms.Label
    $libraryLabel.Location = New-Object System.Drawing.Point(20, 70)
    $libraryLabel.Size = New-Object System.Drawing.Size(100, 20)
    $libraryLabel.Text = "Library Name:"
    $singleForm.Controls.Add($libraryLabel)

    $libraryTextBox = New-Object System.Windows.Forms.TextBox
    $libraryTextBox.Location = New-Object System.Drawing.Point(130, 68)
    $libraryTextBox.Size = New-Object System.Drawing.Size(300, 20)
    $libraryTextBox.Text = "Documents"
    $singleForm.Controls.Add($libraryTextBox)

    $allLibrariesCheckBox = New-Object System.Windows.Forms.CheckBox
    $allLibrariesCheckBox.Location = New-Object System.Drawing.Point(130, 100)
    $allLibrariesCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $allLibrariesCheckBox.Text = "Process all libraries in this site (leave library name empty)"
    $singleForm.Controls.Add($allLibrariesCheckBox)

    $processButton = New-Object System.Windows.Forms.Button
    $processButton.Location = New-Object System.Drawing.Point(200, 200)
    $processButton.Size = New-Object System.Drawing.Size(100, 35)
    $processButton.Text = "Process"
    $processButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($siteUrlTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a site URL.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $singleForm.Hide()
        
        if ($allLibrariesCheckBox.Checked -or [string]::IsNullOrWhiteSpace($libraryTextBox.Text)) {
            Start-ProcessingWithProgress -Mode $Mode -ProcessAll $false -SiteUrl $siteUrlTextBox.Text.Trim()
        } else {
            Start-ProcessingWithProgress -Mode $Mode -ProcessAll $false -SiteUrl $siteUrlTextBox.Text.Trim() -LibraryName $libraryTextBox.Text.Trim()
        }
        
        $singleForm.Close()
    })
    $singleForm.Controls.Add($processButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(320, 200)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $singleForm.Close()
    })
    $singleForm.Controls.Add($cancelButton)

    $folderButton = New-Object System.Windows.Forms.Button
    $folderButton.Location = New-Object System.Drawing.Point(50, 140)
    $folderButton.Size = New-Object System.Drawing.Size(400, 40)
    $folderButton.Text = "Select Specific Folders"
    $folderButton.BackColor = [System.Drawing.Color]::LightBlue
    $folderButton.Add_Click({
    if ([string]::IsNullOrWhiteSpace($siteUrlTextBox.Text) -or [string]::IsNullOrWhiteSpace($libraryTextBox.Text)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter both Site URL and Library Name for folder selection.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    $singleForm.Hide()
    Show-FolderSelectionDialog -SiteUrl $siteUrlTextBox.Text.Trim() -LibraryName $libraryTextBox.Text.Trim() -Mode $Mode
    $singleForm.Close()
    })
    $singleForm.Controls.Add($folderButton)


    $singleForm.ShowDialog()
}

function Get-GraphFoldersRecursive {
    param(
        [string]$DriveId,
        [string]$ParentId = $null,
        [string]$PathPrefix = ""
    )
    $allFolders = @()
    if ($ParentId) {
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ParentId -All | Where-Object { $_.Folder }
    } else {
        $children = Get-MgDriveRootChild -DriveId $DriveId -All | Where-Object { $_.Folder }
    }
    foreach ($child in $children) {
        $thisPath = if ($PathPrefix) { "$PathPrefix/$($child.Name)" } else { $child.Name }
        $allFolders += [PSCustomObject]@{
            Name = $child.Name
            Id = $child.Id
            Path = $thisPath
        }
        $allFolders += Get-GraphFoldersRecursive -DriveId $DriveId -ParentId $child.Id -PathPrefix $thisPath
    }
    return $allFolders
}

function Show-FolderSelectionDialog {
    param (
        [string]$SiteSearchString,
        [string]$LibraryName
    )

    Connect-PnPOnline -Interactive `
                      -Scopes "Sites.Read.All","Files.Read.All"

    $site  = Get-PnPGraphSite   -Search $SiteSearchString | Select-Object -First 1
    if (-not $site) {
        [System.Windows.Forms.MessageBox]::Show("Site not found!") 
        return
    }
    $drive = Get-PnPGraphDrive  -SiteId $site.Id |
             Where-Object Name -eq $LibraryName
    if (-not $drive) {
        [System.Windows.Forms.MessageBox]::Show("Library not found!")
        return
    }

    function Get-FoldersRecursive {
        param($DriveId, $ParentId, $Path)
        $items = Get-PnPGraphFolder -DriveId $DriveId -ItemId $ParentId
        foreach ($f in $items) {
            $fullPath = if ($Path) { "$Path/$($f.Name)" } else { $f.Name }
            [PSCustomObject]@{ Name = $f.Name; Id = $f.Id; Path = $fullPath }
            Get-FoldersRecursive -DriveId $DriveId `
                                 -ParentId $f.Id `
                                 -Path $fullPath
        }
    }
    $allFolders = Get-FoldersRecursive -DriveId $drive.Id -ParentId ""
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object Windows.Forms.Form
    $form.Text = "Select Folder"
    $form.Size = [System.Drawing.Size]::new(700, 600)
    $tree = New-Object Windows.Forms.TreeView
    $tree.Dock = 'Fill'
    $tree.HideSelection = $false
    $root = $tree.Nodes.Add($LibraryName)
    foreach ($folder in $allFolders) {
        $parts = $folder.Path -split '/'
        $node = $root
        foreach ($part in $parts) {
            if ($part -ne "") {
                $found = $node.Nodes.Find($part, $false)
                if ($found.Count -eq 0) {
                    $newNode = $node.Nodes.Add($part, $part)
                    $node = $newNode
                } else {
                    $node = $found[0]
                }
            }
        }
    }
    $form.Controls.Add($tree)
    $okBtn = New-Object Windows.Forms.Button
    $okBtn.Text = "OK"
    $okBtn.Dock = 'Bottom'
    $okBtn.Add_Click({
        if ($tree.SelectedNode -and $tree.SelectedNode.Level -gt 0) {
            $form.Tag = $tree.SelectedNode.FullPath.Replace('', '/')
            $form.Close()
        }
    })
    $form.Controls.Add($okBtn)
    $form.ShowDialog() | Out-Null
    return $form.Tag
}

function Start-FolderProcessingWithProgress {
    param(
        [string]$Mode,
        [string]$SiteUrl,
        [string]$LibraryName,
        [array]$SelectedFolders
    )

    $script:progressForm = New-Object System.Windows.Forms.Form
    $script:progressForm.Text = "Processing Selected Folders"
    $script:progressForm.Size = New-Object System.Drawing.Size(800, 600)
    $script:progressForm.StartPosition = "CenterScreen"
    $script:progressForm.FormBorderStyle = "FixedDialog"
    $script:progressForm.MaximizeBox = $false

    $overallProgressLabel = New-Object System.Windows.Forms.Label
    $overallProgressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $overallProgressLabel.Size = New-Object System.Drawing.Size(760, 20)
    $overallProgressLabel.Text = "Overall Progress:"
    $script:progressForm.Controls.Add($overallProgressLabel)

    $overallProgressBar = New-Object System.Windows.Forms.ProgressBar
    $overallProgressBar.Location = New-Object System.Drawing.Point(20, 45)
    $overallProgressBar.Size = New-Object System.Drawing.Size(600, 25)
    $script:progressForm.Controls.Add($overallProgressBar)

    $overallETALabel = New-Object System.Windows.Forms.Label
    $overallETALabel.Location = New-Object System.Drawing.Point(630, 48)
    $overallETALabel.Size = New-Object System.Drawing.Size(150, 20)
    $overallETALabel.Text = "ETA: Calculating..."
    $script:progressForm.Controls.Add($overallETALabel)

    $currentProgressLabel = New-Object System.Windows.Forms.Label
    $currentProgressLabel.Location = New-Object System.Drawing.Point(20, 85)
    $currentProgressLabel.Size = New-Object System.Drawing.Size(760, 20)
    $currentProgressLabel.Text = "Current Folder:"
    $script:progressForm.Controls.Add($currentProgressLabel)

    $currentProgressBar = New-Object System.Windows.Forms.ProgressBar
    $currentProgressBar.Location = New-Object System.Drawing.Point(20, 110)
    $currentProgressBar.Size = New-Object System.Drawing.Size(600, 25)
    $script:progressForm.Controls.Add($currentProgressBar)

    $currentETALabel = New-Object System.Windows.Forms.Label
    $currentETALabel.Location = New-Object System.Drawing.Point(630, 113)
    $currentETALabel.Size = New-Object System.Drawing.Size(150, 20)
    $currentETALabel.Text = "ETA: Calculating..."
    $script:progressForm.Controls.Add($currentETALabel)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 150)
    $statusLabel.Size = New-Object System.Drawing.Size(760, 40)
    $statusLabel.Text = "Initializing..."
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $script:progressForm.Controls.Add($statusLabel)

    $script:logTextBox = New-Object System.Windows.Forms.TextBox
    $script:logTextBox.Location = New-Object System.Drawing.Point(20, 200)
    $script:logTextBox.Size = New-Object System.Drawing.Size(760, 300)
    $script:logTextBox.Multiline = $true
    $script:logTextBox.ScrollBars = "Vertical"
    $script:logTextBox.ReadOnly = $true
    $script:logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:progressForm.Controls.Add($script:logTextBox)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(350, 520)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $script:cancelRequested = $true
        $cancelButton.Enabled = $false
        $cancelButton.Text = "Cancelling..."
    })
    $script:progressForm.Controls.Add($cancelButton)

    $script:progressForm.Show()

    $backgroundJob = Start-Job -ScriptBlock {
        param($Mode, $SiteUrl, $LibraryName, $SelectedFolders, $ClientId, $KeepVersions, $LogFile)
        
        Import-Module PnP.PowerShell -Force
        
        function Write-JobLog {
            param([string]$Message, [string]$Level = "INFO")
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"
            Add-Content -Path $LogFile -Value $logMessage
            Write-Output @{Type="Log"; Message=$logMessage; Level=$Level}
        }
        
        function Format-JobFileSize {
            param([long]$Size)
            if ($Size -lt 1KB) { return "$Size B" }
            elseif ($Size -lt 1MB) { return "{0:N2} KB" -f ($Size / 1KB) }
            elseif ($Size -lt 1GB) { return "{0:N2} MB" -f ($Size / 1MB) }
            else { return "{0:N2} GB" -f ($Size / 1GB) }
        }
        
        try {
            Write-JobLog "Starting folder processing in $Mode mode" "INFO"
            Write-JobLog "Selected folders: $($SelectedFolders -join ', ')" "INFO"
            
            Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
            Write-JobLog "Connected to site: $SiteUrl" "SUCCESS"
            
            Write-Output @{Type="Progress"; Overall=@{Total=$SelectedFolders.Count; Current=0}}
            
            $processedFolders = 0
            $totalStats = @{
                Folders = $SelectedFolders.Count
                Files = 0
                ExcessVersions = 0
                ExcessSize = 0
                VersionsDeleted = 0
                Errors = 0
            }
            
            foreach ($folderPath in $SelectedFolders) {
                if ($script:cancelRequested) {
                    Write-JobLog "Processing cancelled by user" "WARNING"
                    break
                }
                
                $processedFolders++
                $folderName = Split-Path $folderPath -Leaf
                Write-Output @{Type="Status"; Message="Processing folder: $folderName ($processedFolders of $($SelectedFolders.Count))"}
                Write-Output @{Type="Progress"; Overall=@{Total=$SelectedFolders.Count; Current=$processedFolders}}
                
                try {
                    $folderRelativeUrl = $folderPath.Replace($SiteUrl, "")
                    $items = Get-PnPListItem -List $LibraryName -Query "<View><Query><Where><Eq><FieldRef Name='FileDirRef'/><Value Type='Text'>$folderRelativeUrl</Value></Eq></Where></Query></View>" |
                            Where-Object { $_.FileSystemObjectType -eq "File" }
                    
                    $totalStats.Files += $items.Count
                    Write-JobLog "Found $($items.Count) files in folder: $folderName" "INFO"
                    
                    Write-Output @{Type="Progress"; Current=@{Total=$items.Count; Current=0}}
                    
                    $processedFiles = 0
                    
                    foreach ($item in $items) {
                        if ($script:cancelRequested) { break }
                        
                        $processedFiles++
                        $fileRef = $item["FileRef"]
                        $fileName = Split-Path $fileRef -Leaf
                        Write-Output @{Type="CurrentFile"; File=$fileName}
                        Write-Output @{Type="Progress"; Current=@{Total=$items.Count; Current=$processedFiles}}
                        
                        try {
                            $fileItem = Get-PnPFile -Url $fileRef -AsListItem
                            $file = Get-PnPProperty -ClientObject $fileItem.File -Property Versions
                            $versionCount = $file.Count
                            
                            if ($versionCount -gt $KeepVersions) {
                                $versionsToDelete = $file | Sort-Object -Property Created -Descending | Select-Object -Skip $KeepVersions
                                $deleteCount = $versionsToDelete.Count
                                $totalStats.ExcessVersions += $deleteCount
                                
                                $excessSize = ($versionsToDelete | Measure-Object -Property Size -Sum).Sum
                                $totalStats.ExcessSize += $excessSize
                                
                                $formattedSize = Format-JobFileSize -Size $excessSize
                                Write-JobLog "File: $fileName - $deleteCount excess versions using $formattedSize" "INFO"
                                
                                if ($Mode -eq "run") {
                                    foreach ($version in $versionsToDelete) {
                                        $version.DeleteObject()
                                        $totalStats.VersionsDeleted++
                                    }
                                    Invoke-PnPQuery
                                    Write-JobLog "Deleted $deleteCount versions from $fileName" "SUCCESS"
                                }
                            }
                        } catch {
                            $totalStats.Errors++
                            Write-JobLog "Error processing file $fileName : $($_.Exception.Message)" "ERROR"
                        }
                    }
                    
                } catch {
                    $totalStats.Errors++
                    Write-JobLog "Error processing folder $folderName : $($_.Exception.Message)" "ERROR"
                }
            }
            
            Write-Output @{Type="Complete"; Stats=$totalStats}
            
        } catch {
            Write-JobLog "Critical error: $($_.Exception.Message)" "ERROR"
            Write-Output @{Type="Error"; Message=$_.Exception.Message}
        }
    } -ArgumentList $Mode, $SiteUrl, $LibraryName, $SelectedFolders, $script:clientId, $script:keepVersions, $script:logFile
    
    $overallStartTime = Get-Date
    $currentStartTime = Get-Date
    
    while ($backgroundJob.State -eq "Running") {
        if ($script:cancelRequested) {
            Stop-Job -Job $backgroundJob
            Remove-Job -Job $backgroundJob
            break
        }
        
        $output = Receive-Job -Job $backgroundJob
        
        foreach ($item in $output) {
            switch ($item.Type) {
                "Status" {
                    $statusLabel.Text = $item.Message
                    $currentStartTime = Get-Date 
                }
                "Progress" {
                    if ($item.Overall) {
                        $overallProgressBar.Maximum = $item.Overall.Total
                        $overallProgressBar.Value = [Math]::Min($item.Overall.Current, $item.Overall.Total)
                        
                        $elapsed = (Get-Date) - $overallStartTime
                        if ($item.Overall.Current -gt 0) {
                            $rate = $item.Overall.Current / $elapsed.TotalMinutes
                            $remaining = $item.Overall.Total - $item.Overall.Current
                            if ($rate -gt 0) {
                                $eta = [Math]::Round($remaining / $rate)
                                $overallETALabel.Text = "ETA: $eta min"
                            }
                        }
                    }
                    if ($item.Current) {
                        $currentProgressBar.Maximum = $item.Current.Total
                        $currentProgressBar.Value = [Math]::Min($item.Current.Current, $item.Current.Total)
                        
                        $elapsed = (Get-Date) - $currentStartTime
                        if ($item.Current.Current -gt 0) {
                            $rate = $item.Current.Current / $elapsed.TotalMinutes
                            $remaining = $item.Current.Total - $item.Current.Current
                            if ($rate -gt 0) {
                                $eta = [Math]::Round($remaining / $rate)
                                $currentETALabel.Text = "ETA: $eta min"
                            }
                        }
                    }
                }
                "CurrentFile" {
                    $currentProgressLabel.Text = "Current File: $($item.File)"
                }
                "Complete" {
                    $stats = $item.Stats
                    $formattedSize = Format-FileSize -Size $stats.ExcessSize
                    
                    $summaryMessage = @"
========== FOLDER PROCESSING COMPLETE ==========
Folders processed: $($stats.Folders)
Files processed: $($stats.Files)
Excess versions found: $($stats.ExcessVersions)
Total excess storage: $formattedSize
"@
                    
                    if ($Mode -eq "run") {
                        $summaryMessage += "`nVersions deleted: $($stats.VersionsDeleted)"
                    }
                    
                    $summaryMessage += "`nErrors encountered: $($stats.Errors)"
                    
                    Write-Log $summaryMessage "SUCCESS"
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        $summaryMessage,
                        "Processing Complete",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                    
                    $script:progressForm.Close()
                    return
                }
                "Error" {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Critical error occurred: $($item.Message)",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                    $script:progressForm.Close()
                    return
                }
            }
        }
        
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 500
    }
    
    if ($backgroundJob.State -ne "Running") {
        Remove-Job -Job $backgroundJob
    }
    
    if ($script:cancelRequested) {
        Write-Log "Folder processing was cancelled by user" "WARNING"
        [System.Windows.Forms.MessageBox]::Show(
            "Processing was cancelled.",
            "Cancelled",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    
    $script:progressForm.Close()
}

function Start-ProcessingWithProgress {
    param(
        [string]$Mode,
        [bool]$ProcessAll,
        [string]$SiteUrl = "",
        [string]$LibraryName = ""
    )

    $script:progressForm = New-Object System.Windows.Forms.Form
    $script:progressForm.Text = "Processing SharePoint Files"
    $script:progressForm.Size = New-Object System.Drawing.Size(800, 600)
    $script:progressForm.StartPosition = "CenterScreen"
    $script:progressForm.FormBorderStyle = "FixedDialog"
    $script:progressForm.MaximizeBox = $false

    $overallProgressLabel = New-Object System.Windows.Forms.Label
    $overallProgressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $overallProgressLabel.Size = New-Object System.Drawing.Size(760, 20)
    $overallProgressLabel.Text = "Overall Progress:"
    $script:progressForm.Controls.Add($overallProgressLabel)

    $overallProgressBar = New-Object System.Windows.Forms.ProgressBar
    $overallProgressBar.Location = New-Object System.Drawing.Point(20, 45)
    $overallProgressBar.Size = New-Object System.Drawing.Size(600, 25)
    $script:progressForm.Controls.Add($overallProgressBar)

    $overallETALabel = New-Object System.Windows.Forms.Label
    $overallETALabel.Location = New-Object System.Drawing.Point(630, 48)
    $overallETALabel.Size = New-Object System.Drawing.Size(150, 20)
    $overallETALabel.Text = "ETA: Calculating..."
    $script:progressForm.Controls.Add($overallETALabel)

    $currentProgressLabel = New-Object System.Windows.Forms.Label
    $currentProgressLabel.Location = New-Object System.Drawing.Point(20, 85)
    $currentProgressLabel.Size = New-Object System.Drawing.Size(760, 20)
    $currentProgressLabel.Text = "Current Site:"
    $script:progressForm.Controls.Add($currentProgressLabel)

    $currentProgressBar = New-Object System.Windows.Forms.ProgressBar
    $currentProgressBar.Location = New-Object System.Drawing.Point(20, 110)
    $currentProgressBar.Size = New-Object System.Drawing.Size(600, 25)
    $script:progressForm.Controls.Add($currentProgressBar)

    $currentETALabel = New-Object System.Windows.Forms.Label
    $currentETALabel.Location = New-Object System.Drawing.Point(630, 113)
    $currentETALabel.Size = New-Object System.Drawing.Size(150, 20)
    $currentETALabel.Text = "ETA: Calculating..."
    $script:progressForm.Controls.Add($currentETALabel)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 150)
    $statusLabel.Size = New-Object System.Drawing.Size(760, 40)
    $statusLabel.Text = "Initializing..."
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $script:progressForm.Controls.Add($statusLabel)

    $script:logTextBox = New-Object System.Windows.Forms.TextBox
    $script:logTextBox.Location = New-Object System.Drawing.Point(20, 200)
    $script:logTextBox.Size = New-Object System.Drawing.Size(760, 300)
    $script:logTextBox.Multiline = $true
    $script:logTextBox.ScrollBars = "Vertical"
    $script:logTextBox.ReadOnly = $true
    $script:logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:progressForm.Controls.Add($script:logTextBox)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(350, 520)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $script:cancelRequested = $true
        $cancelButton.Enabled = $false
        $cancelButton.Text = "Cancelling..."
    })
    $script:progressForm.Controls.Add($cancelButton)

    $script:progressForm.Show()

    $startTime = Get-Date
    
    $backgroundJob = Start-Job -ScriptBlock {
        param($Mode, $ProcessAll, $SiteUrl, $LibraryName, $ClientId, $AdminUrl, $KeepVersions, $LogFile)
        
        Import-Module PnP.PowerShell -Force
        
        function Write-JobLog {
            param([string]$Message, [string]$Level = "INFO")
            $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            $logMessage = "[$timestamp] [$Level] $Message"
            Add-Content -Path $LogFile -Value $logMessage
            Write-Output @{Type="Log"; Message=$logMessage; Level=$Level}
        }
        
        function Format-JobFileSize {
            param([long]$Size)
            if ($Size -lt 1KB) { return "$Size B" }
            elseif ($Size -lt 1MB) { return "{0:N2} KB" -f ($Size / 1KB) }
            elseif ($Size -lt 1GB) { return "{0:N2} MB" -f ($Size / 1MB) }
            else { return "{0:N2} GB" -f ($Size / 1GB) }
        }
        
        try {
            Write-JobLog "Starting processing in $Mode mode" "INFO"
            
            if ($ProcessAll) {
                Write-Output @{Type="Status"; Message="Connecting to SharePoint admin site..."}
                Connect-PnPOnline -Url $AdminUrl -Interactive -ClientId $ClientId
                Write-JobLog "Connected to admin site: $AdminUrl" "SUCCESS"
                
                Write-Output @{Type="Status"; Message="Retrieving SharePoint sites..."}
                $sites = Get-PnPTenantSite | Where-Object { $_.Template -ne "SPSPERS" }
                $totalSites = $sites.Count
                Write-JobLog "Found $totalSites sites to process" "INFO"
                
                Write-Output @{Type="Progress"; Overall=@{Total=$totalSites; Current=0}}
                
                $processedSites = 0
                $totalStats = @{
                    Libraries = 0
                    Files = 0
                    ExcessVersions = 0
                    ExcessSize = 0
                    VersionsDeleted = 0
                    Errors = 0
                }
                
                foreach ($site in $sites) {
                    if ($script:cancelRequested) {
                        Write-JobLog "Processing cancelled by user" "WARNING"
                        break
                    }
                    
                    $processedSites++
                    Write-Output @{Type="Status"; Message="Processing site: $($site.Url) ($processedSites of $totalSites)"}
                    Write-Output @{Type="Progress"; Overall=@{Total=$totalSites; Current=$processedSites}}
                    
                    try {
                        Connect-PnPOnline -Url $site.Url -Interactive -ClientId $ClientId
                        Write-JobLog "Connected to site: $($site.Url)" "SUCCESS"
                        
                        $libraries = Get-PnPList | Where-Object {
                            $_.BaseTemplate -eq 101 -and $_.EnableVersioning -eq $true
                        }
                        
                        $totalStats.Libraries += $libraries.Count
                        Write-JobLog "Found $($libraries.Count) document libraries with versioning" "INFO"
                        
                        $libraryIndex = 0
                        
                        foreach ($library in $libraries) {
                            if ($script:cancelRequested) { break }
                            
                            $libraryIndex++
                            Write-Output @{Type="Status"; Message="Processing library: $($library.Title) ($libraryIndex of $($libraries.Count))"}
                            Write-Output @{Type="Progress"; Current=@{Total=$libraries.Count; Current=$libraryIndex}}
                            
                            $items = Get-PnPListItem -List $library.Title -PageSize 1000 -Fields "FileRef" | 
                                    Where-Object { $_.FileSystemObjectType -eq "File" }
                            
                            $totalStats.Files += $items.Count
                            
                            foreach ($item in $items) {
                                if ($script:cancelRequested) { break }
                                
                                $fileRef = $item["FileRef"]
                                Write-Output @{Type="CurrentFile"; File=$fileRef}
                                
                                try {
                                    $fileItem = Get-PnPFile -Url $fileRef -AsListItem
                                    $file = Get-PnPProperty -ClientObject $fileItem.File -Property Versions
                                    $versionCount = $file.Count
                                    
                                    if ($versionCount -gt $KeepVersions) {
                                        $versionsToDelete = $file | Sort-Object -Property Created -Descending | Select-Object -Skip $KeepVersions
                                        $deleteCount = $versionsToDelete.Count
                                        $totalStats.ExcessVersions += $deleteCount
                                        
                                        $excessSize = ($versionsToDelete | Measure-Object -Property Size -Sum).Sum
                                        $totalStats.ExcessSize += $excessSize
                                        
                                        $formattedSize = Format-JobFileSize -Size $excessSize
                                        Write-JobLog "File: $fileRef - $deleteCount excess versions using $formattedSize" "INFO"
                                        
                                        if ($Mode -eq "run") {
                                            foreach ($version in $versionsToDelete) {
                                                $version.DeleteObject()
                                                $totalStats.VersionsDeleted++
                                            }
                                            Invoke-PnPQuery
                                            Write-JobLog "Deleted $deleteCount versions from $fileRef" "SUCCESS"
                                        }
                                    }
                                } catch {
                                    $totalStats.Errors++
                                    Write-JobLog "Error processing file $fileRef : $($_.Exception.Message)" "ERROR"
                                }
                            }
                        }
                    } catch {
                        $totalStats.Errors++
                        Write-JobLog "Error processing site $($site.Url): $($_.Exception.Message)" "ERROR"
                    }
                }
            } else {
                Write-Output @{Type="Status"; Message="Processing single site/library..."}
                
                try {
                    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $ClientId
                    Write-JobLog "Connected to site: $SiteUrl" "SUCCESS"
                    
                    if ($LibraryName) {
                        $libraries = @(Get-PnPList -Identity $LibraryName)
                    } else {
                        $libraries = Get-PnPList | Where-Object {
                            $_.BaseTemplate -eq 101 -and $_.EnableVersioning -eq $true
                        }
                    }
                    
                    Write-Output @{Type="Progress"; Overall=@{Total=$libraries.Count; Current=0}}
                    
                    $processedLibraries = 0
                    $totalStats = @{
                        Libraries = $libraries.Count
                        Files = 0
                        ExcessVersions = 0
                        ExcessSize = 0
                        VersionsDeleted = 0
                        Errors = 0
                    }
                    
                    foreach ($library in $libraries) {
                        if ($script:cancelRequested) { break }
                        
                        $processedLibraries++
                        Write-Output @{Type="Status"; Message="Processing library: $($library.Title) ($processedLibraries of $($libraries.Count))"}
                        Write-Output @{Type="Progress"; Overall=@{Total=$libraries.Count; Current=$processedLibraries}}
                        
                        $items = Get-PnPListItem -List $library.Title -PageSize 1000 -Fields "FileRef" | 
                                Where-Object { $_.FileSystemObjectType -eq "File" }
                        
                        $totalStats.Files += $items.Count
                        Write-Output @{Type="Progress"; Current=@{Total=$items.Count; Current=0}}
                        
                        $processedFiles = 0
                        
                        foreach ($item in $items) {
                            if ($script:cancelRequested) { break }
                            
                            $processedFiles++
                            $fileRef = $item["FileRef"]
                            Write-Output @{Type="CurrentFile"; File=$fileRef}
                            Write-Output @{Type="Progress"; Current=@{Total=$items.Count; Current=$processedFiles}}
                            
                            try {
                                $fileItem = Get-PnPFile -Url $fileRef -AsListItem
                                $file = Get-PnPProperty -ClientObject $fileItem.File -Property Versions
                                $versionCount = $file.Count
                                
                                if ($versionCount -gt $KeepVersions) {
                                    $versionsToDelete = $file | Sort-Object -Property Created -Descending | Select-Object -Skip $KeepVersions
                                    $deleteCount = $versionsToDelete.Count
                                    $totalStats.ExcessVersions += $deleteCount
                                    
                                    $excessSize = ($versionsToDelete | Measure-Object -Property Size -Sum).Sum
                                    $totalStats.ExcessSize += $excessSize
                                    
                                    $formattedSize = Format-JobFileSize -Size $excessSize
                                    Write-JobLog "File: $fileRef - $deleteCount excess versions using $formattedSize" "INFO"
                                    
                                    if ($Mode -eq "run") {
                                        foreach ($version in $versionsToDelete) {
                                            $version.DeleteObject()
                                            $totalStats.VersionsDeleted++
                                        }
                                        Invoke-PnPQuery
                                        Write-JobLog "Deleted $deleteCount versions from $fileRef" "SUCCESS"
                                    }
                                }
                            } catch {
                                $totalStats.Errors++
                                Write-JobLog "Error processing file $fileRef : $($_.Exception.Message)" "ERROR"
                            }
                        }
                    }
                } catch {
                    Write-JobLog "Error connecting to site $SiteUrl : $($_.Exception.Message)" "ERROR"
                }
            }
            
            Write-Output @{Type="Complete"; Stats=$totalStats}
            
        } catch {
            Write-JobLog "Critical error: $($_.Exception.Message)" "ERROR"
            Write-Output @{Type="Error"; Message=$_.Exception.Message}
        }
    } -ArgumentList $Mode, $ProcessAll, $SiteUrl, $LibraryName, $script:clientId, $script:adminUrl, $script:keepVersions, $script:logFile
    
    $overallStartTime = Get-Date
    $currentStartTime = Get-Date
    
    while ($backgroundJob.State -eq "Running") {
        if ($script:cancelRequested) {
            Stop-Job -Job $backgroundJob
            Remove-Job -Job $backgroundJob
            break
        }
        
        $output = Receive-Job -Job $backgroundJob
        
        foreach ($item in $output) {
            switch ($item.Type) {
                "Status" {
                    $statusLabel.Text = $item.Message
                    $currentStartTime = Get-Date 
                }
                "Progress" {
                    if ($item.Overall) {
                        $overallProgressBar.Maximum = $item.Overall.Total
                        $overallProgressBar.Value = [Math]::Min($item.Overall.Current, $item.Overall.Total)
                        
                        $elapsed = (Get-Date) - $overallStartTime
                        if ($item.Overall.Current -gt 0) {
                            $rate = $item.Overall.Current / $elapsed.TotalMinutes
                            $remaining = $item.Overall.Total - $item.Overall.Current
                            if ($rate -gt 0) {
                                $eta = [Math]::Round($remaining / $rate)
                                $overallETALabel.Text = "ETA: $eta min"
                            }
                        }
                    }
                    if ($item.Current) {
                        $currentProgressBar.Maximum = $item.Current.Total
                        $currentProgressBar.Value = [Math]::Min($item.Current.Current, $item.Current.Total)
                        
                        $elapsed = (Get-Date) - $currentStartTime
                        if ($item.Current.Current -gt 0) {
                            $rate = $item.Current.Current / $elapsed.TotalMinutes
                            $remaining = $item.Current.Total - $item.Current.Current
                            if ($rate -gt 0) {
                                $eta = [Math]::Round($remaining / $rate)
                                $currentETALabel.Text = "ETA: $eta min"
                            }
                        }
                    }
                }
                "CurrentFile" {
                    $fileName = Split-Path $item.File -Leaf
                    $currentProgressLabel.Text = "Current File: $fileName"
                }
                "Log" {
                }
                "Complete" {
                    $stats = $item.Stats
                    $formattedSize = Format-FileSize -Size $stats.ExcessSize
                    
                    $summaryMessage = @"
========== PROCESSING COMPLETE ==========
Libraries processed: $($stats.Libraries)
Files processed: $($stats.Files)
Excess versions found: $($stats.ExcessVersions)
Total excess storage: $formattedSize
"@
                    
                    if ($Mode -eq "run") {
                        $summaryMessage += "`nVersions deleted: $($stats.VersionsDeleted)"
                    }
                    
                    $summaryMessage += "`nErrors encountered: $($stats.Errors)"
                    
                    Write-Log $summaryMessage "SUCCESS"
                    
                    [System.Windows.Forms.MessageBox]::Show(
                        $summaryMessage,
                        "Processing Complete",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )
                    
                    $script:progressForm.Close()
                    return
                }
                "Error" {
                    [System.Windows.Forms.MessageBox]::Show(
                        "Critical error occurred: $($item.Message)",
                        "Error",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Error
                    )
                    $script:progressForm.Close()
                    return
                }
            }
        }
        
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 500
    }
    
    if ($backgroundJob.State -ne "Running") {
        Remove-Job -Job $backgroundJob
    }
    
    if ($script:cancelRequested) {
        Write-Log "Processing was cancelled by user" "WARNING"
        [System.Windows.Forms.MessageBox]::Show(
            "Processing was cancelled.",
            "Cancelled",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
    }
    
    $script:progressForm.Close()
}

Write-Host "SharePoint Version Cleaner - GUI Edition" -ForegroundColor Cyan
Write-Host "=========================================" -ForegroundColor Cyan

Ensure-PnPModule

if (Show-InitialSetupDialog) {
    Write-Log "Script initialized with Client ID: $script:clientId" "INFO"
    Write-Log "Admin URL: $script:adminUrl" "INFO"
    Write-Log "Versions to keep: $script:keepVersions" "INFO"
    
    Show-MainMenu
} else {
    Write-Host "Setup cancelled. Exiting..." -ForegroundColor Yellow
}