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

    $scanFolderButton = New-Object System.Windows.Forms.Button
    $scanFolderButton.Location = New-Object System.Drawing.Point(50, 170)
    $scanFolderButton.Size = New-Object System.Drawing.Size(400, 40)
    $scanFolderButton.Text = "Scan Single Folder"
    $scanFolderButton.Add_Click({
        $scanForm.Hide()
        Show-FolderDialog -Mode "calculate"
        $scanForm.Close()
    })
    $scanForm.Controls.Add($scanFolderButton)


    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(200, 260)
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
    $cleanupForm.Size = New-Object System.Drawing.Size(500, 450)
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

    $dryRunCsvButton = New-Object System.Windows.Forms.Button
    $dryRunCsvButton.Location = New-Object System.Drawing.Point(50, 120)
    $dryRunCsvButton.Size = New-Object System.Drawing.Size(400, 40)
    $dryRunCsvButton.Text = "Dry Run - Process Sites from CSV"
    $dryRunCsvButton.BackColor = [System.Drawing.Color]::LightYellow
    $dryRunCsvButton.Add_Click({
        $cleanupForm.Hide()
        Show-CsvProcessingDialog -Mode "dry"
        $cleanupForm.Close()
    })
    $cleanupForm.Controls.Add($dryRunCsvButton)

    $dryRunSingleButton = New-Object System.Windows.Forms.Button
    $dryRunSingleButton.Location = New-Object System.Drawing.Point(50, 170)
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
    $cleanupAllButton.Location = New-Object System.Drawing.Point(50, 220)
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

    $cleanupCsvButton = New-Object System.Windows.Forms.Button
    $cleanupCsvButton.Location = New-Object System.Drawing.Point(50, 270)
    $cleanupCsvButton.Size = New-Object System.Drawing.Size(400, 40)
    $cleanupCsvButton.Text = "CLEANUP - Process Sites from CSV (DELETE VERSIONS)"
    $cleanupCsvButton.BackColor = [System.Drawing.Color]::Red
    $cleanupCsvButton.ForeColor = [System.Drawing.Color]::White
    $cleanupCsvButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This will PERMANENTLY DELETE old file versions from the sites listed in the CSV file. This action cannot be undone. Are you sure?",
            "Confirm CSV Cleanup",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $cleanupForm.Hide()
            Show-CsvProcessingDialog -Mode "run"
            $cleanupForm.Close()
        }
    })
    $cleanupForm.Controls.Add($cleanupCsvButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(200, 380)
    $backButton.Size = New-Object System.Drawing.Size(100, 30)
    $backButton.Text = "Back"
    $backButton.Add_Click({
        $cleanupForm.Close()
    })
    $cleanupForm.Controls.Add($backButton)

    $cleanupForm.ShowDialog()
}

function Show-CsvProcessingDialog {
    param(
        [string]$Mode
    )

    $csvForm = New-Object System.Windows.Forms.Form
    $csvForm.Text = "CSV Site Processing"
    $csvForm.Size = New-Object System.Drawing.Size(600, 400)
    $csvForm.StartPosition = "CenterScreen"
    $csvForm.FormBorderStyle = "FixedDialog"
    $csvForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(560, 40)
    $titleLabel.Text = "Select CSV File with SharePoint Sites"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $csvForm.Controls.Add($titleLabel)

    $instructionLabel = New-Object System.Windows.Forms.Label
    $instructionLabel.Location = New-Object System.Drawing.Point(20, 70)
    $instructionLabel.Size = New-Object System.Drawing.Size(560, 60)
    $instructionLabel.Text = "The CSV file should contain a column with SharePoint site URLs.`nThe column header should be 'SiteUrl' or 'URL'."
    $instructionLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $csvForm.Controls.Add($instructionLabel)

    $csvPathLabel = New-Object System.Windows.Forms.Label
    $csvPathLabel.Location = New-Object System.Drawing.Point(20, 140)
    $csvPathLabel.Size = New-Object System.Drawing.Size(100, 20)
    $csvPathLabel.Text = "CSV File:"
    $csvForm.Controls.Add($csvPathLabel)

    $csvPathTextBox = New-Object System.Windows.Forms.TextBox
    $csvPathTextBox.Location = New-Object System.Drawing.Point(130, 138)
    $csvPathTextBox.Size = New-Object System.Drawing.Size(380, 20)
    $csvPathTextBox.ReadOnly = $true
    $csvForm.Controls.Add($csvPathTextBox)

    $browseButton = New-Object System.Windows.Forms.Button
    $browseButton.Location = New-Object System.Drawing.Point(520, 136)
    $browseButton.Size = New-Object System.Drawing.Size(60, 25)
    $browseButton.Text = "Browse"
    $browseButton.Add_Click({
        $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $openFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        $openFileDialog.Title = "Select CSV File with SharePoint Sites"
        
        if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $csvPathTextBox.Text = $openFileDialog.FileName
        }
    })
    $csvForm.Controls.Add($browseButton)

    $processButton = New-Object System.Windows.Forms.Button
    $processButton.Location = New-Object System.Drawing.Point(200, 300)
    $processButton.Size = New-Object System.Drawing.Size(100, 35)
    $processButton.Text = "Process"
    $processButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($csvPathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please select a CSV file.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }

        if (-not (Test-Path $csvPathTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("The selected CSV file does not exist.", "File Not Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
            return
        }

        try {
            $csvData = Import-Csv -Path $csvPathTextBox.Text
            $urlColumn = $null

            # Check for URL column
            if ($csvData[0].PSObject.Properties.Name -contains "SiteUrl") {
                $urlColumn = "SiteUrl"
            } elseif ($csvData[0].PSObject.Properties.Name -contains "URL") {
                $urlColumn = "URL"
            } else {
                [System.Windows.Forms.MessageBox]::Show("The CSV file must contain a column named 'SiteUrl' or 'URL'.", "Invalid CSV Format", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
                return
            }

            $sites = $csvData | ForEach-Object { $_.$urlColumn } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            if ($sites.Count -eq 0) {
                [System.Windows.Forms.MessageBox]::Show("No valid site URLs found in the CSV file.", "No Sites Found", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
                return
            }

            $csvForm.Hide()
            Start-CsvProcessingWithProgress -Mode $Mode -Sites $sites
            $csvForm.Close()
        } catch {
            [System.Windows.Forms.MessageBox]::Show("Error reading CSV file: $($_.Exception.Message)", "Error", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
        }
    })
    $csvForm.Controls.Add($processButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(320, 300)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $csvForm.Close()
    })
    $csvForm.Controls.Add($cancelButton)

    $csvForm.ShowDialog()
}

function Start-CsvProcessingWithProgress {
    param(
        [string]$Mode,
        [array]$Sites
    )

    $script:progressForm = New-Object System.Windows.Forms.Form
    $script:progressForm.Text = "Processing CSV Sites"
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

    $backgroundJob = Start-Job -ScriptBlock {
        param($Mode, $Sites, $ClientId, $KeepVersions, $LogFile)
        
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
            Write-JobLog "Starting CSV site processing in $Mode mode" "INFO"
            
            $totalStats = @{
                Sites = $Sites.Count
                Libraries = 0
                Files = 0
                ExcessVersions = 0
                ExcessSize = 0
                VersionsDeleted = 0
                Errors = 0
            }
            
            Write-Output @{Type="Progress"; Overall=@{Total=$Sites.Count; Current=0}}
            
            $processedSites = 0
            
            foreach ($siteUrl in $Sites) {
                if ($script:cancelRequested) {
                    Write-JobLog "Processing cancelled by user" "WARNING"
                    break
                }
                
                $processedSites++
                Write-Output @{Type="Status"; Message="Processing site: $siteUrl ($processedSites of $($Sites.Count))"}
                Write-Output @{Type="Progress"; Overall=@{Total=$Sites.Count; Current=$processedSites}}
                
                try {
                    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId
                    Write-JobLog "Connected to site: $siteUrl" "SUCCESS"
                    
                    $libraries = Get-PnPList | Where-Object {
                        $_.BaseTemplate -eq 101 -and $_.EnableVersioning -eq $true
                    }
                    
                    $totalStats.Libraries += $libraries.Count
                    Write-JobLog "Found $($libraries.Count) document libraries with versioning" "INFO"
                    
                    Write-Output @{Type="Progress"; Current=@{Total=$libraries.Count; Current=0}}
                    $processedLibraries = 0
                    
                    foreach ($library in $libraries) {
                        if ($script:cancelRequested) { break }
                        
                        $processedLibraries++
                        Write-Output @{Type="Status"; Message="Processing library: $($library.Title) ($processedLibraries of $($libraries.Count))"}
                        Write-Output @{Type="Progress"; Current=@{Total=$libraries.Count; Current=$processedLibraries}}
                        
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
                    Write-JobLog "Error processing site $siteUrl : $($_.Exception.Message)" "ERROR"
                }
            }
            
            Write-Output @{Type="Complete"; Stats=$totalStats}
            
        } catch {
            Write-JobLog "Critical error: $($_.Exception.Message)" "ERROR"
            Write-Output @{Type="Error"; Message=$_.Exception.Message}
        }
    } -ArgumentList $Mode, $Sites, $script:clientId, $script:keepVersions, $script:logFile
    
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
                    if ($script:logTextBox) {
                        $script:logTextBox.AppendText("$($item.Message)`r`n")
                        $script:logTextBox.ScrollToCaret()
                    }
                }
                "Complete" {
                    $stats = $item.Stats
                    $formattedSize = Format-FileSize -Size $stats.ExcessSize
                    
                    $summaryMessage = @"
========== CSV PROCESSING COMPLETE ==========
Sites processed: $($stats.Sites)
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
                        "CSV Processing Complete",
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

function Show-SingleSiteDialog {
    param(
        [string]$Mode
    )

    $singleForm = New-Object System.Windows.Forms.Form
    $singleForm.Text = "Single Site/Library Selection"
    $singleForm.Size = New-Object System.Drawing.Size(600, 300)
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
    $processButton.Location = New-Object System.Drawing.Point(200, 160)
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
    $cancelButton.Location = New-Object System.Drawing.Point(320, 160)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $singleForm.Close()
    })
    $singleForm.Controls.Add($cancelButton)

    $singleForm.ShowDialog()
}

function Show-FolderDialog {
    param(
        [string]$Mode
    )

    $folderForm = New-Object System.Windows.Forms.Form
    $folderForm.Text = "Folder Selection"
    $folderForm.Size = New-Object System.Drawing.Size(700, 250)
    $folderForm.StartPosition = "CenterScreen"
    $folderForm.FormBorderStyle = "FixedDialog"
    $folderForm.MaximizeBox = $false

    $folderUrlLabel = New-Object System.Windows.Forms.Label
    $folderUrlLabel.Location = New-Object System.Drawing.Point(20, 30)
    $folderUrlLabel.Size = New-Object System.Drawing.Size(100, 20)
    $folderUrlLabel.Text = "Folder URL:"
    $folderForm.Controls.Add($folderUrlLabel)

    $folderUrlTextBox = New-Object System.Windows.Forms.TextBox
    $folderUrlTextBox.Location = New-Object System.Drawing.Point(130, 28)
    $folderUrlTextBox.Size = New-Object System.Drawing.Size(520, 20)
    $folderUrlTextBox.Text = "https://yourtenant.sharepoint.com/sites/sitename/Shared%20Documents/FolderName"
    $folderForm.Controls.Add($folderUrlTextBox)

    $helpLabel = New-Object System.Windows.Forms.Label
    $helpLabel.Location = New-Object System.Drawing.Point(130, 55)
    $helpLabel.Size = New-Object System.Drawing.Size(520, 40)
    $helpLabel.Text = "Tip: Copy the folder URL from SharePoint. Include subfolders if needed.`nExample: .../Shared Documents/Reports/2024/"
    $helpLabel.Font = New-Object System.Drawing.Font("Arial", 8)
    $helpLabel.ForeColor = [System.Drawing.Color]::Gray
    $folderForm.Controls.Add($helpLabel)

    $includeSubfoldersCheckBox = New-Object System.Windows.Forms.CheckBox
    $includeSubfoldersCheckBox.Location = New-Object System.Drawing.Point(130, 105)
    $includeSubfoldersCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $includeSubfoldersCheckBox.Text = "Include all subfolders (recursive)"
    $includeSubfoldersCheckBox.Checked = $true
    $folderForm.Controls.Add($includeSubfoldersCheckBox)

    $processButton = New-Object System.Windows.Forms.Button
    $processButton.Location = New-Object System.Drawing.Point(250, 150)
    $processButton.Size = New-Object System.Drawing.Size(100, 35)
    $processButton.Text = "Process"
    $processButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($folderUrlTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a folder URL.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $folderUrl = $folderUrlTextBox.Text.Trim()
        
        if (-not $folderUrl.StartsWith("http")) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid SharePoint URL starting with http or https.", "Invalid URL", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $folderForm.Hide()
        Start-FolderProcessingWithProgress -Mode $Mode -FolderUrl $folderUrl -IncludeSubfolders $includeSubfoldersCheckBox.Checked
        $folderForm.Close()
    })
    $folderForm.Controls.Add($processButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(370, 150)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $folderForm.Close()
    })
    $folderForm.Controls.Add($cancelButton)

    $folderForm.ShowDialog()
}


function Show-FolderScanOptions {
    $folderScanForm = New-Object System.Windows.Forms.Form
    $folderScanForm.Text = "Folder Scan Options"
    $folderScanForm.Size = New-Object System.Drawing.Size(500, 300)
    $folderScanForm.StartPosition = "CenterScreen"
    $folderScanForm.FormBorderStyle = "FixedDialog"
    $folderScanForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Text = "Select Folder Scan Type"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $folderScanForm.Controls.Add($titleLabel)

    $scanFolderButton = New-Object System.Windows.Forms.Button
    $scanFolderButton.Location = New-Object System.Drawing.Point(50, 70)
    $scanFolderButton.Size = New-Object System.Drawing.Size(400, 40)
    $scanFolderButton.Text = "Scan Specific Folder (Enter Folder URL)"
    $scanFolderButton.Add_Click({
        $folderScanForm.Hide()
        Show-FolderUrlDialog -Mode "calculate"
        $folderScanForm.Close()
    })
    $folderScanForm.Controls.Add($scanFolderButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(200, 210)
    $backButton.Size = New-Object System.Drawing.Size(100, 30)
    $backButton.Text = "Back"
    $backButton.Add_Click({
        $folderScanForm.Close()
    })
    $folderScanForm.Controls.Add($backButton)

    $folderScanForm.ShowDialog()
}

function Show-FolderCleanupOptions {
    $folderCleanupForm = New-Object System.Windows.Forms.Form
    $folderCleanupForm.Text = "Folder Cleanup Options"
    $folderCleanupForm.Size = New-Object System.Drawing.Size(500, 350)
    $folderCleanupForm.StartPosition = "CenterScreen"
    $folderCleanupForm.FormBorderStyle = "FixedDialog"
    $folderCleanupForm.MaximizeBox = $false

    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(20, 20)
    $titleLabel.Size = New-Object System.Drawing.Size(460, 30)
    $titleLabel.Text = "Select Folder Cleanup Type"
    $titleLabel.Font = New-Object System.Drawing.Font("Arial", 14, [System.Drawing.FontStyle]::Bold)
    $titleLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $folderCleanupForm.Controls.Add($titleLabel)

    $dryRunFolderButton = New-Object System.Windows.Forms.Button
    $dryRunFolderButton.Location = New-Object System.Drawing.Point(50, 70)
    $dryRunFolderButton.Size = New-Object System.Drawing.Size(400, 40)
    $dryRunFolderButton.Text = "Dry Run - Specific Folder (Simulate cleanup)"
    $dryRunFolderButton.BackColor = [System.Drawing.Color]::LightYellow
    $dryRunFolderButton.Add_Click({
        $folderCleanupForm.Hide()
        Show-FolderUrlDialog -Mode "dry"
        $folderCleanupForm.Close()
    })
    $folderCleanupForm.Controls.Add($dryRunFolderButton)

    $cleanupFolderButton = New-Object System.Windows.Forms.Button
    $cleanupFolderButton.Location = New-Object System.Drawing.Point(50, 120)
    $cleanupFolderButton.Size = New-Object System.Drawing.Size(400, 40)
    $cleanupFolderButton.Text = "CLEANUP - Specific Folder (DELETE VERSIONS)"
    $cleanupFolderButton.BackColor = [System.Drawing.Color]::Red
    $cleanupFolderButton.ForeColor = [System.Drawing.Color]::White
    $cleanupFolderButton.Add_Click({
        $result = [System.Windows.Forms.MessageBox]::Show(
            "This will PERMANENTLY DELETE old file versions from the specified folder. This action cannot be undone. Are you sure?",
            "Confirm Folder Cleanup",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            $folderCleanupForm.Hide()
            Show-FolderUrlDialog -Mode "run"
            $folderCleanupForm.Close()
        }
    })
    $folderCleanupForm.Controls.Add($cleanupFolderButton)

    $backButton = New-Object System.Windows.Forms.Button
    $backButton.Location = New-Object System.Drawing.Point(200, 180)
    $backButton.Size = New-Object System.Drawing.Size(100, 30)
    $backButton.Text = "Back"
    $backButton.Add_Click({
        $folderCleanupForm.Close()
    })
    $folderCleanupForm.Controls.Add($backButton)

    $folderCleanupForm.ShowDialog()
}

function Show-FolderUrlDialog {
    param(
        [string]$Mode
    )

    $folderForm = New-Object System.Windows.Forms.Form
    $folderForm.Text = "Folder URL Selection"
    $folderForm.Size = New-Object System.Drawing.Size(650, 350)
    $folderForm.StartPosition = "CenterScreen"
    $folderForm.FormBorderStyle = "FixedDialog"
    $folderForm.MaximizeBox = $false

    $instructionLabel = New-Object System.Windows.Forms.Label
    $instructionLabel.Location = New-Object System.Drawing.Point(20, 20)
    $instructionLabel.Size = New-Object System.Drawing.Size(600, 40)
    $instructionLabel.Text = "Enter the full SharePoint folder URL. Example:`nhttps://yourtenant.sharepoint.com/sites/sitename/Shared Documents/FolderName"
    $instructionLabel.Font = New-Object System.Drawing.Font("Arial", 9)
    $folderForm.Controls.Add($instructionLabel)

    $folderUrlLabel = New-Object System.Windows.Forms.Label
    $folderUrlLabel.Location = New-Object System.Drawing.Point(20, 80)
    $folderUrlLabel.Size = New-Object System.Drawing.Size(100, 20)
    $folderUrlLabel.Text = "Folder URL:"
    $folderForm.Controls.Add($folderUrlLabel)

    $folderUrlTextBox = New-Object System.Windows.Forms.TextBox
    $folderUrlTextBox.Location = New-Object System.Drawing.Point(20, 105)
    $folderUrlTextBox.Size = New-Object System.Drawing.Size(590, 20)
    $folderUrlTextBox.Text = "https://yourtenant.sharepoint.com/sites/sitename/Shared Documents/FolderName"
    $folderForm.Controls.Add($folderUrlTextBox)

    $includeSubfoldersCheckBox = New-Object System.Windows.Forms.CheckBox
    $includeSubfoldersCheckBox.Location = New-Object System.Drawing.Point(20, 140)
    $includeSubfoldersCheckBox.Size = New-Object System.Drawing.Size(300, 20)
    $includeSubfoldersCheckBox.Text = "Include all subfolders recursively"
    $includeSubfoldersCheckBox.Checked = $true
    $folderForm.Controls.Add($includeSubfoldersCheckBox)

    $helpLabel = New-Object System.Windows.Forms.Label
    $helpLabel.Location = New-Object System.Drawing.Point(20, 170)
    $helpLabel.Size = New-Object System.Drawing.Size(590, 60)
    $helpLabel.Text = "URL Format Help:`n• Navigate to your folder in SharePoint`n• Copy the URL from the browser address bar`n• Make sure it includes the full path to the specific folder"
    $helpLabel.Font = New-Object System.Drawing.Font("Arial", 8)
    $helpLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    $folderForm.Controls.Add($helpLabel)

    $processButton = New-Object System.Windows.Forms.Button
    $processButton.Location = New-Object System.Drawing.Point(200, 250)
    $processButton.Size = New-Object System.Drawing.Size(100, 35)
    $processButton.Text = "Process"
    $processButton.Add_Click({
        if ([string]::IsNullOrWhiteSpace($folderUrlTextBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a folder URL.", "Missing Information", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $folderUrl = $folderUrlTextBox.Text.Trim()
        
        if (-not $folderUrl.StartsWith("http")) {
            [System.Windows.Forms.MessageBox]::Show("Please enter a valid SharePoint URL starting with http or https.", "Invalid URL", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
            return
        }
        
        $folderForm.Hide()
        Start-FolderProcessingWithProgress -Mode $Mode -FolderUrl $folderUrl -IncludeSubfolders $includeSubfoldersCheckBox.Checked
        $folderForm.Close()
    })
    $folderForm.Controls.Add($processButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(320, 250)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 35)
    $cancelButton.Text = "Cancel"
    $cancelButton.Add_Click({
        $folderForm.Close()
    })
    $folderForm.Controls.Add($cancelButton)

    $folderForm.ShowDialog()
}

function Start-FolderProcessingWithProgress {
    param(
        [string]$Mode,
        [string]$FolderUrl,
        [bool]$IncludeSubfolders
    )

    $script:progressForm = New-Object System.Windows.Forms.Form
    $script:progressForm.Text = "Processing SharePoint Folder"
    $script:progressForm.Size = New-Object System.Drawing.Size(800, 600)
    $script:progressForm.StartPosition = "CenterScreen"
    $script:progressForm.FormBorderStyle = "FixedDialog"
    $script:progressForm.MaximizeBox = $false

    # Create form controls
    $progressLabel = New-Object System.Windows.Forms.Label
    $progressLabel.Location = New-Object System.Drawing.Point(20, 20)
    $progressLabel.Size = New-Object System.Drawing.Size(760, 20)
    $progressLabel.Text = "Folder Processing Progress:"
    $script:progressForm.Controls.Add($progressLabel)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.Location = New-Object System.Drawing.Point(20, 45)
    $progressBar.Size = New-Object System.Drawing.Size(600, 25)
    $script:progressForm.Controls.Add($progressBar)

    $etaLabel = New-Object System.Windows.Forms.Label
    $etaLabel.Location = New-Object System.Drawing.Point(630, 48)
    $etaLabel.Size = New-Object System.Drawing.Size(150, 20)
    $etaLabel.Text = "ETA: Calculating..."
    $script:progressForm.Controls.Add($etaLabel)

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.Location = New-Object System.Drawing.Point(20, 90)
    $statusLabel.Size = New-Object System.Drawing.Size(760, 40)
    $statusLabel.Text = "Initializing folder processing..."
    $statusLabel.Font = New-Object System.Drawing.Font("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $script:progressForm.Controls.Add($statusLabel)

    $script:logTextBox = New-Object System.Windows.Forms.TextBox
    $script:logTextBox.Location = New-Object System.Drawing.Point(20, 140)
    $script:logTextBox.Size = New-Object System.Drawing.Size(760, 360)
    $script:logTextBox.Multiline = $true
    $script:logTextBox.ScrollBars = "Vertical"
    $script:logTextBox.ReadOnly = $true
    $script:logTextBox.Font = New-Object System.Drawing.Font("Consolas", 9)
    $script:progressForm.Controls.Add($script:logTextBox)

    $closeButton = New-Object System.Windows.Forms.Button
    $closeButton.Location = New-Object System.Drawing.Point(350, 520)
    $closeButton.Size = New-Object System.Drawing.Size(100, 30)
    $closeButton.Text = "Close"
    $closeButton.Enabled = $false
    $closeButton.Add_Click({
        $script:progressForm.Close()
    })
    $script:progressForm.Controls.Add($closeButton)

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Location = New-Object System.Drawing.Point(460, 520)
    $cancelButton.Size = New-Object System.Drawing.Size(100, 30)
    $cancelButton.Text = "Cancel"
    $script:progressForm.Controls.Add($cancelButton)

    # Variables to store job and timer
    $script:backgroundJob = $null
    $script:jobTimer = $null
    $script:isClosing = $false

    $cancelButton.Add_Click({
        if ($script:backgroundJob) {
            $script:cancelRequested = $true
            $cancelButton.Enabled = $false
            $cancelButton.Text = "Cancelling..."
            
            try {
                Stop-Job -Job $script:backgroundJob -ErrorAction SilentlyContinue
                Remove-Job -Job $script:backgroundJob -Force -ErrorAction SilentlyContinue
            } catch {}

            if ($script:jobTimer) {
                $script:jobTimer.Stop()
                $script:jobTimer.Dispose()
            }
            
            $closeButton.Enabled = $true
        }
    })

    $script:progressForm.Add_FormClosing({
        param($sender, $e)
        
        if (-not $script:isClosing -and -not $closeButton.Enabled) {
            $e.Cancel = $true
            return
        }
        
        $script:isClosing = $true
        
        if ($script:jobTimer) {
            $script:jobTimer.Stop()
            $script:jobTimer.Dispose()
        }
        
        if ($script:backgroundJob) {
            try {
                Stop-Job -Job $script:backgroundJob -ErrorAction SilentlyContinue
                Remove-Job -Job $script:backgroundJob -Force -ErrorAction SilentlyContinue
            } catch {}
        }
    })

    $script:progressForm.Add_Shown({
        try {
            $uri = [System.Uri]::new($FolderUrl)
            $sitePath = $uri.AbsolutePath
            $siteUrl = $uri.Scheme + "://" + $uri.Host

            if ($sitePath -match "^(/sites/[^/]+|/teams/[^/]+)") {
                $siteUrl += $matches[1]
            }

            Write-Log "Connecting to SharePoint site: $siteUrl" "INFO"
            $statusLabel.Text = "Connecting to SharePoint site..."
            
            try {
                Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $script:clientId -ErrorAction Stop
                Write-Log "Connected to SharePoint site successfully" "SUCCESS"
            } catch {
                $errorMessage = "Failed to connect to SharePoint site: $($_.Exception.Message)"
                Write-Log $errorMessage "ERROR"
                [System.Windows.Forms.MessageBox]::Show(
                    $errorMessage,
                    "Connection Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
                $closeButton.Enabled = $true
                return
            }
            
            $script:backgroundJob = Start-Job -ScriptBlock {
                param($Mode, $FolderUrl, $IncludeSubfolders, $KeepVersions, $LogFile, $ClientId)
                
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
                
                function Get-AllFiles {
                    param(
                        [string]$FolderUrl,
                        [bool]$Recursive
                    )
                    
                    $allFiles = @()
                    
                    $officeExtensions = @(
                        "docx", "doc",     # Word
                        "xlsx", "xls",     # Excel
                        "pptx", "ppt",     # PowerPoint
                        "vsdx", "vsd",     # Visio
                        "one",             # OneNote
                        "pdf",
                        "csv",
                        "txt"             # PDF (not strictly Office but often versioned)
                    )
                    
                    try {
                        $uri = [System.Uri]::new($FolderUrl)
                        $fullPath = [System.Web.HttpUtility]::UrlDecode($uri.AbsolutePath)
                        $fullPath = $fullPath.TrimEnd('/')
                        
                        $siteUrl = $uri.Scheme + "://" + $uri.Host
                        $sitePath = ""
                        $relativePath = $fullPath

                        if ($fullPath -match "^(/sites/[^/]+|/teams/[^/]+)") {
                            $sitePath = $matches[1]
                            $siteUrl += $sitePath
                            $relativePath = $fullPath.Substring($sitePath.Length)
                        }

                        Write-JobLog "Processing URL components:" "INFO"
                        Write-JobLog "Site URL: $siteUrl" "INFO"
                        Write-JobLog "Site Path: $sitePath" "INFO"
                        Write-JobLog "Relative Path: $relativePath" "INFO"

                        # Connect to the site
                        Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId
                        Write-JobLog "Connected to site for file processing" "SUCCESS"

                        $pathParts = $relativePath.Split('/', [System.StringSplitOptions]::RemoveEmptyEntries)
                        $libraryName = $pathParts[0]
                        if ($libraryName -eq "Shared Documents") {
                            $libraryName = "Documents"  # Internal name is usually just "Documents"
                        }
                        Write-JobLog "Using library name: $libraryName" "INFO"

                        function Process-Folder {
                            param(
                                [string]$FolderPath
                            )
                            
                            Write-JobLog "Processing folder: $FolderPath" "INFO"
                            $folderFiles = @()
                            
                            try {
                                $items = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderPath -ItemType File
                                Write-JobLog "Found $($items.Count) files in folder $FolderPath" "INFO"
                                
                                foreach ($file in $items) {
                                    try {
                                        $extension = [System.IO.Path]::GetExtension($file.Name).TrimStart('.')
                                        if ($officeExtensions -contains $extension.ToLower()) {
                                            $listItem = Get-PnPFile -Url $file.ServerRelativeUrl -AsListItem
                                            if ($listItem -and $listItem.File) {
                                                $fileObj = Get-PnPProperty -ClientObject $listItem.File -Property Versions
                                                if ($fileObj) {
                                                    $folderFiles += @{
                                                        ListItem = $listItem
                                                        File = $fileObj
                                                        Name = $file.Name
                                                        ServerRelativeUrl = $file.ServerRelativeUrl
                                                    }
                                                    Write-JobLog "Added file: $($file.Name)" "INFO"
                                                }
                                            }
                                        }
                                    } catch {
                                        Write-JobLog "Error processing file $($file.Name): $($_.Exception.Message)" "ERROR"
                                        continue
                                    }
                                }
                                
                                if ($Recursive) {
                                    try {
                                        $subFolders = Get-PnPFolderItem -FolderSiteRelativeUrl $FolderPath -ItemType Folder
                                        Write-JobLog "Found $($subFolders.Count) subfolders in $FolderPath" "INFO"
                                        
                                        foreach ($subFolder in $subFolders) {
                                            Write-JobLog "Processing subfolder: $($subFolder.Name)" "INFO"
                                            $subFolderFiles = Process-Folder -FolderPath $subFolder.ServerRelativeUrl
                                            if ($subFolderFiles -and $subFolderFiles.Count -gt 0) {
                                                $folderFiles += $subFolderFiles
                                                Write-JobLog "Added $($subFolderFiles.Count) files from subfolder $($subFolder.Name)" "INFO"
                                            }
                                        }
                                    } catch {
                                        Write-JobLog "Error getting subfolders for $FolderPath $($_.Exception.Message)" "ERROR"
                                    }
                                }
                            } catch {
                                Write-JobLog "Error processing folder $FolderPath $($_.Exception.Message)" "ERROR"
                            }
                            
                            Write-JobLog "Total files found in $FolderPath $($folderFiles.Count)" "INFO"
                            return $folderFiles
                        }

                        $allFiles = Process-Folder -FolderPath $relativePath
                        
                        Write-JobLog "Total Office files found: $($allFiles.Count)" "INFO"
                        return $allFiles
                        
                    } catch {
                        Write-JobLog "Error in Get-AllFiles: $($_.Exception.Message)" "ERROR"
                        Write-JobLog "Stack Trace: $($_.Exception.StackTrace)" "ERROR"
                        return @()
                    }
                }

                try {
                    Write-JobLog "Starting folder processing in $Mode mode" "INFO"
                    
                    $totalStats = @{
                        Files = 0
                        ExcessVersions = 0
                        ExcessSize = 0
                        VersionsDeleted = 0
                        Errors = 0
                    }

                    $files = Get-AllFiles -FolderUrl $FolderUrl -Recursive $IncludeSubfolders
                    $totalStats.Files = $files.Count
                    Write-JobLog "Total files found: $($files.Count)" "INFO"

                    Write-Output @{Type="Progress"; Total=$files.Count; Current=0}
                    $processedFiles = 0

                    foreach ($file in $files) {
                        $processedFiles++
                        Write-Output @{Type="Status"; Message="Processing file: $($file.Name) ($processedFiles of $($files.Count))"}
                        Write-Output @{Type="Progress"; Total=$files.Count; Current=$processedFiles}

                        try {
                            if ($file.File.Versions) {
                                $versionCount = $file.File.Versions.Count
                                
                                if ($versionCount -gt $KeepVersions) {
                                    $versionsToDelete = $file.File.Versions | Sort-Object -Property Created -Descending | Select-Object -Skip $KeepVersions
                                    $deleteCount = $versionsToDelete.Count
                                    $totalStats.ExcessVersions += $deleteCount
                                    
                                    $excessSize = ($versionsToDelete | Measure-Object -Property Size -Sum).Sum
                                    $totalStats.ExcessSize += $excessSize
                                    
                                    $formattedSize = Format-JobFileSize -Size $excessSize
                                    Write-JobLog "File: $($file.Name) - $deleteCount excess versions using $formattedSize" "INFO"
                                    
                                    if ($Mode -eq "run") {
                                        foreach ($version in $versionsToDelete) {
                                            try {
                                                $version.DeleteObject()
                                                $totalStats.VersionsDeleted++
                                            } catch {
                                                Write-JobLog "Error deleting version for file $($file.Name): $($_.Exception.Message)" "ERROR"
                                                $totalStats.Errors++
                                            }
                                        }
                                        Invoke-PnPQuery
                                        Write-JobLog "Deleted $deleteCount versions from $($file.Name)" "SUCCESS"
                                    }
                                } else {
                                    Write-JobLog "File: $($file.Name) - No excess versions found (current: $versionCount)" "INFO"
                                }
                            } else {
                                Write-JobLog "File: $($file.Name) - No versions found" "INFO"
                            }
                        } catch {
                            $totalStats.Errors++
                            Write-JobLog "Error processing file $($file.Name): $($_.Exception.Message)" "ERROR"
                        }
                    }

                    Write-Output @{Type="Complete"; Stats=$totalStats}
                    
                } catch {
                    Write-JobLog "Critical error: $($_.Exception.Message)" "ERROR"
                    Write-Output @{Type="Error"; Message=$_.Exception.Message}
                }
            } -ArgumentList $Mode, $FolderUrl, $IncludeSubfolders, $script:keepVersions, $script:logFile, $script:clientId

            $script:jobTimer = New-Object System.Windows.Forms.Timer
            $script:jobTimer.Interval = 500
            $script:jobTimer.Add_Tick({
                if ($script:cancelRequested -or $script:isClosing) {
                    $script:jobTimer.Stop()
                    $closeButton.Enabled = $true
                    return
                }
                
                if ($script:backgroundJob) {
                    $output = Receive-Job -Job $script:backgroundJob
                    
                    foreach ($item in $output) {
                        switch ($item.Type) {
                            "Status" {
                                $statusLabel.Text = $item.Message
                            }
                            "Progress" {
                                $progressBar.Maximum = $item.Total
                                $progressBar.Value = [Math]::Min($item.Current, $item.Total)
                                
                                if ($item.Current -gt 0) {
                                    $percent = [Math]::Round(($item.Current / $item.Total) * 100)
                                    $statusLabel.Text = "$($item.Message) - $percent% Complete"
                                }
                            }
                            "Log" {
                                if ($script:logTextBox) {
                                    $script:logTextBox.AppendText("$($item.Message)`r`n")
                                    $script:logTextBox.ScrollToCaret()
                                }
                            }
                            "Complete" {
                                $stats = $item.Stats
                                $formattedSize = Format-FileSize -Size $stats.ExcessSize
                                
                                $summaryMessage = @"
========== FOLDER PROCESSING COMPLETE ==========
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
                                    "Folder Processing Complete",
                                    [System.Windows.Forms.MessageBoxButtons]::OK,
                                    [System.Windows.Forms.MessageBoxIcon]::Information
                                )
                                
                                $closeButton.Enabled = $true
                                $cancelButton.Enabled = $false
                                return
                            }
                            "Error" {
                                [System.Windows.Forms.MessageBox]::Show(
                                    "Error occurred: $($item.Message)",
                                    "Error",
                                    [System.Windows.Forms.MessageBoxButtons]::OK,
                                    [System.Windows.Forms.MessageBoxIcon]::Error
                                )
                                $closeButton.Enabled = $true
                                $cancelButton.Enabled = $false
                                return
                            }
                        }
                    }
                    
                    if ($script:backgroundJob.State -ne "Running") {
                        $script:jobTimer.Stop()
                        Remove-Job -Job $script:backgroundJob -Force
                        $closeButton.Enabled = $true
                        $cancelButton.Enabled = $false
                    }
                }
            })
            $script:jobTimer.Start()
            
        } catch {
            $errorMessage = "Error initializing folder processing: $($_.Exception.Message)"
            Write-Log $errorMessage "ERROR"
            [System.Windows.Forms.MessageBox]::Show(
                $errorMessage,
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
            $closeButton.Enabled = $true
        }
    })

    $script:progressForm.ShowDialog()
}

function Start-ProcessingWithProgress {
    param(
        [string]$Mode,
        [bool]$ProcessAll,
        [string]$SiteUrl = "",
        [string]$LibraryName = "",
        [string]$FolderUrl = "",
        [bool]$IncludeSubfolders = $true
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
            } elseif ($FolderUrl) {
                Write-Output @{Type="Status"; Message="Processing specific folder..."}

                try {
                    $uri = [System.Uri]$FolderUrl
                    $sitePath = $uri.AbsolutePath
                    $siteUrl = $uri.Scheme + "://" + $uri.Host

                    if ($sitePath -match "^(/sites/[^/]+|/teams/[^/]+)") {
                        $siteUrl += $matches[1]
                    } else {
                        throw "Invalid folder URL format. Must contain /sites/ or /teams/ in the path."
                    }

                    Write-JobLog "Extracted site URL: $siteUrl" "INFO"
                    Connect-PnPOnline -Url $siteUrl -Interactive -ClientId $ClientId
                    Write-JobLog "Connected to site for folder processing" "SUCCESS"

                    $folderPath = $uri.AbsolutePath.Replace("%20", " ")
                    $relativeRoot = $siteUrl.Replace($uri.Scheme + "://" + $uri.Host, "")
                    $folderPath = $folderPath.Substring($relativeRoot.Length)

                    Write-JobLog "Final folder path: $folderPath" "DEBUG"

                    $totalStats = @{
                        Libraries = 1
                        Files = 0
                        ExcessVersions = 0
                        ExcessSize = 0
                        VersionsDeleted = 0
                        Errors = 0
                    }

                    Write-Output @{Type="Progress"; Overall=@{Total=1; Current=0}}
                    Write-Output @{Type="Status"; Message="Scanning folder for files..."}

                    if ($IncludeSubfolders) {
                        $camlQuery = "<View Scope='RecursiveAll'><Query><Where><BeginsWith><FieldRef Name='FileDirRef'/><Value Type='Text'>$folderPath</Value></BeginsWith></Where></Query></View>"
                        $items = Get-PnPListItem -List "Documents" -Query $camlQuery | Where-Object { $_.FileSystemObjectType -eq "File" }
                    } else {
                        $folder = Get-PnPFolder -Url $folderPath
                        $items = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType File
                        $items = $items | ForEach-Object { Get-PnPFile -Url $_.ServerRelativeUrl -AsListItem }
                    }

                    $totalStats.Files = $items.Count
                    Write-JobLog "Found $($items.Count) files in folder" "INFO"

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

                    Write-Output @{Type="Progress"; Overall=@{Total=1; Current=1}}
                    Write-Output @{Type="Complete"; Stats=$totalStats}

                } catch {
                    Write-JobLog "Error processing folder $FolderUrl : $($_.Exception.Message)" "ERROR"
                    Write-Output @{Type="Error"; Message=$_.Exception.Message}
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