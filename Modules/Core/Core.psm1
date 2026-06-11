<# 
.SYNOPSIS
    Core utility functions for SharePoint Version Cleaner
    Includes logging, file formatting, and session management
#>

function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        
        [Parameter(Mandatory=$false)]
        [ValidateSet("DEBUG", "INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO",
        
        [Parameter(Mandatory=$false)]
        [switch]$SkipUI,
        
        [Parameter(Mandatory=$false)]
        [switch]$SkipConsole
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file - check both script and global scope
    $logFilePath = $script:logFile
    if (-not $logFilePath) {
        $logFilePath = $global:logFile
    }
    
    if ($logFilePath) {
        try {
            Add-Content -Path $logFilePath -Value $logMessage
        } catch {
            # Silently fail if log file is locked or unavailable
        }
    }
    
    # Write to console unless explicitly skipped
    if (-not $SkipConsole) {
        switch ($Level) {
            "DEBUG" {
                Write-Host $logMessage -ForegroundColor Gray
            }
            "ERROR" {
                Write-Host $logMessage -ForegroundColor Red
            }
            "WARNING" {
                Write-Host $logMessage -ForegroundColor Yellow
            }
            "SUCCESS" {
                Write-Host $logMessage -ForegroundColor Green
            }
            default {
                Write-Host $logMessage -ForegroundColor White
            }
        }
    }
    
    # Only update UI if we're on the main thread and not skipping UI
    if (-not $SkipUI -and $script:progressForm -and $script:logTextBox) {
        try {
            # Check if we're on the UI thread
            if ($script:logTextBox.InvokeRequired) {
                return
            }
            
            $script:logTextBox.AppendText("$logMessage`r`n")
            $script:logTextBox.ScrollToCaret()
        } catch {
            # Ignore UI update errors to prevent infinite loops
        }
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

function Enable-ReportButton {
    if ($script:reportButton -and -not $script:reportButton.IsDisposed) {
        try {
            $script:reportButton.Enabled = $true
            $script:reportButton.BackColor = [System.Drawing.Color]::LightGreen
            $script:reportButton.ForeColor = [System.Drawing.Color]::Black
        } catch {
            # Ignore UI update errors
        }
    }
}

function Ensure-PnPModule {
    param (
        [string]$ModuleName = "PnP.PowerShell",
        [switch]$SkipUpdate
    )

    $module = Get-Module -ListAvailable -Name $ModuleName
    
    if ($module) {
        # Only check for updates if not explicitly skipped and not already checked this session
        if (-not $SkipUpdate -and -not $global:modulesCheckedForUpdates) {
            Write-Host "--- Checking for updates for $ModuleName ---" -ForegroundColor Cyan
            Write-Log "Checking for updates for $ModuleName..." "INFO"
            
            try {
                Update-Module -Name $ModuleName -ErrorAction SilentlyContinue
                Write-Host "Module is up to date or has been updated." -ForegroundColor Green
            } catch {
                Write-Host "Unable to check for updates (check your internet connection)." -ForegroundColor Yellow
            }
            
            # Mark that we've checked for updates this session
            $global:modulesCheckedForUpdates = $true
        }
    } else {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "$ModuleName is not installed. Would you like to install it now?",
            "Required Module",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            try {
                Write-Host "Installing $ModuleName... Please wait." -ForegroundColor Yellow
                Install-Module $ModuleName -AllowClobber -Force -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show("$ModuleName has been installed successfully.", "Success")
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Installation failed: $_", "Error")
                exit 1
            }
        } else {
            exit 1
        }
    }
}

function Initialize-SessionState {
    # Configuration variables (with defaults) - use $global: for cross-module access
    $global:clientId = "your-entraid-client-app-id"
    $global:adminUrl = "https://your-tenant-admin.sharepoint.com/"
    $global:defaultSiteUrl = "https://tenant.sharepoint.com/sites/sitename"
    $global:keepVersions = 25
    $global:logDirectory = "Logs"
    
    # Module update check flag (to avoid checking on every startup)
    $global:modulesCheckedForUpdates = $false
    
    # Global authentication token and connection management
    $global:globalAccessToken = $null
    $global:globalSharePointToken = $null
    $global:globalTokenExpiry = $null
    $global:isGloballyAuthenticated = $false

    # Cache for connections and sites
    $global:connectedSites = @{}
    $global:cachedSites = @()
    $global:lastSitesFetch = $null

    # Session statistics for report generation
    $global:sessionStats = @{
        SessionStart = Get-Date
        TotalOperations = 0
        ScanOperations = @()
        CleanupOperations = @()
        AffectedSites = @{}
        TotalStorageSaved = 0
        TotalVersionsDeleted = 0
        TotalFilesProcessed = 0
        TotalLibrariesProcessed = 0
        TotalSitesProcessed = 0
        TotalErrors = 0
        HasPerformedOperations = $false
    }

    # Processing state
    $global:cancelRequested = $false
    $global:currentBackgroundJob = $null
    $global:progressForm = $null
    $global:logTextBox = $null
    $global:reportButton = $null
    
    Write-Log "Session state initialized" "INFO"
}

# Export public functions
Export-ModuleMember -Function @(
    'Write-Log',
    'Format-FileSize',
    'Enable-ReportButton',
    'Ensure-PnPModule',
    'Initialize-SessionState'
)
