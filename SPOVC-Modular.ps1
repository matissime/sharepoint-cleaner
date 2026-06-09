#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Version Cleaner (SPOVC) - Modular Edition
.DESCRIPTION
    A modular PowerShell 7+ GUI application for managing file version cleanup in SharePoint Online.
    Provides scan, dry-run, and cleanup modes with HTML session reporting and comprehensive logging.
.AUTHOR
    SharePoint Administration Team
.VERSION
    2.0.0 (Modular)
.NOTES
    - Requires PowerShell 7+
    - Requires PnP.PowerShell module (installed automatically on first run)
    - Configuration stored in config.json
    - Logs and reports stored in Logs/ directory
#>

# Enforce PowerShell 7+
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Host "This script requires PowerShell 7 or later." -ForegroundColor Red
    Write-Host "Current version: $($PSVersionTable.PSVersion)" -ForegroundColor Yellow
    Write-Host "Download PowerShell 7+ from: https://github.com/PowerShell/PowerShell/releases" -ForegroundColor Cyan
    exit 1
}

# Add WinForms assembly
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# Script location
$scriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path

# Import all modules in dependency order
try {
    Write-Host "Loading modules..." -ForegroundColor Cyan
    
    Import-Module "$scriptPath/Modules/Core/Core.psm1" -Force
    Import-Module "$scriptPath/Modules/Core/Config.psm1" -Force
    Import-Module "$scriptPath/Modules/Authentication/Authentication.psm1" -Force
    Import-Module "$scriptPath/Modules/Reporting/Reporting.psm1" -Force
    Import-Module "$scriptPath/Modules/UI/UI.psm1" -Force
    Import-Module "$scriptPath/Modules/Processing/SharePointOps.psm1" -Force
    Import-Module "$scriptPath/Modules/Processing/Processing.psm1" -Force
    
    Write-Host "Modules loaded successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error loading modules: $_" -ForegroundColor Red
    exit 1
}

# Initialize session state first (sets defaults)
try {
    Write-Host "Initializing session..." -ForegroundColor Cyan
    Initialize-SessionState
    Ensure-PnPModule -SkipUpdate  # Skip update check on startup for faster loading
    Write-Host "Session initialized!" -ForegroundColor Green
} catch {
    Write-Host "Error initializing session: $_" -ForegroundColor Red
    exit 1
}

# Load configuration
try {
    Write-Host "Loading configuration..." -ForegroundColor Cyan
    
    $configPath = Join-Path $scriptPath "config.json"
    
    # If config doesn't exist, create it
    if (-not (Test-Path $configPath)) {
        Write-Host "Configuration file not found: $configPath" -ForegroundColor Yellow
        Write-Host "Creating default config.json..." -ForegroundColor Yellow
        
        $defaultConfig = @{
            settings = @{
                clientId = "your-entraid-client-app-id"
                adminUrl = "https://your-tenant-admin.sharepoint.com/"
                keepVersions = 25
            }
            logging = @{
                directory = "Logs"
                enableFileLog = $true
                enableConsoleLog = $true
            }
            inactiveSites = @{
                defaultDaysInactive = 90
                defaultMaxStorageMB = 100
                includeEmptySites = $false
            }
        }
        
        $defaultConfig | ConvertTo-Json | Set-Content $configPath -Force
        Write-Host "Created default config.json at: $configPath" -ForegroundColor Yellow
        Write-Host "⚠️  Please edit config.json with your Entra app ID and admin URL" -ForegroundColor Yellow
    }
    
    # Load and apply configuration
    $config = Load-Configuration -ConfigPath $configPath
    Apply-Configuration -Config $config
    
    Write-Host "Configuration loaded successfully!" -ForegroundColor Green
} catch {
    Write-Host "Error loading configuration: $_" -ForegroundColor Red
    exit 1
}

# Session state and PnP already initialized above

# Global variable to store selected sites for processing
$global:selectedSitesForProcessing = @()

# Main execution
function Main {
    try {
        # Initialize authentication
        Write-Host "Initializing authentication..." -ForegroundColor Cyan
        Initialize-GlobalAuthentication
        Write-Host "Authentication successful!" -ForegroundColor Green
        
        # Load and cache sites at startup
        Write-Host "Loading SharePoint sites in background..." -ForegroundColor Cyan
        try {
            $global:cachedAllSites = Get-PnPTenantSite -Detailed | Where-Object { $_.Status -eq "Active" } | Sort-Object Title
            Write-Host "Cached $($global:cachedAllSites.Count) active sites" -ForegroundColor Green
        } catch {
            Write-Host "Warning: Could not load all sites at startup: $_" -ForegroundColor Yellow
            $global:cachedAllSites = @()
        }
        
        # Show main menu loop
        while ($true) {
            $operation = Show-MainMenu
            
            if ($null -eq $operation -or $operation -eq "EXIT") {
                # User cancelled or clicked Exit
                Write-Host "Exiting..." -ForegroundColor Yellow
                break
            }
            
            # Handle Disconnect operation
            if ($operation -eq "DISCONNECT") {
                Write-Host "Clearing authentication cache..." -ForegroundColor Cyan
                Clear-AuthenticationCache
                [System.Windows.Forms.MessageBox]::Show("Authentication tokens cleared. You will be prompted to authenticate on your next operation.", "Disconnected", "OK", "Information")
                continue
            }
            
            # Handle Select Sites operation
            if ($operation -eq "SELECT_SITES") {
                Write-Host "Opening site selector..." -ForegroundColor Cyan
                $selectedSites = Show-SitesBrowserDialog
                
                if ($null -ne $selectedSites -and $selectedSites.Count -gt 0) {
                    $global:selectedSitesForProcessing = $selectedSites
                    Write-Host "Selected $($selectedSites.Count) site(s) for processing" -ForegroundColor Green
                    [System.Windows.Forms.MessageBox]::Show("Selected $($selectedSites.Count) site(s):`n`n$($selectedSites | ForEach-Object {$_.Title} | Join-String -Separator "`n")", "Sites Selected", "OK", "Information")
                } else {
                    Write-Host "No sites selected." -ForegroundColor Yellow
                }
                continue
            }
            
            # Check if sites are selected for SCAN/DRYRUN/CLEANUP
            if ($operation -in @("SCAN", "DRYRUN", "CLEANUP")) {
                if ($global:selectedSitesForProcessing.Count -eq 0) {
                    $msg = "No sites selected for processing.`n`nPlease click 'Select Sites' on the main menu first."
                    [System.Windows.Forms.MessageBox]::Show($msg, "No Sites Selected", "OK", "Warning")
                    continue
                }
                
                $selectedSites = $global:selectedSitesForProcessing
            }
            
            # SAFETY CHECK: Double confirmation for CLEANUP mode
            if ($operation -eq "CLEANUP") {
                $confirmMsg = @"
⚠️  DESTRUCTIVE OPERATION WARNING ⚠️

You are about to PERMANENTLY DELETE excess file versions from SharePoint.

This action:
✓ Will PERMANENTLY DELETE file versions
✓ CANNOT be undone
✓ Will FREE storage space
✓ KEEPS $($global:keepVersions) most recent versions per file

Are you absolutely sure you want to proceed?
"@
                $result1 = [System.Windows.Forms.MessageBox]::Show($confirmMsg, "Confirm Cleanup Operation", [System.Windows.Forms.MessageBoxButtons]::YesNo, [System.Windows.Forms.MessageBoxIcon]::Warning)
                
                if ($result1 -ne [System.Windows.Forms.DialogResult]::Yes) {
                    Write-Host "Cleanup operation cancelled by user." -ForegroundColor Yellow
                    continue
                }

                # SECOND confirmation
                $confirmMsg2 = @"
FINAL CONFIRMATION REQUIRED

Type 'YES' below to confirm you want to permanently delete excess versions:

(This is the FINAL step - you cannot undo this operation)
"@
                $result2 = [System.Windows.Forms.MessageBox]::Show($confirmMsg2, "FINAL CONFIRMATION", [System.Windows.Forms.MessageBoxButtons]::OKCancel, [System.Windows.Forms.MessageBoxIcon]::Stop)
                
                if ($result2 -ne [System.Windows.Forms.DialogResult]::OK) {
                    Write-Host "Cleanup operation cancelled by user at final confirmation." -ForegroundColor Yellow
                    continue
                }
            }
            
            # Execute operation with selected sites
            Write-Host "Starting $operation operation on $($selectedSites.Count) site(s)..." -ForegroundColor Cyan
            Start-ProcessingWithProgress -Mode $operation -Scope "SelectedSites" -SelectedSites $selectedSites
            
            # Ask if user wants to continue
            $continueResponse = [System.Windows.Forms.MessageBox]::Show(
                "Operation complete. Return to main menu?",
                "SPOVC",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question
            )
            
            if ($continueResponse -ne [System.Windows.Forms.DialogResult]::Yes) {
                break
            }
        }
        
        Write-Host "`nThank you for using SharePoint Version Cleaner!" -ForegroundColor Cyan
        Write-Log "SharePoint Version Cleaner closed by user" "INFO"
        
    } catch {
        Write-Host "Error in main execution: $_" -ForegroundColor Red
        Write-Log "Fatal error: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Fatal error: $_", "Error", "OK", "Error")
    }
}

# Run main
Main
