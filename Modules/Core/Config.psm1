<# 
.SYNOPSIS
    Configuration management for SharePoint Version Cleaner
    Handles loading and parsing config.json
#>

function Load-Configuration {
    param(
        [Parameter(Mandatory=$false)]
        [string]$ConfigPath
    )
    
    # Resolve config path - use provided path or search for it
    if ([string]::IsNullOrEmpty($ConfigPath)) {
        # Try to find config.json in common locations
        $searchPaths = @(
            "$PSScriptRoot\..\..\config.json",
            "$PSScriptRoot\config.json",
            ".\config.json",
            "config.json"
        )
        
        foreach ($path in $searchPaths) {
            if (Test-Path $path) {
                $ConfigPath = $path
                break
            }
        }
    }
    
    try {
        if (-not (Test-Path $ConfigPath)) {
            Write-Log "Configuration file not found at $ConfigPath, using defaults" "WARNING"
            return $null
        }

        # Resolve to absolute path
        $resolvedPath = (Resolve-Path $ConfigPath).Path
        $config = Get-Content $resolvedPath -Raw | ConvertFrom-Json
        Write-Log "Configuration loaded from $resolvedPath" "INFO"
        return $config
    } catch {
        Write-Log "Error loading configuration: $($_.Exception.Message)" "ERROR"
        return $null
    }
}

function Get-ConfigValue {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Config,
        
        [Parameter(Mandatory=$true)]
        [string]$Key,
        
        [object]$DefaultValue
    )
    
    if ($null -eq $Config) {
        return $DefaultValue
    }
    
    $keys = $Key -split '\.'
    $value = $Config
    
    foreach ($k in $keys) {
        if ($value -is [PSObject] -and $value.PSObject.Properties.Name -contains $k) {
            $value = $value.$k
        } else {
            return $DefaultValue
        }
    }
    
    return $value
}

function Apply-Configuration {
    param(
        [Parameter(Mandatory=$true)]
        [object]$Config
    )
    
    if ($null -eq $Config) {
        Write-Log "No configuration file loaded, using defaults" "INFO"
        # Ensure global variables are set to defaults (already done in Initialize-SessionState)
        return
    }
    
    # Apply settings from config, falling back to current global defaults
    $clientId = Get-ConfigValue -Config $Config -Key "settings.clientId"
    if ($clientId) { $global:clientId = $clientId }
    
    $adminUrl = Get-ConfigValue -Config $Config -Key "settings.adminUrl"
    if ($adminUrl) { $global:adminUrl = $adminUrl }
    
    $keepVersions = Get-ConfigValue -Config $Config -Key "settings.keepVersions"
    if ($keepVersions) { $global:keepVersions = $keepVersions }
    
    $defaultSiteUrl = Get-ConfigValue -Config $Config -Key "settings.defaultSiteUrl"
    if ($defaultSiteUrl) { $global:defaultSiteUrl = $defaultSiteUrl }
    
    # Apply logging settings
    $logDir = Get-ConfigValue -Config $Config -Key "logging.directory" -DefaultValue "Logs"
    if ($logDir) {
        # Resolve log directory path (relative to script root or absolute)
        if ([System.IO.Path]::IsPathRooted($logDir)) {
            $global:logDirectory = $logDir
        } else {
            $global:logDirectory = Join-Path (Split-Path $PSScriptRoot -Parent) $logDir
        }
        
        if (!(Test-Path -Path $global:logDirectory)) {
            New-Item -ItemType Directory -Path $global:logDirectory -Force | Out-Null
        }
    }
    
    Write-Log "Configuration applied successfully" "INFO"
    Write-Log "  Client ID: $($global:clientId.Substring(0, [Math]::Min(8, $global:clientId.Length)))..." "INFO"
    Write-Log "  Admin URL: $($global:adminUrl)" "INFO"
    if ($global:defaultSiteUrl) {
        Write-Log "  Default Site URL: $($global:defaultSiteUrl)" "INFO"
    }
    Write-Log "  Keep Versions: $($global:keepVersions)" "INFO"
}

# Export public functions
Export-ModuleMember -Function @(
    'Load-Configuration',
    'Get-ConfigValue',
    'Apply-Configuration'
)
