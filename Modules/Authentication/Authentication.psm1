<# 
.SYNOPSIS
    Authentication module for SharePoint Version Cleaner
    Handles JWT token caching, expiry checking, and connection management
#>

function Initialize-GlobalAuthentication {
    param(
        [switch]$Force
    )
    
    # Check if already authenticated and token is still valid
    if ($global:isGloballyAuthenticated -and $global:globalTokenExpiry -and (Get-Date) -lt $global:globalTokenExpiry -and -not $Force) {
        Write-Log "Using existing global authentication (expires: $($global:globalTokenExpiry))" "INFO"
        return $true
    }
    
    try {
        Write-Log "Initializing global authentication..." "INFO"
        Write-Log "Connecting to: $($global:adminUrl)" "INFO"
        
        # Connect with interactive authentication
        Connect-PnPOnline -Url $global:adminUrl -Interactive -ClientId $global:clientId -ErrorAction Stop
        Write-Log "Connected to PnP Online successfully" "INFO"
        
        Write-Log "Attempting to retrieve access token from PnP connection..." "INFO"
        try {
            $global:globalSharePointToken = Get-PnPAccessToken -ErrorAction Stop
            
            # Validate token format (should be JWT with 3 parts separated by dots)
            if ($global:globalSharePointToken -notmatch '^\S+\.\S+\.\S+$') {
                Write-Log "WARNING: Token does not match JWT format pattern" "WARNING"
            } else {
                Write-Log "Token retrieved successfully (length: $($global:globalSharePointToken.Length) chars)" "INFO"
            }
            
        } catch {
            Write-Log "WARNING: Could not retrieve access token: $($_.Exception.Message)" "WARNING"
            Write-Log "Using fallback: relying on PnP connection instead of cached token" "INFO"
            # Set a flag that we're using connection-based auth
            $global:globalSharePointToken = "CONNECTION_BASED_AUTH"
        }
        
        if (-not $global:globalSharePointToken) {
            throw "Could not establish authentication"
        }
        
        $global:globalAccessToken = $global:globalSharePointToken
        
        # Parse token expiry if we have a real token
        if ($global:globalSharePointToken -ne "CONNECTION_BASED_AUTH") {
            Write-Log "Parsing access token for expiry information..." "INFO"
            $global:globalTokenExpiry = Get-TokenExpiry -Token $global:globalAccessToken
            Write-Log "Token expires at: $($global:globalTokenExpiry)" "INFO"
        } else {
            Write-Log "Using connection-based authentication (no token expiry to parse)" "INFO"
            $global:globalTokenExpiry = (Get-Date).AddHours(1)  # Fallback 1-hour expiry
        }
        
        $global:isGloballyAuthenticated = $true
        Write-Log "Global authentication successful!" "SUCCESS"
        return $true
        
    } catch {
        $global:isGloballyAuthenticated = $false
        $global:globalAccessToken = $null
        $global:globalSharePointToken = $null
        $global:globalTokenExpiry = $null
        Write-Log "Global authentication failed: $($_.Exception.Message)" "ERROR"
        throw $_
    }
}

function Connect-SharePoint {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Url,
        [switch]$SuppressSuccessMessage,
        [switch]$UseCache
    )
    
    # For admin operations, use global authentication if available
    if ($Url -eq $global:adminUrl -and $global:isGloballyAuthenticated -and $global:globalTokenExpiry -and (Get-Date) -lt $global:globalTokenExpiry) {
        if (-not $SuppressSuccessMessage) {
            Write-Log "Already connected to admin URL, reusing connection" "INFO"
        }
        return
    }
    
    # For site operations, check if already connected
    $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
    if ($currentConnection -and $currentConnection.Url -eq $Url) {
        if (-not $SuppressSuccessMessage) {
            Write-Log "Already connected to $Url, reusing existing connection" "INFO"
        }
        return
    }
    
    # Connect using interactive authentication
    try {
        if (-not $SuppressSuccessMessage) {
            Write-Log "Connecting to $Url using interactive authentication" "INFO"
        }
        Connect-PnPOnline -Url $Url -Interactive -ClientId $global:clientId -ErrorAction Stop
        
        $global:connectedSites[$Url] = Get-Date
        
        if (-not $SuppressSuccessMessage) {
            Write-Log "Successfully connected to $Url using interactive authentication" "SUCCESS"
        }
    } catch {
        $errorMessage = "Failed to connect to $Url : $($_.Exception.Message)"
        if (-not $SuppressSuccessMessage) {
            Write-Log $errorMessage "ERROR"
        }
        throw $errorMessage
    }
}

function Get-TokenExpiry {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Token
    )
    
    try {
        # JWT tokens have 3 parts separated by dots: header.payload.signature
        $tokenParts = $Token -split '\.'
        if ($tokenParts.Count -lt 3) {
            Write-Log "Invalid JWT token format (expected 3 parts, got $($tokenParts.Count))" "WARNING"
            return (Get-Date).AddHours(1)
        }
        
        Write-Log "Token has correct format (3 parts). Part lengths: header=$($tokenParts[0].Length), payload=$($tokenParts[1].Length), sig=$($tokenParts[2].Length)" "INFO"
        
        # Extract and decode the payload (second part)
        $payloadBase64 = $tokenParts[1]
        
        # Add padding if needed for base64 decoding
        $payloadPadded = $payloadBase64
        while ($payloadPadded.Length % 4 -ne 0) {
            $payloadPadded += "="
        }
        
        try {
            $decodedBytes = [Convert]::FromBase64String($payloadPadded)
            $decodedPayload = [System.Text.Encoding]::UTF8.GetString($decodedBytes)
            Write-Log "Successfully decoded payload (length: $($decodedPayload.Length) chars)" "INFO"
        } catch {
            Write-Log "ERROR: Failed to decode base64 payload: $($_.Exception.Message)" "WARNING"
            return (Get-Date).AddHours(1)
        }
        
        try {
            $payloadObject = $decodedPayload | ConvertFrom-Json
            Write-Log "Successfully parsed JSON payload" "INFO"
        } catch {
            Write-Log "ERROR: Failed to parse payload as JSON: $($_.Exception.Message). Payload preview: $($decodedPayload.Substring(0, [Math]::Min(200, $decodedPayload.Length)))" "WARNING"
            return (Get-Date).AddHours(1)
        }
        
        if ($payloadObject.exp) {
            Write-Log "Found 'exp' field in payload: $($payloadObject.exp) (type: $($payloadObject.exp.GetType().Name))" "INFO"
            try {
                $expValue = [int64]$payloadObject.exp
                Write-Log "Converted exp to int64: $expValue" "INFO"
                $expiryDate = [DateTimeOffset]::FromUnixTimeSeconds($expValue).DateTime
                Write-Log "Token expiry parsed successfully: $expiryDate" "INFO"
                return $expiryDate
            } catch {
                Write-Log "Error converting/parsing exp timestamp '$($payloadObject.exp)': $($_.Exception.Message). Using 1-hour default." "WARNING"
                return (Get-Date).AddHours(1)
            }
        } else {
            Write-Log "No 'exp' field in token payload. Using 1-hour default expiry." "WARNING"
            return (Get-Date).AddHours(1)
        }
    } catch {
        Write-Log "UNEXPECTED ERROR in Get-TokenExpiry: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" "WARNING"
        return (Get-Date).AddHours(1)
    }
}

function IsJwtTokenExpired {
    param(
        [datetime]$ExpiryDate
    )
    
    return (Get-Date) -ge $ExpiryDate
}

function Clear-AuthenticationCache {
    <#
    .SYNOPSIS
        Clears all cached authentication tokens and disconnects from SharePoint
    #>
    try {
        # Disconnect from any active PnP connections
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
        
        # Clear global authentication state
        $global:globalAccessToken = $null
        $global:globalSharePointToken = $null
        $global:globalTokenExpiry = $null
        $global:isGloballyAuthenticated = $false
        $global:connectedSites = @{}
        $global:cachedSites = @()
        $global:lastSitesFetch = $null
        
        Write-Log "Authentication cache cleared successfully" "SUCCESS"
        return $true
    } catch {
        Write-Log "Error clearing authentication cache: $($_.Exception.Message)" "WARNING"
        return $false
    }
}

# Export public functions
Export-ModuleMember -Function @(
    'Initialize-GlobalAuthentication',
    'Connect-SharePoint',
    'Clear-AuthenticationCache',
    'Get-TokenExpiry',
    'IsJwtTokenExpired'
)

