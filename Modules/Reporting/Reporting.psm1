<# 
.SYNOPSIS
    Reporting module for SharePoint Version Cleaner
    Handles session statistics tracking and HTML report generation
#>

function Add-SessionOperation {
    param(
        [string]$Type, # "Scan" or "Cleanup" or "DryRun"
        [string]$Scope, # "AllSites" or "SingleSite"
        [string]$SiteUrl = "",
        [string]$LibraryName = "",
        [hashtable]$Stats
    )
    
    $operation = @{
        Timestamp = Get-Date
        Type = $Type
        Scope = $Scope
        SiteUrl = $SiteUrl
        LibraryName = $LibraryName
        Stats = $Stats
    }
    
    $global:sessionStats.TotalOperations++
    $global:sessionStats.HasPerformedOperations = $true
    
    if ($Type -eq "Scan") {
        $global:sessionStats.ScanOperations += $operation
    } elseif ($Type -eq "Cleanup" -or $Type -eq "DryRun") {
        $global:sessionStats.CleanupOperations += $operation
        if ($Type -eq "Cleanup") {
            $global:sessionStats.TotalVersionsDeleted += $Stats.VersionsDeleted
            $global:sessionStats.TotalStorageSaved += $Stats.ExcessSize
        }
    } elseif ($Type -eq "InactiveSitesScan" -or $Type -eq "InactiveSitesDeletion") {
        $global:sessionStats.CleanupOperations += $operation
        if ($Type -eq "InactiveSitesDeletion") {
            $global:sessionStats.TotalStorageSaved += $Stats.ExcessSize
        }
    }
    
    # Track affected sites
    if ($Scope -eq "AllSites") {
        if ($Stats.ContainsKey('Sites')) {
            $global:sessionStats.TotalSitesProcessed += $Stats.Sites
        }
    } else {
        if ($SiteUrl -and $SiteUrl -ne "") {
            $siteKey = $SiteUrl.ToLower()
            if (-not $global:sessionStats.AffectedSites.ContainsKey($siteKey)) {
                $global:sessionStats.AffectedSites[$siteKey] = @{
                    Url = $SiteUrl
                    Operations = @()
                    TotalFiles = 0
                    TotalLibraries = 0
                    TotalVersionsDeleted = 0
                    TotalStorageSaved = 0
                }
            }
            
            $global:sessionStats.AffectedSites[$siteKey].Operations += $operation
            $global:sessionStats.AffectedSites[$siteKey].TotalFiles += $Stats.Files
            $global:sessionStats.AffectedSites[$siteKey].TotalLibraries += $Stats.Libraries
            if ($Stats.ContainsKey('VersionsDeleted')) {
                $global:sessionStats.AffectedSites[$siteKey].TotalVersionsDeleted += $Stats.VersionsDeleted
            }
            if ($Stats.ContainsKey('ExcessSize')) {
                $global:sessionStats.AffectedSites[$siteKey].TotalStorageSaved += $Stats.ExcessSize
            }
            
            $global:sessionStats.TotalSitesProcessed = $global:sessionStats.AffectedSites.Count
        }
    }
    
    $global:sessionStats.TotalFilesProcessed += $Stats.Files
    $global:sessionStats.TotalLibrariesProcessed += $Stats.Libraries
    $global:sessionStats.TotalErrors += $Stats.Errors
    
    Write-Log "Session operation recorded: $Type - $Scope" "INFO"
}

function Generate-SessionReport {
    try {
        if (-not $global:sessionStats.HasPerformedOperations) {
            [System.Windows.Forms.MessageBox]::Show(
                "No operations have been performed yet. Please run a scan or cleanup operation first.",
                "No Data Available",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            return
        }
        
        $reportPath = "$global:logDirectory\SharePoint_Version_Cleaner_Report_$script:dateTime.html"
        
        $html = Build-HtmlReport
        $html | Out-File -FilePath $reportPath -Encoding UTF8
        
        Write-Log "Session report generated: $reportPath" "SUCCESS"
        Start-Process $reportPath
        
        [System.Windows.Forms.MessageBox]::Show(
            "Session report has been generated and opened in your browser.`n`nFile location: $reportPath",
            "Report Generated",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        
    } catch {
        Write-Log "Error generating session report: $($_.Exception.Message)" "ERROR"
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to generate session report: $($_.Exception.Message)",
            "Report Generation Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
}

function Build-HtmlReport {
    $html = @"
<!DOCTYPE html>
<html>
<head>
    <title>SharePoint Version Cleaner - Session Report</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; background-color: #f5f5f5; }
        .header { background-color: #0078d4; color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; }
        .header h1 { margin: 0; }
        .header p { margin: 5px 0; opacity: 0.9; }
        .section { background-color: white; padding: 20px; border-radius: 8px; margin-bottom: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        .section h2 { color: #0078d4; margin-top: 0; border-bottom: 2px solid #0078d4; padding-bottom: 10px; }
        .stats-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px; }
        .stat-box { background-color: #f8f9fa; padding: 15px; border-left: 4px solid #0078d4; border-radius: 4px; }
        .stat-number { font-size: 24px; font-weight: bold; color: #0078d4; }
        .stat-label { color: #666; font-size: 14px; }
        table { width: 100%; border-collapse: collapse; margin-top: 15px; }
        th, td { padding: 12px; text-align: left; border-bottom: 1px solid #ddd; }
        th { background-color: #0078d4; color: white; }
        tr:hover { background-color: #f5f5f5; }
        .success { color: #28a745; font-weight: bold; }
        .error { color: #dc3545; font-weight: bold; }
    </style>
</head>
<body>
    <div class="header">
        <h1>📊 SharePoint Version Cleaner - Session Report</h1>
        <p>Generated on: $(Get-Date -Format "dddd, MMMM dd, yyyy 'at' HH:mm:ss")</p>
        <p>Session Duration: $((Get-Date) - $global:sessionStats.SessionStart | ForEach-Object { "{0:hh\:mm\:ss}" -f $_ })</p>
    </div>
    
    <div class="section">
        <h2>📈 Session Overview</h2>
        <div class="stats-grid">
            <div class="stat-box">
                <div class="stat-number">$($global:sessionStats.TotalOperations)</div>
                <div class="stat-label">Total Operations</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($global:sessionStats.TotalSitesProcessed)</div>
                <div class="stat-label">Sites Processed</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($global:sessionStats.TotalLibrariesProcessed)</div>
                <div class="stat-label">Libraries Processed</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($global:sessionStats.TotalFilesProcessed)</div>
                <div class="stat-label">Files Processed</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($global:sessionStats.TotalVersionsDeleted)</div>
                <div class="stat-label">Versions Deleted</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$(Format-FileSize -Size $global:sessionStats.TotalStorageSaved)</div>
                <div class="stat-label">Storage Saved</div>
            </div>
        </div>
    </div>
"@

    if ($global:sessionStats.ScanOperations.Count -gt 0) {
        $html += @"
    <div class="section">
        <h2>🔍 Scan Operations</h2>
        <table>
            <tr>
                <th>Time</th>
                <th>Scope</th>
                <th>Target</th>
                <th>Libraries</th>
                <th>Files</th>
                <th>Potential Savings</th>
                <th>Errors</th>
            </tr>
"@
        foreach ($op in $global:sessionStats.ScanOperations) {
            $target = if ($op.Scope -eq "AllSites") { "All Sites" } else {
                if ($op.LibraryName) { "$($op.SiteUrl) - $($op.LibraryName)" } else { $op.SiteUrl }
            }
            $html += @"
            <tr>
                <td>$($op.Timestamp.ToString("HH:mm:ss"))</td>
                <td>$($op.Scope)</td>
                <td>$target</td>
                <td>$($op.Stats.Libraries)</td>
                <td>$($op.Stats.Files)</td>
                <td>$(Format-FileSize -Size $op.Stats.ExcessSize)</td>
                <td class="$(if ($op.Stats.Errors -gt 0) { 'error' } else { 'success' })">$($op.Stats.Errors)</td>
            </tr>
"@
        }
        $html += "</table></div>"
    }

    if ($global:sessionStats.CleanupOperations.Count -gt 0) {
        $html += @"
    <div class="section">
        <h2>🧹 Cleanup Operations</h2>
        <table>
            <tr>
                <th>Time</th>
                <th>Type</th>
                <th>Scope</th>
                <th>Versions Deleted</th>
                <th>Storage Saved</th>
                <th>Errors</th>
            </tr>
"@
        foreach ($op in $global:sessionStats.CleanupOperations) {
            $html += @"
            <tr>
                <td>$($op.Timestamp.ToString("HH:mm:ss"))</td>
                <td>$($op.Type)</td>
                <td>$($op.Scope)</td>
                <td>$(if ($op.Stats.ContainsKey('VersionsDeleted')) { $op.Stats.VersionsDeleted } else { 'N/A' })</td>
                <td>$(if ($op.Stats.ContainsKey('ExcessSize')) { Format-FileSize -Size $op.Stats.ExcessSize } else { 'N/A' })</td>
                <td class="$(if ($op.Stats.Errors -gt 0) { 'error' } else { 'success' })">$($op.Stats.Errors)</td>
            </tr>
"@
        }
        $html += "</table></div>"
    }

    if ($global:sessionStats.AffectedSites.Count -gt 0) {
        $html += @"
    <div class="section">
        <h2>🌐 Affected SharePoint Sites</h2>
        <table>
            <tr>
                <th>Site URL</th>
                <th>Libraries</th>
                <th>Files</th>
                <th>Versions Deleted</th>
                <th>Storage Saved</th>
            </tr>
"@
        foreach ($siteInfo in $global:sessionStats.AffectedSites.Values) {
            $html += @"
            <tr>
                <td><a href="$($siteInfo.Url)" target="_blank">$($siteInfo.Url)</a></td>
                <td>$($siteInfo.TotalLibraries)</td>
                <td>$($siteInfo.TotalFiles)</td>
                <td>$($siteInfo.TotalVersionsDeleted)</td>
                <td>$(Format-FileSize -Size $siteInfo.TotalStorageSaved)</td>
            </tr>
"@
        }
        $html += "</table></div>"
    }

    $html += @"
    <div class="section">
        <h2>📋 Summary</h2>
        <p>This session processed <strong>$($global:sessionStats.TotalFilesProcessed)</strong> files across <strong>$($global:sessionStats.TotalLibrariesProcessed)</strong> libraries in <strong>$($global:sessionStats.TotalSitesProcessed)</strong> sites.</p>
        $(if ($global:sessionStats.TotalVersionsDeleted -gt 0) {
            "<p class='success'>✅ Successfully deleted <strong>$($global:sessionStats.TotalVersionsDeleted)</strong> excess file versions, saving <strong>$(Format-FileSize -Size $global:sessionStats.TotalStorageSaved)</strong> of storage space.</p>"
        })
        $(if ($global:sessionStats.TotalErrors -gt 0) {
            "<p class='error'>⚠️ Encountered <strong>$($global:sessionStats.TotalErrors)</strong> errors during processing.</p>"
        } else {
            "<p class='success'>✅ All operations completed successfully with no errors.</p>"
        })
    </div>
</body>
</html>
"@
    return $html
}

# Export public functions
Export-ModuleMember -Function @(
    'Add-SessionOperation',
    'Generate-SessionReport'
)

