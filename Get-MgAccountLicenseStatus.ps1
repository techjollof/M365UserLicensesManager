<#
.SYNOPSIS
    Gets licensed and unlicensed users in Microsoft 365.

.DESCRIPTION
    This function retrieves all users in Microsoft 365 and categorizes them as licensed or unlicensed.
    It requires the Microsoft Graph PowerShell module to be installed and connected first.

.PARAMETER ExportToCSV
    Specifies whether to export the results to CSV files. Default is $false.

.PARAMETER OutputPath
    Specifies the directory path where CSV files will be saved if ExportToCSV is $true.
    Default is the current user's Downloads folder.

.EXAMPLE
    Get-MgAccountLicenseStatus
    Displays licensed and unlicensed users in the console.

.EXAMPLE
    Get-MgAccountLicenseStatus -ExportToCSV -OutputPath "C:\Reports"
    Exports licensed and unlicensed users to CSV files in the specified directory.

.NOTES
    Requires Microsoft Graph PowerShell module.
    Run Connect-MgGraph first with appropriate permissions (User.Read.All minimum).
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [bool]$ExportToCSV = $true,

    [parameter(Mandatory = $false)]
    [switch]$DisplayInConsole,
        
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "$Home\Downloads"
)

# Check if Microsoft Graph module is available
if (-not (Get-Module -ListAvailable -Name Microsoft.Graph)) {
    Write-Error "Microsoft Graph PowerShell module is not installed. Please install it first using: Install-Module Microsoft.Graph -Scope CurrentUser -Force"
    return
}

# Check if we're connected to Microsoft Graph
try {
    $context = Get-MgContext
    if (-not $context) {
        Write-Error "Not connected to Microsoft Graph. Please run Connect-MgGraph first with appropriate permissions."
        return
    }
}
catch {
    Write-Error "Error checking Microsoft Graph connection. Please ensure you're connected with Connect-MgGraph -Scopes User.Read.All, Organization.Read.All"
    return
}
    

try {
    Write-Verbose "Retrieving all users..."
    $allUsers = Get-MgUser -ConsistencyLevel eventual -All -Property Id, DisplayName, UserPrincipalName, AssignedLicenses, LicenseAssignmentStates, LicenseDetails, AssignedPlans
        
    $licensedUsers = $allUsers | Where-Object { $_.AssignedLicenses.Count -gt 0 }
    $unlicensedUsers = $allUsers | Where-Object { $_.AssignedLicenses.Count -eq 0 }

    $licensedOutput = $licensedUsers | Select-Object DisplayName, UserPrincipalName, AssignedLicenses, AssignedPlans, AssignedLicenses, @{Name = "LicenseCount"; Expression = { $_.AssignedLicenses.Count } }
    $unlicensedOutput = $unlicensedUsers | Select-Object DisplayName, UserPrincipalName

    if ($DisplayInConsole) {
        Write-Host "`nLicensed Users ($($licensedUsers.Count)):" -ForegroundColor Green
        $licensedOutput | Format-Table

        Write-Host "`nUnlicensed Users ($($unlicensedUsers.Count)):" -ForegroundColor Yellow
        $unlicensedOutput | Format-Table
    }
    else {
        Write-Host "`nLicensed Users ($($licensedUsers.Count)):" -ForegroundColor Green
        Write-Host "`nUnlicensed Users ($($unlicensedUsers.Count)):" -ForegroundColor Yellow
    }

    if ($ExportToCSV) {
        if (-not (Test-Path -Path $OutputPath)) {
            New-Item -ItemType Directory -Path $OutputPath | Out-Null
        }

        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $licensedPath = Join-Path -Path $OutputPath -ChildPath "LicensedUsers_$timestamp.csv"
        $unlicensedPath = Join-Path -Path $OutputPath -ChildPath "UnlicensedUsers_$timestamp.csv"

        $licensedOutput | Export-Csv -Path $licensedPath -NoTypeInformation -Encoding UTF8
        $unlicensedOutput | Export-Csv -Path $unlicensedPath -NoTypeInformation -Encoding UTF8

        Write-Host "`nExported licensed users to: $licensedPath" -ForegroundColor Cyan
        Write-Host "Exported unlicensed users to: $unlicensedPath" -ForegroundColor Cyan
    }

    # Optional: return results
    return [PSCustomObject]@{
        LicensedUsers   = $licensedOutput
        UnlicensedUsers = $unlicensedOutput
    }
}
catch {
    Write-Error "An error occurred while retrieving user data: $_"
}
