<#
.SYNOPSIS
    Moves a user from one Microsoft 365 license plan to another.

.DESCRIPTION
    This function removes a specified Microsoft 365 license SKU from a user and assigns a new one.
    It optionally disables specific service plans in the new license.
    Requires Microsoft.Graph PowerShell module and an active connection via Connect-MgGraph.

.PARAMETER UserIds
    One or more UPNs or paths to CSV files containing user identifiers.

.PARAMETER LicenseToRemove
    The SKU part number of the current license (e.g., "SPE_E3").

.PARAMETER LicenseToAdd
    The SKU part number of the new license (e.g., "SPE_E5").

.PARAMETER DisabledPlans
    Optional. Array of service plan GUIDs to disable in the assigned license (e.g., Exchange, Yammer, etc.).

.PARAMETER SelectLicense
    If set, allows you to pick SKUs interactively from a grid.

.EXAMPLE
    Set-MgUserLicenseUpgrade -UserIds "user@domain.com" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5"

.EXAMPLE
    Set-MgUserLicenseUpgrade -UserIds ".\users.csv" -SelectLicense

.EXAMPLE
    Set-MgUserLicenseUpgrade -UserIds "user@domain.com" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5" -DisabledPlans "12345678-aaaa-bbbb-cccc-1234567890ab"

#>

# Requires -Modules Microsoft.Graph.Users, Microsoft.Graph.Authentication

[CmdletBinding(SupportsShouldProcess = $true)]
param (
    # User Input Parameters
    [Parameter(Mandatory = $true, Position = 0)]
    [Alias('UserPrincipalName', 'UPNs', 'UserId', 'UserSource')]
    [string[]]$UserIds,
    
    # License Management Parameters
    [Parameter(Mandatory, ParameterSetName = 'ManualSelection')]
    [Alias('RemoveSku', 'OldLicense', 'CSku', 'CurrentLicense')]
    [string]$LicenseToRemove,
    
    [Parameter(Mandatory, ParameterSetName = 'ManualSelection')]
    [Alias('AddSku', 'NewLicense', 'NSku', 'TargetLicense')]
    [string]$LicenseToAdd,
    
    [Parameter(ParameterSetName = 'ManualSelection')]
    [Alias('DisabledService', 'RemovePlans', 'RemoveServices', 'DisabledServicePlans')]
    [string[]]$DisabledPlans,
    
    # Interactive Mode Parameters
    [Parameter(Mandatory, ParameterSetName = 'InteractiveMode')]
    [Alias('UseLicensesPicker', 'InteractiveLicenseSelect')]
    [switch]$SelectLicense,
    
    [Parameter(ParameterSetName = 'InteractiveMode')]
    [Alias('PickPlansToDisable', 'InteractivePlanSelect')]
    [switch]$SelectDisabledPlans,
    
    # Advanced Options
    [Parameter()]
    [Alias('PreservePlanState')]
    [switch]$KeepExistingPlanState
)
    
    
function Get-ProcessedUserIds {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$UserIds
    )
        
    # Initialize an array to hold all processed user identifiers
    $processedUserIds = @()
        
    foreach ($userIdItem in $UserIds) {
        # Check if the item is a path to a CSV file
        if ($userIdItem -like "*.csv" -and (Test-Path -Path $userIdItem -PathType Leaf)) {
            try {
                # Import the CSV file
                $csvData = Import-Csv -Path $userIdItem -ErrorAction Stop
        
                if ($csvData.Count -eq 0) {
                    Write-Warning "CSV file '$userIdItem' is empty."
                    continue
                }
        
                # Check if CSV has structured columns
                if ($csvData[0].PSObject.Properties.Count -gt 1) {
                    # Structured CSV - look for common column names
                    $columnNames = $csvData[0].PSObject.Properties.Name
        
                    # Try to find a column with common identifiers
                    $idColumn = $columnNames | Where-Object {
                        $_ -match 'User(ID|PrincipalName|Email(Address)?)' -or
                        $_ -eq 'ID' -or 
                        $_ -eq 'Email' -or 
                        $_ -eq 'EmailAddress'
                    } | Select-Object -First 1
        
                    if ($idColumn) {
                        $processedUserIds += $csvData.$idColumn
                    }
                    else {
                        Write-Warning "Could not identify user identifier column in CSV file '$userIdItem'. Using the first column."
                        $firstColumn = $columnNames[0]
                        $processedUserIds += $csvData.$firstColumn
                    }
                }
                else {
                    # Unstructured CSV - assume each row is a single user ID
                    $processedUserIds += $csvData | ForEach-Object { $_.PSObject.Properties.Value }
                }
            }
            catch {
                Write-Error "Failed to process CSV file '$userIdItem': $_"
                continue
            }
        }
        else {
            # Not a CSV file, treat as direct user ID
            $processedUserIds += $userIdItem
        }
    }
        
    return ($processedUserIds | Where-Object { $_ -and ($_ -is [string]) } | ForEach-Object { $_.Trim() } | Sort-Object -Unique)
}
    
    
########################
    
    
# Define log file relative to script path
$logFile = Join-Path -Path $PSScriptRoot -ChildPath "Logs\LicenseUpgrade_$(Get-Date -Format 'yyyy-MM-dd').log"
    
# Ensure log directory exists
$logDir = Split-Path $logFile
if (-not (Test-Path $logDir)) {
    New-Item -Path $logDir -ItemType Directory | Out-Null
}
    
# Function to log messages
function Write-Log {
    param (
        [string]$message,
        [string]$type = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    "$timestamp [$type] $message" | Out-File -FilePath $logFile -Append
}
    
    
################################
    
    
try {
    Write-Verbose "Getting available subscribed SKUs..."
    $skus = Get-MgSubscribedSku -All
    
    if ($SelectLicense) {
        Write-Verbose "Auto license selection is enabled. Selecting SKUs via GUI..."
        $fromSkus = $skus | Out-GridView -Title "Select SKUs to REMOVE (Current)" -PassThru
        $toSkus = $skus | Out-GridView -Title "Select SKUs to ASSIGN (New)" -PassThru
    }
    else {
        Write-Verbose "Auto license selection is disabled. Using provided SKU part numbers."
        $fromSkus = $skus | Where-Object { $_.SkuPartNumber -eq $LicenseToRemove }
        $toSkus = $skus | Where-Object { $_.SkuPartNumber -eq $LicenseToAdd }
    }
    
    if (-not $fromSkus) {
        Write-Log "Source SKU(s) not found." -type "ERROR"
        return
    }
    
    if (-not $toSkus) {
        Write-Log "Target SKU(s) not found." -type "ERROR"
        return
    }
    
    # Build arrays for add and remove operations
    $fromLicenseIds = $fromSkus | ForEach-Object { $_.SkuId }
    # Build addLicenses array with disabled plans only if present in the SKU
    if ($SelectDisabledPlans -or $DisabledPlans) {
        if ($SelectDisabledPlans) {
            $allPlans = $toSkus.ServicePlans | Where-Object { $_.AppliesTo -eq "user" } | Sort-Object ServicePlanName -Unique
            $selectedPlans = $allPlans | Out-GridView -Title "Select service plans to DISABLE" -PassThru
            write-Log "Selected plans: $($selectedPlans.ServicePlanName -join ', ')" -LogPath $LogPath -Level "INFO"
            # Store the selected plan IDs in $DisabledPlans
            $DisabledPlans = $selectedPlans.ServicePlanId
        }
        else {
            $DisabledPlans = $DisabledPlans | ForEach-Object { $_.Trim() } | Sort-Object -Unique
        }
    }
    
    
    $allUsers = Get-MgUser -ConsistencyLevel eventual -All -Property Id, DisplayName, UserPrincipalName, AssignedLicenses, LicenseAssignmentStates, LicenseDetails, AssignedPlans
        
}
catch {
    Write-Error "An error occurred while preparing license information: $_"
    return
}
    
# Get valid user IDs
$ResolvedUserIds = Get-ProcessedUserIds -UserIds $UserIds
if ($ResolvedUserIds.Count -eq 0) {
    Write-Error "No valid user IDs found."
    return
}
    
# Logging path setup
$LogPath = Join-Path -Path $PSScriptRoot -ChildPath "LicenseChange.log"
    
foreach ($user in $ResolvedUserIds) {
    
    $isLicenseAdded = $false
    $isLicenseRemove = $false
    
    $fromSkuStr = ($fromSkus | Select-Object -ExpandProperty SkuPartNumber) -join ', '
    $toSkuStr = ($toSkus   | Select-Object -ExpandProperty SkuPartNumber) -join ', '
        
    # Fetch user license information (current and existing AssignedLicenses and DisabledPlans)
    $userLicenseInfo = ($allUsers | Where-Object { $_.UserPrincipalName.ToLower() -eq $user.ToLower() }).AssignedLicenses
    $userExistingLicenses = $userLicenseInfo.SkuId
    $userExistingDisabledPlans = $userLicenseInfo.DisabledPlans
    
    # Create the toLicenses structure
    $toLicenses = foreach ($sku in $toSkus) {
        $availablePlans = $sku.ServicePlans
    
        # Get newly selected plans to disable based on input
        $newlySelectedToDisable = $availablePlans | Where-Object {
            $_.ServicePlanId -in $DisabledPlans -or $_.ServicePlanName -in $DisabledPlans
        }
    
        # Combine the new disabled plans with the existing disabled plans (make sure both are in the right SKU) and remove duplicates
        $existingForSkuDisabledPlan = $null
        if ($KeepExistingPlanState) {
            $existingForSkuDisabledPlan = $availablePlans | Where-Object {
                $_.ServicePlanId -in $userExistingDisabledPlans -or $_.ServicePlanName -in $userExistingDisabledPlans
            }
        }
        # $combinedDisabledPlans
        $combinedDisabledPlans = @(
            $newlySelectedToDisable.ServicePlanId
            $existingForSkuDisabledPlan.ServicePlanId
        ) | Where-Object { $_ } | Sort-Object -Unique
            
    
        if ($combinedDisabledPlans.Count -gt 0) {
            # If there are disabled plans, include them in the result
            @{
                skuId         = $sku.SkuId
                disabledPlans = @($combinedDisabledPlans)
            }
        }
        else {
            # No disabled plans, only the SKU
            @{ skuId = $sku.SkuId }
        }
    }
        
    
    ##################################### License Removal and Addition #####################################
    
    if ($PSCmdlet.ShouldProcess("$user", "Remove license '$fromSkuStr' and assign '$toSkuStr'")) {
        try {
    
            # Check if the user has any licenses assigned before proceeding with removal
            if ($userExistingLicenses) {
                $filteredFromLicenses = $fromLicenseIds | Where-Object { $_ -in $userExistingLicenses }
    
                # Step 1: Remove licenses
                if ($filteredFromLicenses.Count -gt 0) {
                    $removeBody = @{
                        removeLicenses = @($filteredFromLicenses)
                        addLicenses    = @()
                    } | ConvertTo-Json -Depth 10
    
                    # Write-Host "Removing license for...." -ForegroundColor Yellow
                    # $removeBody
    
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$user/assignLicense" -Body $removeBody -ContentType "application/json" | Out-Null
                    $isLicenseRemove = $true
                    if ($isLicenseRemove) {
                        Write-Log "License removed for $user from [$fromSkuStr] to [$toSkuStr]" -LogPath $LogPath -Level "INFO"
                    }
                }
            }
            else {
                Write-Log "No matching licenses found for removal for user $user." -LogPath $LogPath -Level "WARNING"
            }
    
            # Step 2: Add licenses: Check if the user already has the new license assigned before adding it
            # $filteredToLicenses = $toLicenses | Where-Object { $_.skuId -notin $userExistingLicenses }
            if ($toLicenses) {
    
                $addBody = @{
                    removeLicenses = @()
                    addLicenses    = @($toLicenses)
                } | ConvertTo-Json -Depth 10
    
                # write-host "Adding license for...." -ForegroundColor Yellow
                # $addBody
    
                Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$user/assignLicense" -Body $addBody -ContentType "application/json"  | Out-Null
                $isLicenseAdded = $true
                if ($isLicenseAdded) {
                    Write-Log "License updated for $user from [$fromSkuStr] to [$toSkuStr]" -LogPath $LogPath -Level "INFO"
                }
            }
            else {
                Write-Log "No matching licenses found for addition or already assigned to the user $user." -LogPath $LogPath -Level "WARNING"
            }
    
        }
        catch {
            Write-Error "Failed to update license for $user : $($_)"
            Write-Log "Failed to update license for $user : $($_)" -LogPath $LogPath -Level "ERROR"
        }
        Write-Host "✅ License processing completed for $($ResolvedUserIds.Count) user(s)." -ForegroundColor Green
        Write-Log "✅ License processing completed for $($ResolvedUserIds.Count) user(s)." -LogPath $LogPath -Level "INFO"
    }
    else {
        $dryRunMsg = "🧪 [Dry Run] Would upgrade license for $user from [$fromSkuStr] to [$toSkuStr]"
        Write-Host $dryRunMsg -ForegroundColor Gray
        Write-Log $dryRunMsg -LogPath $LogPath -Level "DRYRUN"
    }
}