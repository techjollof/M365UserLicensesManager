
# Parameters for the script
[CmdletBinding(DefaultParameterSetName = 'ManualPassword', SupportsShouldProcess = $true)]
param (
    [Parameter(Mandatory, Position = 0)]
    [string]$UserCsvFilePath,

    [Parameter(Position = 1)]
    [string]$ResultExportFilePath,

    # License Assignment Parameters
    [Parameter(ParameterSetName = "DirectLicense")]
    [string[]]$AssignedLicenseSkus,

    [Parameter(ParameterSetName = "PromptForLicense")]
    [switch]$PromptForAssignedLicenseSkus,

    # License Plan Disabling Parameters
    [Parameter(ParameterSetName = "DirectLicense")]
    [Parameter(ParameterSetName = "DirectDisablePlans")]
    [string[]]$DisableLicensePlans,

    [Parameter(ParameterSetName = "PromptForLicense")]
    [Parameter(ParameterSetName = "PromptForDisablePlans")]
    [switch]$PromptForDisableLicensePlans,

    # User Creation Parameters (available in all parameter sets)
    [Parameter(ParameterSetName = "DirectLicense")]
    [Parameter(ParameterSetName = "PromptForLicense")]
    [ValidateNotNullOrEmpty()]
    [string]$Password,

    [Parameter(ParameterSetName = "DirectLicense")]
    [Parameter(ParameterSetName = "PromptForLicense")]
    [switch]$AutoGeneratePassword,

    [Parameter()]
    [ValidateRange(10, 20)]
    [int]$AutoPasswordLength = 12,

    [switch]$SamePasswordForAll,

    # Password change switch (available in all parameter sets)
    [Parameter()]
    [bool]$ForcePasswordChange = $true,

    [Parameter()]
    [string]$LogFilePath
)

# Validate mutual exclusivity (though parameter sets should prevent this)
if ($PSBoundParameters.ContainsKey('AssignedLicenseSkus') -and $PromptForAssignedLicenseSkus) {
    throw "Parameters -AssignedLicenseSkus and -PromptForAssignedLicenseSkus are mutually exclusive and cannot be used together."
}

if ($PSBoundParameters.ContainsKey('Password') -and $AutoGeneratePassword) {
    throw "Parameters -Password and -AutoGeneratePassword are mutually exclusive and cannot be used together."
}

# Set default log file path if LogFilePath is not provided
if (-not $LogFilePath) {
    $LogFilePath = Join-Path $PSScriptRoot "UserCreationLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
}

# Function to write log entries
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = "[$timestamp][$Level] $Message"
    Add-Content -Path $LogFilePath -Value $entry
    Write-Verbose $entry
}

# Function to handle retries on transient errors (throttling or server errors)
function Invoke-WithRetry {
    param (
        [scriptblock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$DelaySeconds = 5
    )
    $attempt = 0
    do {
        try {
            $attempt++
            return & $ScriptBlock
        }
        catch {
            if ($_.Exception.Response.StatusCode.value__ -in 429, 500, 503 -and $attempt -lt $MaxRetries) {
                Write-Log "Transient error encountered. Retrying in $DelaySeconds seconds..." "WARN"
                Start-Sleep -Seconds $DelaySeconds
            } else {
                throw
            }
        }
    } while ($attempt -lt $MaxRetries)
}

$users = Import-Csv -Path $UserCsvFilePath

function New-RandomPassword {
    [CmdletBinding()]
    param (
        [int]$Length = 16,
        [switch]$EnforceComplexity
    )

    $lower = "abcdefghijklmnopqrstuvwxyz".ToCharArray()
    $upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray()
    $numbers = "0123456789".ToCharArray()
    $special = "!@#$%^&*()_+-=[]{}|;:,.<>?".ToCharArray()
    $allChars = $lower + $upper + $numbers + $special

    if ($EnforceComplexity) {
        $password = @(
            $lower | Get-Random
            $upper | Get-Random
            $numbers | Get-Random
            $special | Get-Random
        )
        $password += 1..($Length - 4) | ForEach-Object { $allChars | Get-Random }
        $password = $password | Sort-Object { Get-Random }
    }
    else {
        $password = 1..$Length | ForEach-Object { $allChars | Get-Random }
    }
    return -join $password
}

if ($AutoGeneratePassword) {
    if ($SamePasswordForAll) {
        $Password = New-RandomPassword -Length $AutoPasswordLength
    }
}
elseif (-not $Password) {
    Throw "Password must be defined or use -AutoPassword. Add -SamePasswordForAll if you want to use the same password for all users."
}

$PasswordProfile = @{
    Password                      = $Password
    ForceChangePasswordNextSignIn = $ForcePasswordChange
}

$results = @()
$addLicenses = @()
$licenseParams = $null  # Initialize to avoid undefined return

# Handle License Assignment (Mutually Exclusive)
if ($AssignedLicenseSkus) {
    Write-Log "AssignedLicenseSkus: $AssignedLicenseSkus"
    $AssignedSku = Get-MgSubscribedSku -All | Where-Object { $_.SkuPartNumber -in $AssignedLicenseSkus -or $_.SkuId -in $AssignedLicenseSkus }

    if (-not $AssignedSku) {
        Write-Warning "No matching SKUs found. User will be created without licenses."
    }
}
elseif ($PromptForAssignedLicenseSkus) {
    Write-Log "Please select SKUs from the list below."
    $AssignedSku = Get-MgSubscribedSku -All | Where-Object { ($_.PrepaidUnits.Enabled -ne $_.ConsumedUnits) } | Out-GridView -PassThru -Title "Select License SKUs to Assign"

    if (-not $AssignedSku) {
        Write-Warning "No license selected. Continuing without license assignment."
    }
}

# Handle Disabled Plans
if ($PromptForDisableLicensePlans -and $AssignedSku) {
    Write-Log "Please select service plans to disable from the list below."
    $SelectedDisableLicensePlans = $AssignedSku.ServicePlans | Where-Object { $_.AppliesTo -eq "User" -and $_.ProvisioningStatus -eq "Success" } | Out-GridView -PassThru -Title "Select Service Plans to Disable"
    [Object]$DisableLicensePlans = $SelectedDisableLicensePlans | Select-Object ServicePlanId, ServicePlanName
}

# Build License
if ($AssignedSku) {
    foreach ($sku in $AssignedSku) {
        $licenseAssignment = @{ skuId = $sku.SkuId }

        if ($DisableLicensePlans) {
            # Log the processing of service plans for the current SKU
            Write-Log "Processing the plans for SKU: $($sku.SkuPartNumber)" "WARN"
            
            # Handle disabling service plans based on the prompt flag
            $disabledPlans = if ($PromptForDisableLicensePlans) {
                $sku.ServicePlans | Where-Object { $_.ServicePlanId -in $DisableLicensePlans.ServicePlanId -or $_.ServicePlanName -in $DisableLicensePlans.ServicePlanName }
            } else {
                $sku.ServicePlans | Where-Object { $_.ServicePlanId -in $DisableLicensePlans -or $_.ServicePlanName -in $DisableLicensePlans }
            }

            if ($disabledPlans) {
                $licenseAssignment['disabledPlans'] = @($disabledPlans.ServicePlanId)
            }
        }
        # Add license assignment to the list of licenses to assign
        $addLicenses += $licenseAssignment
    }

    # Prepare the license assignment parameters
    $licenseParams = @{
        addLicenses    = $addLicenses
        removeLicenses = @()
    } | ConvertTo-Json -Depth 10
}

foreach ($user in $users) {

    $newUserParams = @{}
    # Convert to truly custom object to support custom fields
    $result = @{}

    $user.PSObject.Properties | ForEach-Object {
        if ($_.Value) {
            $newUserParams[$_.Name] = $_.Value
        }
        $result[$_.Name] = $_.Value
    }

    try {

        # Handle password
        if ($AutoGeneratePassword -and -not $SamePasswordForAll) {
            $Password = New-RandomPassword -Length $AutoPasswordLength
            $PasswordProfile.Password = $Password
        }

        $newUserParams['PasswordProfile'] = $PasswordProfile
        $newUserParams['AccountEnabled'] = $true

        if ($PSCmdlet.ShouldProcess($newUserParams.UserPrincipalName, "Create New MgUser")) {
            $newUser = Invoke-WithRetry { New-MgUser @newUserParams }
            Write-Log "Successfully created user $($newUser.UserPrincipalName)"

            if ($licenseParams) {
                Invoke-WithRetry {
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($newUser.Id)/assignLicense" -Body $licenseParams -ContentType "application/json" | Out-Null
                }

                $result['AssignedLicenseSkus'] = $AssignedSku.SkuPartNumber -join ','
                $result['AssignedLicenseId'] = $AssignedSku.SkuId -join ','

                if ($DisableLicensePlans) {
                    $result['DisabledLicensePlans'] = $disabledPlans.ServicePlanName -join ','
                }

                Write-Log "Assigned licenses to user $($newUser.UserPrincipalName)"
            }

            $result.Password = $Password
            $result['Status'] = 'Success'
            $result['UserId'] = $newUser.Id
        }
    }
    catch {
        Write-Log "Failed to create user $($user.UserPrincipalName): $_" "ERROR"
        Write-Log "Error details: $($_)" "ERROR"
        $result['Status'] = 'Failed'
        $result['ErrorMessage'] = $_.Exception.Message
    }

    $results += [PSCustomObject]$result
}

# Export results if any
# Set default export file path if ResultExportFilePath is not provided
if (-not $ResultExportFilePath) {
    $ResultExportFilePath = Join-Path $PSScriptRoot "UserCreationLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
}

# Export results if any
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $ResultExportFilePath -NoTypeInformation -Force
    Write-Log "Results exported to $ResultExportFilePath"
}
else {
    Write-Warning "No results to export"
}

Write-Log "User creation script completed. $(($results | Where-Object Status -eq 'Success').Count) succeeded, $(($results | Where-Object Status -eq 'Failed').Count) failed."
