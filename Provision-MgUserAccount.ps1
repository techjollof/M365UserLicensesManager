<#
.SYNOPSIS
    Bulk creates Microsoft 365 users based on a CSV input file with optional license assignment and password handling.

.DESCRIPTION
    This script creates Microsoft 365 users using Microsoft Graph API based on data in a CSV file.
    It supports both static and auto-generated passwords, interactive or direct license assignment,
    and the optional disabling of service plans. Results and logs are exported to user-defined paths.
    Passwords can be generated per user or applied uniformly across all accounts.

.PARAMETER UserIdsCsv
    Specifies the path to the CSV file containing user data. Required columns: UserPrincipalName, DisplayName, MailNickname.

.PARAMETER ResultExportFilePath
    Specifies the path and base filename where the user creation results will be exported.
    A timestamp will be appended automatically to the filename.

.PARAMETER AssignLicenses
    Specifies the license SKU(s) to assign directly to the user(s). Mutually exclusive with -PromptForAssignedLicenseSkus.

.PARAMETER SelectLicenses
    Prompts the user to select license SKU(s) interactively. Mutually exclusive with -AssignedLicenseSkus.

.PARAMETER DisablePlans
    Specifies one or more service plans to be disabled during license assignment. Mutually exclusive with -PromptForDisableLicensePlans.

.PARAMETER SelectDisabledPlans
    Prompts the user to select service plans to disable interactively. Mutually exclusive with -DisableLicensePlans.

.PARAMETER Password
    Specifies a static password to assign to all users. Mutually exclusive with -AutoGeneratePassword.

.PARAMETER AutoGeneratePassword
    Automatically generates random passwords for users. Mutually exclusive with -Password.

.PARAMETER AutoPasswordLength
    Specifies the length of auto-generated passwords. Must be between 10 and 20. Default is 12.

.PARAMETER SamePasswordForAll
    When used with -AutoGeneratePassword, applies the same password to all users instead of generating unique ones.

.PARAMETER ForcePasswordChange
    If set to $true, users will be prompted to change their password upon first login. Default is true.

.PARAMETER LogFilePath
    Specifies the full path to the log file for this script run. If not specified, a default log file is created.

.EXAMPLE
    .\Provision-MgUserAccount.ps1 -UserCsvFilePath ".\Users.csv" -Password "SecureP@ssw0rd" -AssignedLicenseSkus "ENTERPRISEPACK"
    Creates users with a static password and directly assigns the specified license.

.EXAMPLE
    .\Provision-MgUserAccount.ps1 -UserCsvFilePath ".\users.csv" -AutoGeneratePassword -PromptForAssignedLicenseSkus -PromptForDisableLicensePlans -SamePasswordForAll
    Generates one password for all users and prompts interactively for license and plan choices.

.EXAMPLE
    .\Provision-MgUserAccount.ps1 -UserCsvFilePath "C:\Users.csv" -AutoGeneratePassword
    Creates users using unique, auto-generated passwords with no license assignment.

.EXAMPLE
    .\Provision-MgUserAccount.ps1 -UserCsvFilePath "C:\Users.csv" -AssignedLicenseSkus "ENTERPRISEPACK" -DisableLicensePlans "EXCHANGE_S_ENTERPRISE"
    Assigns a license and disables the Exchange Online service plan.

.EXAMPLE
    .\Provision-MgUserAccount.ps1 -UserCsvFilePath "C:\Users.csv" -PromptForAssignedLicenseSkus -PromptForDisableLicensePlans
    Opens a GUI to select licenses and plans to disable.

.INPUTS
    CSV file input with user data.

.OUTPUTS
    CSV file containing success/failure details per user. Optional log file output.

.NOTES
    Ensure the Microsoft Graph module is installed and the script is authenticated to Microsoft 365 with required permissions.

.FUNCTIONALITY
    Bulk user provisioning and license assignment in Microsoft 365
#>

# Parameters for the script
[CmdletBinding(DefaultParameterSetName = 'ManualPassword', SupportsShouldProcess = $true)]
param (
    # **User Input Parameters**
    [Parameter(Mandatory, Position = 0)]
    [Alias('UserIds', 'UserSource')]
    [ValidateScript({Test-Path $_ -PathType Leaf})]
    [string]$UserIdsCsv,

    [Parameter(Position = 1)]
    [string]$ResultExportFilePath,

    # **License Assignment (Mutually Exclusive)**
    [Parameter(ParameterSetName = "DirectLicense")]
    [Alias('LicenseSkus')]
    [string[]]$AssignLicenses,

    [Parameter(ParameterSetName = "PromptForLicense")]
    [Alias('PromptForAssignedLicenseSkus')]
    [switch]$SelectLicenses,

    # **Service Plan Management**
    [Parameter(ParameterSetName = "DirectLicense")]
    [Parameter(ParameterSetName = "DirectDisablePlans")]
    [Alias('DisableLicensePlans')]
    [string[]]$DisablePlans,

    [Parameter(ParameterSetName = "PromptForLicense")]
    [Parameter(ParameterSetName = "PromptForDisablePlans")]
    [Alias('PromptForDisableLicensePlans')]
    [switch]$SelectDisabledPlans,

    # **Password Configuration (Mutually Exclusive)**
    [Parameter(ParameterSetName = "ManualPassword")]
    [ValidateNotNullOrEmpty()]
    [string]$Password,

    [Parameter(ParameterSetName = "AutoPassword")]
    [switch]$AutoGeneratePassword,

    [Parameter(ParameterSetName = "AutoPassword")]
    [ValidateRange(10, 20)]
    [int]$PasswordLength = 12,

    [Parameter(ParameterSetName = "AutoPassword")]
    [switch]$SamePasswordForAll,

    # **Common Options**
    [Parameter()]
    [bool]$ForcePasswordChange = $true,

    [Parameter()]
    [string]$LogFilePath
)


################## Variables ###################
$createUserAccounts = @()
$addLicenses = @()
$licenseParams = $null  # Initialize to avoid undefined return
$ProcessUsers = @() 
if (-not $LogFilePath) {
    $LogFilePath = Join-Path $PSScriptRoot "UserCreationLogInf_$(Get-Date -Format 'yyyyMMdd').log"
}


# Function to write log entries
function Write-Log {
    param (
        [string]$Message,
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $lineNumber = $MyInvocation.ScriptLineNumber
    $entry = "[$timestamp][$Level][Line $lineNumber] $Message"

    Add-Content -Path $LogFilePath -Value $entry
    Write-Verbose $entry
}


#################### Parameter and Path Validation

# Validate mutual exclusivity (though parameter sets should prevent this)
if ($PSBoundParameters.ContainsKey('AssignLicenses') -and $SelectLicenses) {
    Write-Log "Parameters -AssignedLicenseSkus and -PromptForAssignedLicenseSkus are mutually exclusive and cannot be used together." -Level "ERROR"
    throw "Parameters -AssignedLicenseSkus and -PromptForAssignedLicenseSkus are mutually exclusive and cannot be used together."
}

if ($PSBoundParameters.ContainsKey('Password') -and $AutoGeneratePassword) {
    Write-Log "Parameters -Password and -AutoGeneratePassword are mutually exclusive and cannot be used together." -Level "ERROR"
    throw "Parameters -Password and -AutoGeneratePassword are mutually exclusive and cannot be used together."
}

# Validate if the CSV file exists, import it, and check for required columns
if (-not (Test-Path $UserIdsCsv)) {
    Write-Log "The specified CSV file '$UserIdsCsv' does not exist." "ERROR"
    return
}



# Define retry logic with throttling check
function Invoke-WithRetry {
    param (
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$DelayBetweenRetries = 5
    )
    
    $attempt = 0
    $success = $false
    while ($attempt -lt $MaxRetries -and -not $success) {
        try {
            $attempt++
            return & $ScriptBlock
        }
        catch {
            # Check if the error is throttling (HTTP 429) or bandwidth exceeded (HTTP 509)
            if ($_.Exception -match "HTTP 429|HTTP 509") {
                Write-Log "Attempt $attempt failed due to error $_. Retrying in $DelayBetweenRetries seconds..." "INFO"
                Start-Sleep -Seconds $DelayBetweenRetries
            }
            else {
                # If it's not a throttling or bandwidth exceeded error, log it and stop retrying
                if ($attempt -ge $MaxRetries) {
                    Write-Log "Max retries reached. Giving up." "ERROR"
                    throw $_  # Rethrow the last error after max retries
                }
                break  # Exit loop for non-throttling errors
            }
        }
    }
}

# Password generator
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

function Export-Results {
    param (
        [Parameter(Mandatory = $true)]
        [Object]$ResultData,

        [string]$ResultExportPath,

        [string]$newFileName # Optional custom file name (without extension or date)
    )

    $exportDate = Get-Date -Format 'yyyy_MM_dd_HH'

    # Use default base name if not provided
    $baseName = if ([System.IO.Path]::GetFileNameWithoutExtension($ResultExportPath)) {
        [System.IO.Path]::GetFileNameWithoutExtension($ResultExportPath)
    }else {
        [System.IO.Path]::GetFileNameWithoutExtension($newFileName)
    }

    # Strip any extension and append timestamp + .csv
    $finalFileName = "$baseName-$exportDate.csv"

    if ($ResultExportPath) {
        # Extract directory path or fallback to PSScriptRoot if it's just a filename
        $directoryPath = [System.IO.Path]::GetDirectoryName($ResultExportPath)
        if ([string]::IsNullOrWhiteSpace($directoryPath)) {
            $directoryPath = $PSScriptRoot
        }
    } else {
        $directoryPath = $PSScriptRoot
    }

    # Construct final export path
    $finalFilePath = [System.IO.Path]::Combine($directoryPath, $finalFileName)

    # If file already exists, append 6-digit random number to avoid overwrite
    if (Test-Path $finalFilePath) {
        $randomSuffix = Get-Random -Minimum 100000 -Maximum 1000000
        $finalFileName = "$baseName-$exportDate-$randomSuffix.csv"
        $finalFilePath = [System.IO.Path]::Combine($directoryPath, $finalFileName)
    }

    # Export the data
    $ResultData | Export-Csv -Path $finalFilePath -NoTypeInformation -Force

    # Log export result
    Write-Log "Results exported to $finalFilePath." "INFO"
}


#################### Password Processing and Configurations ###################


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


#################  Check and Correct Object Properties if DisplayName or MailNickname value is missing ####################

Write-Host "`nStep 1: Starting import of the specified CSV file containing user account data..." -ForegroundColor Cyan


$users = Import-Csv -Path $UserIdsCsv

if ($users.Count -eq 0) {
    Write-Log "The CSV file '$UserIdsCsv' provided is empty." "ERROR"
    return
}

################# Validation csv data #######################

Write-Host "`nStep 2: Validating and processing user properties 'UserPrincipalName', 'DisplayName', and 'MailNickname' retrieved from the imported CSV file..." -ForegroundColor Yellow

# Validate required columns (UserPrincipalName, DisplayName, and MailNickname)
$requiredColumns = @("UserPrincipalName", "DisplayName", "MailNickname")
foreach ($col in $requiredColumns) {
    if (-not ($users[0].PSObject.Properties.Name -contains $col)) {
        Write-Log "The CSV file is missing the required column '$col'." "ERROR"
        return
    }
}

Write-Log "Starting user validation... Checking 'UserPrincipalName', 'DisplayName', and 'MailNickname'. If 'GivenName' and 'Surname' exist, they'll be used to generate missing fields." "INFO"

foreach ($user in $users) {
    $hasGiven  = $user.PSObject.Properties.Name -contains "GivenName"
    $hasSurname = $user.PSObject.Properties.Name -contains "Surname"

    # Validate UPN format early
    $isValidUPN = $user.UserPrincipalName -match "^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$"
    if (-not $user.UserPrincipalName -or -not $isValidUPN) {
        $user | Add-Member -MemberType NoteProperty -Name "ValidationStatus" -Value "Invalid" -Force
        Write-Log "Invalid or missing UPN: $($user | ConvertTo-Json -Compress)" "WARN"
        $ProcessUsers += $user
        continue
    }

    # Generate DisplayName if missing
    if (-not $user.DisplayName) {
        switch -Wildcard ($true) {
            { $hasGiven -and $hasSurname -and $user.GivenName -and $user.Surname } { $user.DisplayName = "$($user.GivenName) $($user.Surname)"; break }
            { $hasGiven -and $user.GivenName } { $user.DisplayName = $user.GivenName; break }
            { $hasSurname -and $user.Surname } { $user.DisplayName = $user.Surname; break }
            default {
                Write-Log "Missing DisplayName and insufficient data to generate it for UPN: $($user.UserPrincipalName)" "WARN"
            }
        }
    }

    # Generate MailNickName if missing
    if (-not $user.MailNickName) {
        switch -Wildcard ($true) {
            { $user.DisplayName } { $user.MailNickName = $user.DisplayName.Replace(" ", "."); break }
            { $hasGiven -and $hasSurname -and $user.GivenName -and $user.Surname } { $user.MailNickName = "$($user.GivenName).$($user.Surname)"; break }
            { $hasGiven -and $user.GivenName } { $user.MailNickName = $user.GivenName; break }
            { $hasSurname -and $user.Surname } { $user.MailNickName = $user.Surname; break }
            { $user.UserPrincipalName } { $user.MailNickName = $user.UserPrincipalName.Split('@')[0]; break }
            default {
                Write-Log "Missing MailNickName and insufficient data to generate it for UPN: $($user.UserPrincipalName)" "WARN"
            }
        }
    }

    # Final validation for required fields
    if (-not $user.DisplayName -or -not $user.MailNickName) {
        $user | Add-Member -MemberType NoteProperty -Name "ValidationStatus" -Value "Invalid" -Force
        Write-Log "Incomplete user record for UPN: $($user.UserPrincipalName)" "WARN"
    } else {
        $user | Add-Member -MemberType NoteProperty -Name "ValidationStatus" -Value "Valid" -Force
    }

    $ProcessUsers += $user
}


# Split valid and invalid users
$validProcessedUsers = $ProcessUsers | Where-Object { $_.ValidationStatus -eq "Valid" } | Select-Object * -ExcludeProperty ValidationStatus
$invalidProcessedUsers = $ProcessUsers | Where-Object { $_.ValidationStatus -eq "Invalid" }

Write-Log "Processed $($users.Count) users. Valid: $($validProcessedUsers.Count), Invalid: $($invalidProcessedUsers.Count). All processed object has exported" "INFO"
Export-Results -ResultData $ProcessUsers -ResultExportPath $ResultExportFilePath -newFileName "Detailed_ProcessUser"

if (!$validProcessedUsers) {
    Write-Log "There are not valid account and unable to proceed with the account creation, verify the CSV file, especially the mandatory columns 'UserPrincipalName', 'DisplayName', and 'MailNickname' "
    return
}

##################################### User license processing ####################################

Write-Host "`nStep 3: Retrieving available Microsoft 365 licenses and evaluating applicable service plans..." -ForegroundColor Magenta


# Handle License Assignment (Mutually Exclusive)
if ($AssignLicenses) {
    Write-Log "AssignLicenses: $AssignLicenses"
    $AssignedSku = Get-MgSubscribedSku -All | Where-Object { $_.SkuPartNumber -in $AssignLicenses -or $_.SkuId -in $AssignLicenses }

    if (-not $AssignedSku) {
        Write-Warning "No matching SKUs found. User will be created without licenses."
    }
}
elseif ($SelectLicenses) {
    Write-Log "Please select SKUs from the list below."
    $AssignedSku = Get-MgSubscribedSku -All | Where-Object { ($_.PrepaidUnits.Enabled -ne $_.ConsumedUnits) } | Out-GridView -PassThru -Title "Select License SKUs to Assign"

    if (-not $AssignedSku) {
        Write-Warning "No license selected. Continuing without license assignment."
    }
}

# Handle Disabled Plans
if ($SelectDisabledPlans -and $AssignedSku) {
    Write-Log "Please select service plans to disable from the list below."
    $SelectedDisableLicensePlans = $AssignedSku.ServicePlans | Where-Object { $_.AppliesTo -eq "User" -and $_.ProvisioningStatus -eq "Success" } | Out-GridView -PassThru -Title "Select Service Plans to Disable"
    [Object]$DisablePlans = $SelectedDisableLicensePlans | Select-Object ServicePlanId, ServicePlanName
}

# Build License
if ($AssignedSku) {
    foreach ($sku in $AssignedSku) {
        $licenseAssignment = @{ skuId = $sku.SkuId }

        if ($DisablePlans) {
            # Log the processing of service plans for the current SKU
            Write-Log "Processing the plans for SKU: $($sku.SkuPartNumber)" "WARN"
            
            # Handle disabling service plans based on the prompt flag
            $disabledPlans = if ($SelectDisabledPlans) {
                $sku.ServicePlans | Where-Object { $_.ServicePlanId -in $DisablePlans.ServicePlanId -or $_.ServicePlanName -in $DisablePlans.ServicePlanName }
            }
            else {
                $sku.ServicePlans | Where-Object { $_.ServicePlanId -in $DisablePlans -or $_.ServicePlanName -in $DisablePlans }
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


############################# Object Creation and license assignment #################################
Write-Host "`nStep 4: Creating user accounts and assigning licenses where available. If no license is found, only the user account will be created...`n" -ForegroundColor Green

$counter = 0
$totalUsers = $validProcessedUsers.Count

foreach ($user in $validProcessedUsers) {

    $newUserParams = @{}
    $result = @{}
    $counter++

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
            # Create a new user with retry logic
            try {
                $newUser = Invoke-WithRetry {
                    try {
                        return (New-MgUser @newUserParams -ErrorAction Stop)
                    }
                    catch {
                        # Log the specific error, instead of showing in the console
                        Write-Log "Failed to create user $($newUserParams.UserPrincipalName): $($_.ErrorDetails.Message.split("`n")[0])" "ERROR"
                        throw $_  
                    }
                }

                if ($newUser) {
                    Write-Log "Successfully created user $($newUser.UserPrincipalName)" -Level "INFO"
                }
            }
            catch {
                # Log error if user creation fails
                Write-Log "Failed to create user $($newUserParams.UserPrincipalName): $_" "ERROR"
            }
            

            if ($licenseParams -and $newUser) {
                Invoke-WithRetry {
                    Invoke-MgGraphRequest -Method POST -Uri "https://graph.microsoft.com/v1.0/users/$($newUser.Id)/assignLicense" -Body $licenseParams -ContentType "application/json" | Out-Null
                }

                $result['AssignLicenses'] = $AssignedSku.SkuPartNumber -join ','
                $result['AssignedLicenseId'] = $AssignedSku.SkuId -join ','

                if ($DisablePlans) {
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

    # Report progress every 10 users
    if ($counter % 10 -eq 0 -or $counter -eq $totalUsers) {
        Write-Host "`tProgress: $counter of $totalUsers users processed.`n" -ForegroundColor DarkCyan
    }

    $createUserAccounts += [PSCustomObject]$result
}


# Export results if any
if ($createUserAccounts.Count -gt 0) {
    Export-Results -ResultData $createUserAccounts -ResultExportPath $ResultExportFilePath -newFileName "CreatedUserAccountInfo"
}
else {
    Write-Warning "No user accounts to created"
}

Write-Log "User creation script completed. $(($createUserAccounts | Where-Object Status -eq 'Success').Count) succeeded, $(($createUserAccounts | Where-Object Status -eq 'Failed').Count) failed."
