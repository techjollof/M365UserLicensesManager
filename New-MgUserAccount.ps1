# function New-MgUserAccount {
[CmdletBinding()]
param (
    [Parameter(Mandatory, Position = 0)]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$UserCsvFilePath,
        
    [Parameter(Mandatory, Position = 1)]
    [string]$ResultExportFilePath,

    [Parameter()]
    [switch]$AssignLicense,
    
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LicenseSku,

    [Parameter(ParameterSetName = "ManualPassword")]
    [ValidateNotNullOrEmpty()]
    [string]$Password,

    [Parameter(ParameterSetName = "AutoPassword")]
    [switch]$AutoPassword,
        
    [Parameter(ParameterSetName = "AutoPassword")]
    [ValidateRange(10, 20)]
    [int]$AutoPasswordLength = 12,

    [Parameter(ParameterSetName = "AutoPassword")]
    [switch]$SamePasswordForAll
)

# Import the CSV file
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
        $password = $password | Sort-Object { Get-Random } # Shuffle
    }
    else {
        $password = 1..$Length | ForEach-Object { $allChars | Get-Random }
    }
    
    return -join $password
}

# Create a password profile
if ($AutoPassword) {
    $Password = New-RandomPassword -Length $AutoPasswordLength
}
elseif (-not $Password) {
    Throw "Password must be defined for account creation, or you use the -AutoPassword to automatically generate password for the account. You can also add -SamePasswordForAll and all users will be created with the same password."
}

$PasswordProfile = @{
    Password                      = $Password
    ForceChangePasswordNextSignIn = $true
}

# Initialize an array to store user creation results
$results = @()

# Assign a license to the new user
if ($AssignLicense) {
    if (-not $LicenseSku) {
        $AssignLicenseSku = Get-MgSubscribedSku -All | Out-GridView -PassThru -Title "Select a License SKU"
        if (-not $AssignLicenseSku) {
            Write-Warning "No license selected. Continuing without license assignment."
            $AssignLicense = $false
        }
    }
    else {
        $AssignLicenseSku = Get-MgSubscribedSku -All | Where-Object SkuPartNumber -eq $LicenseSku
        
        if (-not $AssignLicenseSku) {
            Write-Warning "License SKU '$LicenseSku' not found. Continuing without license assignment."
            $AssignLicense = $false
        }
    }
}

# Loop through each user in the CSV file
foreach ($user in $users) {
    try {
        $newUserParams = @{}
        $user.PSObject.Properties | ForEach-Object {
            if ($_.Value) {
                $newUserParams[$_.Name] = $_.Value
            }
        }

        if ($AutoPassword -and $SamePasswordForAll -eq $false) {
            $Password = New-RandomPassword -Length $AutoPasswordLength
            $PasswordProfile.Password = $Password
        }

        # Add required parameters
        $newUserParams['PasswordProfile'] = $PasswordProfile
        $newUserParams['AccountEnabled'] = $true

        # Create a new user
        # $newUser = New-MgUser @newUserParams
        Write-Host "Successfully created user $($newUser.UserPrincipalName)" -ForegroundColor Green

        Start-Sleep -Seconds 3

        if ($AssignLicense -and $LicenseSku) {
            $licenseParams = @{
                UserId         = $newUser.Id
                AddLicenses    = @( @{ SkuId = $LicenseSku.SkuId } )
                RemoveLicenses = @()
            }
            # Set-MgUserLicense @licenseParams
            Write-Host "Assigned license $($LicenseSku.SkuPartNumber) to user $($newUser.UserPrincipalName)" -ForegroundColor Cyan
        }

        # Add success result to the results array
        $result = $user.PSObject.Copy()
        $result | Add-Member -MemberType NoteProperty -Name 'Status' -Value 'Success'
        $result | Add-Member -MemberType NoteProperty -Name 'AccountPassword' -Value $Password
        $result | Add-Member -MemberType NoteProperty -Name 'UserId' -Value $newUser.Id
        if ($AssignLicense -and $UserSku) {
            $result | Add-Member -MemberType NoteProperty -Name 'AssignedLicense' -Value $UserSku.SkuPartNumber
        }
    }
    catch {
        # Handle errors
        Write-Error "Failed to create user $($user.UserPrincipalName): $_"
        $result = $user.PSObject.Copy()
        $result | Add-Member -MemberType NoteProperty -Name 'Status' -Value 'Failed'
        $result | Add-Member -MemberType NoteProperty -Name 'ErrorMessage' -Value $_.Exception.Message
    }

    # Add the user result to the results array
    $results += $result
}

# Export the results to a CSV file
if ($results) {
    $results | Export-Csv -Path $ResultExportFilePath -NoTypeInformation -Force
    Write-Host "Results exported to $ResultExportFilePath" -ForegroundColor Green
}
else {
    Write-Warning "No results to export"
}
    
return $results
# }