# 📘 Microsoft 365 User Management Scripts

This repository contains two PowerShell scripts for managing Microsoft 365 users via Microsoft Graph API:

1. **`Perform-MgUserLicenseUpgrade`** - Migrates users between license SKUs with optional service plan management.
2. **`Provision-MgUserAccount`** - Provisions new users with optional license assignment and password management.

---

## 📦 `Perform-MgUserLicenseUpgrade`

### 📝 Overview

Migrates Microsoft 365 users from one license SKU to another, supporting optional service plan disabling, interactive GUI selection, and CSV bulk processing.

### 🚀 Features

- Migrate users between SKUs (e.g., E3 → E5)
- Disable specific service plans (Exchange, Yammer, etc.)
- Process single UPNs or bulk CSV imports
- Interactive license picker (`Out-GridView`)
- Retain existing disabled plans (optional)
- Comprehensive logging

### 📌 Parameters

| Parameter               | Description |
|-------------------------|-------------|
| **-UserIds**            | UPN(s) or CSV path(s) containing users (supports mixed input like `"user@contoso.com", "path/to/file.csv"`) |
| **-LicenseToRemove**    | SKU to remove (e.g., `SPE_E3`) |
| **-LicenseToAdd**       | SKU to assign (e.g., `SPE_E5`) |
| **-DisabledPlans**      | Service plan GUIDs to disable |
| **-SelectLicenses**  | Interactive license selection (makes `-LicenseToRemove/Add` optional) |
| **-SelectDisabledPlans** | Interactive plan disabling |
| **-KeepExistingPlanState** | Preserve existing disabled plans during migration |

### 💡 Examples

```powershell
# Single user migration
Perform-MgUserLicenseUpgrade -UserIds "user@contoso.com" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5"

# Bulk CSV migration with plan disabling
Perform-MgUserLicenseUpgrade -UserIds "users.csv" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5" -DisabledPlans "e95bec33-7c88..."

# Interactive mode
Perform-MgUserLicenseUpgrade -UserIds "users.csv" -SelectLicenses -SelectDisabledPlans

# Mixed input processing (UPNs + CSV)
Perform-MgUserLicenseUpgrade -UserIds "user1@contoso.com", "user2@contoso.com", ".\more_users.csv" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5"

# Interactive mode with mixed inputs
Perform-MgUserLicenseUpgrade -UserIds "user@contoso.com", ".\departments\sales.csv" -SelectLicenses -SelectDisabledPlans

# Bulk migration with plan retention
Perform-MgUserLicenseUpgrade -UserIds ".\all_users.csv" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5" -KeepExistingPlanState
```

---

## 📘 `Provision-MgUserAccount`

### 📄 Overview

Automates Microsoft 365 user provisioning with CSV support, license assignment, and password management.

### 🚀 Features

- Bulk user creation from CSV
- Auto/preset password generation
- License assignment (direct or interactive)
- Service plan disabling
- Validation and error handling
- Detailed logging

### 📌 Parameters

| Parameter | Description |
|-----------|-------------|
| -UserIdsCsv | Path to user CSV file |
| -Password | Static password for all users |
| -AutoGeneratePassword | Auto-generate passwords |
| -AutoPasswordLength | Password length (10-20 characters) |
| -SamePasswordForAll | Use same password for all users |
| -ForcePasswordChange | Force password change at first login |
| -AssignedLicense | License SKUs to assign |
| -SelectLicenses | Interactive license selection |
| -DisablePlans | Plans to disable |
| -SelectDisabledPlans | Interactive plan disabling |
| -ResultExportFilePath | Custom path for results export |
| -LogFilePath | Custom path for log file |

### 💡 Examples

```powershell
# Basic provisioning with auto passwords
.\Provision-MgUserAccount.ps1 -UserIdsCsv "users.csv" -AutoGeneratePassword

# With license assignment
.\Provision-MgUserAccount.ps1 -UserIdsCsv "users.csv" -AssignedLicense "ENTERPRISEPACK"

# Interactive mode
.\Provision-MgUserAccount.ps1 -UserIdsCsv "users.csv" -SelectLicenses -SelectDisabledPlans

# Auto-fill demonstration with minimal CSV
.\Provision-MgUserAccount.ps1 -UserIdsCsv ".\minimal_users.csv" -AutoGeneratePassword
# CSV only needs: UserPrincipalName
# Script auto-generates: DisplayName, MailNickname

# Comprehensive example
.\Provision-MgUserAccount.ps1 -UserIdsCsv ".\full_users.csv" `
    -AutoGeneratePassword -AutoPasswordLength 14 -SamePasswordForAll `
    -SelectLicenses -SelectDisabledPlans `
    -ResultExportFilePath ".\reports\Q3_onboarding.csv"
```

---

## 📂 Common Prerequisites

- **Modules**:  
  `Microsoft.Graph.Users`, `Microsoft.Graph.Identity.DirectoryManagement`, `Microsoft.Graph.Authentication`
  
- **Installation**:  

  ```powershell
  Install-Module Microsoft.Graph -Scope CurrentUser
  ```

- **Authentication**:  

  ```powershell
  Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"
  ```

- **Permissions**:  
  - User management rights
  - License management rights

---

## 📄 CSV Formats

### User Provisioning CSV

```csv
UserPrincipalName,DisplayName,MailNickname,GivenName,Surname
user1@contoso.com,User One,user1,User,One
```

### License Migration CSV

```csv
UserPrincipalName
user1@contoso.com
user2@contoso.com
```

---

## 📁 Output & Logging

Both scripts generate:

- Timestamped log files
- Detailed operation reports
- Error tracking

Default Log locations:

- Script execution directory

---

## 🔒 Security Best Practices

- Use least-privilege permissions
- Avoid hardcoding credentials
- Store passwords securely
- Test with small batches first

---

## 📚 References

- [Microsoft Graph API Documentation](https://learn.microsoft.com/en-us/graph/)
- [Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [Microsoft 365 Licensing Guide](https://learn.microsoft.com/en-us/microsoft-365/)
