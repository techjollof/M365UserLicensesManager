# Microsoft 365 User Provisioning and License Management

## 📑 Table of Contents

- [📘 Synopsis](#-synopsis)
- [🧩 Available Scripts](#-available-scripts)
  - [📦 `About Perform-MgUserLicenseUpgrade.ps1`](#-about-perform-mguserlicenseupgradeps1)
    - [📝 Overview](#-overview)
    - [🚀 Key Features](#-key-features)
    - [📌 Parameters](#-parameters)
    - [💡 Usage Examples](#-usage-examples)
  - [👥 `Provision-MgUserAccount.ps1`](#-provision-mguseraccountps1)
    - [📄 Overview](#-overview-1)
    - [🚀 Key Features](#-key-features-1)
    - [📌 Parameters](#-parameters-1)
    - [💡 Usage Examples](#-usage-examples-1)
- [📂 CSV File Formats](#-csv-file-formats)
- [📁 Output & Logging](#-output--logging)
- [🔐 Prerequisites & Permissions](#-prerequisites--permissions)
  - [🧰 Required Modules](#-required-modules)
  - [🔑 Authentication](#-authentication)
  - [📋 Required Graph Permissions](#-required-graph-permissions)
- [🔒 Security Best Practices](#-security-best-practices)
- [📚 References](#-references)

## 📘 Synopsis

This repository provides **PowerShell-based automation tools** for managing Microsoft 365 users using the **Microsoft Graph API**. The toolkit includes two feature-rich scripts designed to streamline **user provisioning** and **license management** for both small-scale and enterprise environments.

## 🧩 Available Scripts

1. **`Perform-MgUserLicenseUpgrade.ps1`** – Automates Microsoft 365 **license migrations**, with optional service plan adjustments, interactive selection, and bulk CSV support.

2. **`Provision-MgUserAccount.ps1`** – Creates **new user accounts** with optional license assignment, automated password generation, and full lifecycle logging.

## 📦 `About Perform-MgUserLicenseUpgrade.ps1`

### 📝 Overview

This script simplifies **license SKU transitions** (e.g., migrating users from E3 to E5), with optional service plan toggling. It supports both **interactive** and **CSV-based** user selection and offers **logging and plan preservation** capabilities for enterprise-grade compliance.

### 🚀 Key Features

- Seamlessly migrate users between license SKUs (e.g., `SPE_E3 → SPE_E5`)
- Disable specific **service plans** (e.g., Exchange, Yammer)
- Supports **single UPNs** or **bulk CSV inputs**
- **Interactive license and plan selection** via `Out-GridView`
- Option to **retain existing disabled service plans**
- **Detailed logging** and exportable result summaries

### 📌 Parameters

| Parameter                   | Description |
|-|-|
| `-UserIds`                 | One or more UPNs or path(s) to CSV file(s) containing users |
| `-LicenseToRemove`         | License SKU ID to remove (e.g., `SPE_E3`) |
| `-LicenseToAdd`            | License SKU ID to assign (e.g., `SPE_E5`) |
| `-DisabledPlans`           | List of service plan GUIDs to disable |
| `-SelectLicenses`          | Launches an interactive license selector |
| `-SelectDisabledPlans`     | Launches an interactive plan selection UI |
| `-KeepExistingPlanState`   | Preserves currently disabled service plans |

### 💡 Usage Examples

#### ✅ **Example 1: Single User License Upgrade**

Migrate a single user from E3 to E5 by specifying both licenses.

```powershell
Perform-MgUserLicenseUpgrade -UserIds "user@contoso.com" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5"
```

> Ideal for testing or handling specific user upgrades manually.

#### ✅ **Example 2: Bulk CSV Migration with Disabled Plans**

 Migrates all users listed in `users.csv`, removes their E3 license, assigns E5, and disables specified service plans (using GUIDs).

```powershell
Perform-MgUserLicenseUpgrade -UserIds "users.csv" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5" -DisabledPlans "e95bec33-7c88..."
```

> Perfect for migrating departments or batches of users while selectively disabling unused services like Yammer or Stream.

#### ✅ **Example 3: Fully Interactive Mode**

 Load users from a CSV and choose licenses and service plans via a GUI (`Out-GridView`).

```powershell
Perform-MgUserLicenseUpgrade -UserIds "users.csv" -SelectLicenses -SelectDisabledPlans
```

> Best when you're unsure of exact SKU IDs or plan GUIDs — reduces chances of error.

#### ✅ **Example 4: Mixed Input (UPNs + CSV)**

 Process a mix of direct UPNs and one or more CSV files.

```powershell
Perform-MgUserLicenseUpgrade -UserIds "user1@contoso.com", "user2@contoso.com", ".\more_users.csv" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5"
```

> Lets you handle ad-hoc migrations without preprocessing data into a single CSV file.

#### ✅ **Example 5: Mixed Inputs with Interactive Mode**

 Use both manual UPNs and a CSV list with interactive license and plan selection.

```powershell
Perform-MgUserLicenseUpgrade -UserIds "user@contoso.com", ".\departments\sales.csv" -SelectLicenses -SelectDisabledPlans
```

> A flexible option when some users need manual review while others are listed in CSVs.

#### ✅ **Example 6: Preserve Existing Plan State**

 Migrate licenses while keeping current disabled service plans untouched.

```powershell
Perform-MgUserLicenseUpgrade -UserIds ".\all_users.csv" -LicenseToRemove "SPE_E3" -LicenseToAdd "SPE_E5" -KeepExistingPlanState
```

> Maintains user-specific plan customizations (like disabled Teams or Yammer) through license changes.

📄 **More Details**: [Perform-MgUserLicenseUpgrade.md](./Perform-MgUserLicenseUpgrade.md)

## 👥 `Provision-MgUserAccount.ps1`

### 📄 Overview

Automates the creation of new Microsoft 365 user accounts, including **password setup**, **license assignment**, and **plan configuration**, with full **CSV-based or interactive support**. Designed for high-volume onboarding and consistent policy application.

### 🚀 Key Features

- Bulk user creation via **CSV import**
- **Password generation** (static or dynamic)
- Assign licenses via **parameters or interactive UI**
- Disable specific service plans as needed
- Supports **input validation**, logging, and export
- Auto-fills missing fields (e.g., `DisplayName`, `MailNickname`) if only `UserPrincipalName` is provided

### 📌 Parameters

| Parameter                | Description |
|--|-|
| `-UserIdsCsv`           | Path to CSV file containing user data |
| `-Password`             | Optional static password for all users |
| `-AutoGeneratePassword` | Enables dynamic password generation |
| `-AutoPasswordLength`   | Length for auto-generated passwords (10–20 chars) |
| `-SamePasswordForAll`   | Apply one generated password to all users |
| `-ForcePasswordChange`  | Forces password change on first login |
| `-AssignedLicense`      | License SKUs to assign directly |
| `-SelectLicenses`       | Opens license picker GUI |
| `-DisablePlans`         | GUIDs of service plans to disable |
| `-SelectDisabledPlans`  | Opens interactive service plan picker |
| `-ResultExportFilePath` | Path to export success/failure report |
| `-LogFilePath`          | Path to store execution logs |

### 💡 Usage Examples

#### ✅ **Example 1: Auto-Generate Passwords**

 Creates users from a CSV with automatically generated passwords.

```powershell
.\Provision-MgUserAccount.ps1 -UserIdsCsv "users.csv" -AutoGeneratePassword
```

> Useful for secure onboarding without handling passwords manually.

#### ✅ **Example 2: License Assignment with Provisioning**

 Assigns a license during user creation.

```powershell
.\Provision-MgUserAccount.ps1 -UserIdsCsv "users.csv" -AssignedLicense "ENTERPRISEPACK"
```

> When license assignment is mandatory during user setup (e.g., E1, E3).

#### ✅ **Example 3: Fully Interactive Setup**

 Select license SKUs and plans interactively while importing from CSV.

```powershell
.\Provision-MgUserAccount.ps1 -UserIdsCsv "users.csv" -SelectLicenses -SelectDisabledPlans
```

> Enables visual selection without needing to know SKU or plan IDs.

#### ✅ **Example 4: Minimal Input with Auto-Fill**

 Provision users from a CSV that only includes UPNs. Other fields like `DisplayName` and `MailNickname` will be auto-filled.

```powershell
.\Provision-MgUserAccount.ps1 -UserIdsCsv ".\minimal_users.csv" -AutoGeneratePassword
```

> Reduces complexity when only usernames are available. Great for quick trials or onboarding from minimal data exports.

#### ✅ **Example 5: Full Automation with Export**

 Complete setup: generate passwords, select licenses, select disabled plans, and export results to a report file.

```powershell
.\Provision-MgUserAccount.ps1 -UserIdsCsv ".\full_users.csv" `
    -AutoGeneratePassword -AutoPasswordLength 14 -SamePasswordForAll `
    -SelectLicenses -SelectDisabledPlans `
    -ResultExportFilePath ".\reports\Q3_onboarding.csv"
```

> End-to-end onboarding scenario with custom report path and all options enabled.

📄 **Full details**: [Provision-MgUserAccount.md](./Provision-MgUserAccount.md)

## 📂 CSV File Formats

### 👤 User Provisioning

```csv
UserPrincipalName,DisplayName,MailNickname,GivenName,Surname
user1@contoso.com,User One,user1,User,One
```

### 🔄 License Migration

```csv
UserPrincipalName
user1@contoso.com
user2@contoso.com
```

## 📁 Output & Logging

Both scripts automatically generate:

- **Timestamped log files**  
- **Exportable result summaries** (CSV)  
- **Detailed error tracking**  

> Logs are saved to the script directory by default unless a custom path is provided.

## 🔐 Prerequisites & Permissions

### 🧰 Required Modules

```powershell
Install-Module Microsoft.Graph -Scope CurrentUser
```

### 🔑 Authentication

```powershell
Connect-MgGraph -Scopes "User.ReadWrite.All", "Directory.ReadWrite.All"
```

### 📋 Required Graph Permissions

- `User.ReadWrite.All`
- `Directory.ReadWrite.All`

## 🔒 Security Best Practices

- Use **least-privilege accounts** for execution
- Avoid storing or hardcoding passwords
- Secure CSVs and logs with appropriate access controls
- Always test in a **non-production environment** first

## 📚 References

- [📖 Microsoft Graph API Docs](https://learn.microsoft.com/en-us/graph/)
- [📘 Microsoft Graph PowerShell SDK](https://learn.microsoft.com/en-us/powershell/microsoftgraph/)
- [📋 Microsoft 365 Licensing Guide](https://learn.microsoft.com/en-us/microsoft-365/)
