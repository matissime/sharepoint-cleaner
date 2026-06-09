# SharePoint Version Cleaner (SPOVC)

A modular PowerShell 7+ GUI application for managing file version cleanup in SharePoint Online document libraries. Provides scan, dry-run, and cleanup modes with HTML session reporting and comprehensive logging.

## 🎯 Key Features

- **Single & Batch Operations**: Cleanup individual sites/libraries or entire tenant
- **Configurable Retention**: Externalized JSON config (default: keep 25 versions)
- **Windows Forms GUI**: Real-time progress tracking with cancellation support
- **Session Reporting**: HTML reports with per-site metrics and storage savings
- **JWT Token Caching**: Efficient authentication with fallback to interactive auth
- **Modular Architecture**: Clean separation of concerns for maintainability and testing

## 📁 Project Structure

```
SharePoint-Version-Cleaner/
├── SPOVC-Modular.ps1                 ← Main entry point (new modular version)
├── config.json                       ← External configuration (no more hardcoding!)
│
├── Modules/                          ← Modular PowerShell architecture
│   ├── Core/
│   │   ├── Core.psm1                 (Logging, utilities, session state)
│   │   └── Config.psm1               (Configuration management)
│   │
│   ├── Authentication/
│   │   └── Authentication.psm1       (JWT tokens, SharePoint connections)
│   │
│   ├── Reporting/
│   │   └── Reporting.psm1            (Session statistics, HTML reports)
│   │
│   ├── UI/
│   │   └── UI.psm1                   (Windows Forms dialogs)
│   │
│   ├── Processing/
│   │   └── Processing.psm1           (Job orchestration, background processing)
│   │
│   └── Processing/
│       └── SharePointOps.psm1        (SharePoint API operations, version management)
│
├── Logs/                             ← Runtime logs & session reports (auto-created)
│   ├── SPVersionCleaner_*.log
│   └── SharePoint_Version_Cleaner_Report_*.html
│
├── Documentation/
│   ├── README.md                     ← This file
│   ├── ARCHITECTURE.md               ← Visual module structure & data flows
│   ├── MODULAR_SUMMARY.md            ← Quick reference guide
│   └── REFACTORING_GUIDE.md          ← Complete implementation status
│
└── .github/
    └── copilot-instructions.md       ← AI agent reference
```

## 🚀 Getting Started

### Prerequisites

- **PowerShell 7+** (verify with `$PSVersionTable.PSVersion.Major`)
- **PnP.PowerShell module** (installed automatically on first run)
- **SharePoint Online admin permissions**
- **Entra (Azure AD) app registration** (see "Entra App Setup" below)

### Quick Start

1. **Clone or download** this repository
2. **Edit `config.json`** with your tenant details:
   ```json
   {
     "settings": {
       "clientId": "your-entra-app-id",
       "adminUrl": "https://your-tenant-admin.sharepoint.com/",
       "keepVersions": 25
     }
   }
   ```
3. **Run as Administrator**:
   ```powershell
   cd SharePoint-Version-Cleaner
   .\SPOVC-Modular.ps1
   ```
4. **Select operation mode** from the GUI menu:
   - **Calculate/Estimate** (scan-only preview)
   - **Dry-Run** (simulate changes without deletion)
   - **Cleanup** (apply version removal)

### Testing Workflow

⚠️ **Always follow this sequence**:

1. Use **Calculate/Estimate** first to preview what will be deleted
2. Review the HTML report in `Logs/`
3. Run **Dry-Run** (if available) to see actual impact
4. Only then proceed with **Cleanup** mode
5. Check logs for detailed operation traces

## 📚 Documentation

| Document | Purpose |
|----------|---------|
| [ARCHITECTURE.md](./ARCHITECTURE.md) | Visual module structure, dependencies, data flows |
| [MODULAR_SUMMARY.md](./MODULAR_SUMMARY.md) | Executive summary, quick reference |
| [REFACTORING_GUIDE.md](./REFACTORING_GUIDE.md) | Implementation status, migration checklist |
| [.github/copilot-instructions.md](./.github/copilot-instructions.md) | Technical reference for developers |

## 🔐 Entra App Setup

To use SPOVC with SharePoint Online, you need to register an application in Microsoft Entra (Azure AD) and configure it with the required permissions. This clientId goes into `config.json`.

### Prerequisites
- PowerShell 7
- Registered app for PnP-Online (Entra)

### Register a PnP App
1. Go to the [Microsoft Entra Admin Center](https://entra.microsoft.com/#view/Microsoft_AAD_RegisteredApps/ApplicationsListBlade/quickStartType~/null/sourceType/Microsoft_AAD_IAM?Microsoft_AAD_IAM_legacyAADRedirect=true).
2. Select **New Registration** in the top left corner.
3. Enter a name for your application and click **Register** at the bottom left corner.

### Assign Required Permissions
1. After registration, you will see the application overview screen.
2. **Copy your Application (client) ID** and store it for later use.
3. Click on **API Permissions** in the left menu.
4. Remove the default `Graph.User` permission or any other permissions added by default.
5. Click **Add a permission** > **SharePoint** > **Delegated permissions**.
6. Add the following permissions:
   - `AllSites.FullControl`
   - `AllSites.Manage`
   - `AllSites.Read`
   - `AllSites.Write`
7. Click **Add permissions** to save.

### Configure Authentication
1. Click on the **Authentication** tab in the left menu.
2. Click **Add a platform** > **Mobile and desktop applications**.
3. In the custom redirect URI field, enter:
   ```
   http://localhost
   ```
4. Save the changes by clicking **Configure** at the bottom right of the screen.

That’s it! Your Entra app registration is complete and ready to use with the SharePoint Version Cleaner scripts.

## ⚙️ Configuration

All settings are stored in `config.json` - no need to edit the PowerShell script!

### config.json Reference

```json
{
  "settings": {
    "clientId": "your-entraid-client-app-id",
    "adminUrl": "https://yourtenantname-admin.sharepoint.com/",
    "keepVersions": 25
  },
  "logging": {
    "directory": "Logs",
    "enableFileLog": true,
    "enableConsoleLog": true
  },
  "inactiveSites": {
    "defaultDaysInactive": 90,
    "defaultMaxStorageMB": 100,
    "includeEmptySites": false
  }
}
```

### Configuration Options

| Option | Default | Description |
|--------|---------|-------------|
| `clientId` | (required) | Entra app ID from your app registration |
| `adminUrl` | (required) | SharePoint admin URL for your tenant |
| `keepVersions` | 25 | Number of recent versions to retain per file |
| `enableFileLog` | true | Write detailed logs to `Logs/` directory |
| `enableConsoleLog` | true | Display logs in PowerShell console |
| `defaultDaysInactive` | 90 | Consider sites inactive after N days |
| `defaultMaxStorageMB` | 100 | Storage threshold for cleanup recommendations |

## ✅ Best Practices & Notes

- **Always test in a non-production environment first**
- **Start with Calculate/Estimate mode** to preview changes
- **Review HTML reports** before running cleanup
- **Check log files** (`Logs/`) for detailed operation traces
- **Ensure you have appropriate SharePoint admin permissions**
- **Backup site data** before large-scale cleanup operations
- **Monitor storage savings** after running cleanup

## 🏗️ Modular Architecture Benefits

SPOVC is built with clean modular design:

- **Testable**: Each module can be tested independently
- **Maintainable**: Changes isolated to specific modules
- **Reusable**: Import modules into other PowerShell scripts
- **Configurable**: External JSON config eliminates script editing
- **Transparent**: Clear separation of concerns (Auth, UI, Processing, Reporting)

For developers, see:
- [ARCHITECTURE.md](./ARCHITECTURE.md) - Module dependencies and data flow
- [REFACTORING_GUIDE.md](./REFACTORING_GUIDE.md) - Implementation details
- [.github/copilot-instructions.md](./.github/copilot-instructions.md) - Technical reference

