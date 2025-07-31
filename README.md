# SharePoint Version Cleaner

This repository contains PowerShell scripts and guides for cleaning up file versions in SharePoint Online document libraries. The tools help administrators manage and reduce the number of file versions stored, saving space and improving performance.

## Project Structure

- **Backup Versions/**
  - `SPOVC.ps1`: Legacy or backup version of the main version cleaner script.
  - `SPOVC-GraphPicker.ps1`: A version of the script that uses Microsoft Graph and provides a graphical picker for site selection.

- **SharePoint Version Cleaner/**
  - `SPOVC.ps1`: The main PowerShell script for cleaning up SharePoint file versions. Includes a Windows Forms GUI for user interaction.
  - `SharePoint File Version Cleaner Guide.docx`: User guide with detailed instructions and screenshots.

- **SharePoint Version Cleaner - CSV Support/**
  - `SPOVC.ps1`: Enhanced version of the cleaner script that supports batch processing of multiple SharePoint sites via a CSV file.
  - `SharePoint File Version Cleaner Guide.docx`: User guide for the CSV-enabled version.
  - `exampleCSV.csv`: Example CSV file format for batch processing. The file should contain a column `SiteUrl` with one SharePoint site URL per line.

## Getting Started

1. **Prerequisites**
   - Windows PowerShell 5.1 or PowerShell 7+
   - Required modules: `PnP.PowerShell`, `Microsoft.Graph` (for GraphPicker version)
   - SharePoint Online admin permissions

2. **Running the Script**
   - Download or clone this repository.
   - Open PowerShell as Administrator.
   - Navigate to the desired script directory.
   - Run the script:
     ```powershell
     .\SPOVC.ps1
     ```
   - For the CSV version, edit `exampleCSV.csv` with your site URLs and run the script in the `SharePoint Version Cleaner - CSV Support` folder.

3. **Guides**

## Configuring an Entra (Azure AD) App for PnP-Online

To use these scripts with SharePoint Online, you need to register an application in Microsoft Entra (Azure AD) and grant it the necessary permissions. This is required for authentication and to allow the scripts to access your SharePoint sites.

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

## Example CSV Format

```
SiteUrl
https://yourtenant.sharepoint.com/sites/site1
https://yourtenant.sharepoint.com/sites/site2
https://yourtenant.sharepoint.com/sites/site3
```

## Notes
- Always test scripts in a non-production environment first.
- Ensure you have appropriate permissions before running cleanup operations.
- Logs are generated in a `Logs` subfolder for each run.

## License & Support

This project is provided **free to use by anyone**.

However, please note that it is **not actively maintained or supported**. These scripts were originally created for a specific internal project and are shared here in the hope that they may be useful to others.

Feel free to use, modify, or adapt them as needed, but no guarantees are provided regarding updates, bug fixes, or support.
 
Use at your own risk. Contributions and suggestions are welcome.

