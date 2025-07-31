# SharePoint Version Cleaner - CSV Support

This folder contains an enhanced version of the SharePoint Version Cleaner script that supports batch processing of multiple sites using a CSV file.

## Files

- **SPOVC.ps1**: PowerShell script for cleaning up file versions across multiple SharePoint sites listed in a CSV file.
- **exampleCSV.csv**: Example CSV file. List your site URLs under the `SiteUrl` column.

## Example CSV Format

```
SiteUrl
https://yourtenant.sharepoint.com/sites/site1
https://yourtenant.sharepoint.com/sites/site2
https://yourtenant.sharepoint.com/sites/site3
```

## Usage

1. Edit `exampleCSV.csv` to include your SharePoint site URLs.
2. Open PowerShell as Administrator.
3. Navigate to this folder.
4. Run the script:
   ```powershell
   .\SPOVC.ps1
   ```
5. Follow the prompts and refer to the guide for more information.

## Notes
- Ensure you have the required permissions and modules installed (`PnP.PowerShell`).
- Logs are saved in the `Logs` subfolder. 