# Document Library Inheritance Reset | BETA

![shutterstock_1476215399-4226488549](https://github.com/DravenWB/Microsoft_PowerShell_Scripts/assets/46582061/f3c93143-35ad-45ab-931f-3043d83bffdb)

<sup> Unfortunately, not this kind... </sup>

## Description

This script is designed to reset the inheritance of items within a specific folder of a SharePoint Online document library. It makes heavy use of PnP.PowerShell in order to accomplish this but does so successfully in both testing and live environments.

### Functions
- Reset's document inheritance to parent. (Most commonly the document library / Parent Folder)
- File output for review prior to operational execution.
- Document count confirmation capabilities prior to runtime.
- File output report for post-execution records.
   - Amount of records per file configureable. 
- Error handling.

## Dependencies
- Due to PnP's recent updates, the module requires PowerShell 7 or greater.
   - The script includes a check to both review current installation and install the module if missing.
- Dependencies directory required.
   - Contains several required .ps1 files.
   - Must be in same directory relationship to function. (Directory error handling and better structuring planned.)
   - Recommended to download the entire "Document Inheritance Reset" folder to operate.

## Limitations
- This scripts limitations are implied upon the system it is run on. As CAML does not match PnP Queries in item counts, the full amount total needs to be pulled via PnP.
   - This is reduced by only pulling minor data per file for the total count and then handled in batches for the processing but is still fairly resource intensive.
- The script is currently only able to be run for full document libraries and cannot target specific folders. While the script originally had this functionality, it appears to have broken with an update to a dependency but may be re-added in the future.

## Documentation
- [PnP.PowerShell](https://github.com/pnp/powershell)
- [Connect-PnPOnline](https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html)
- [Get-PnPContext](https://pnp.github.io/powershell/cmdlets/Get-PnPContext.html)
- [Write-Progress](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-progress?view=powershell-7.4)
- [Convert-String | Microsoft](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convert-string?view=powershell-5.1)
- [Installing PowerShell on Windows - PowerShell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4#installing-the-msi-package)
- [About Classes - PowerShell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_classes?view=powershell-7.4)
- [Everything You Wanted to Know About Exceptions - PowerShell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.4)
