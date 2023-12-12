# Document Library Inheritance Reset

![shutterstock_1476215399-4226488549](https://github.com/DravenWB/Microsoft_PowerShell_Scripts/assets/46582061/f3c93143-35ad-45ab-931f-3043d83bffdb)

<sup> Unfortunately, not this kind... </sup>

## Description

This script is designed to reset the inheritance of items within a specific folder of a SharePoint Online document library. It makes heavy use of PnP.PowerShell in order to accomplish this but does so successfully in both testing and live environments.

### Functions
- Reset's document inheritance to parent. (Most commonly the document library.)
- File output for review prior to operational execution.
- File output report for post-execution records.
- Light error handling.

## Dependencies
- Due to PnP's recent updates, the module requires PowerShell 7 or greater.
   - The script includes a check to both review current installation and install the module if missing.

## Limitations
- At this time, the script gets all files and folders within a specific directory. To specify it further, you need to scope down the initial pull command to exclude folder types. This functionality may be added in the future.
- Many of the user input properties intitially can be shortened for ease of use though this will take some additional time with string management that I simply don't have at this very moment.
- Depending on the amount of items you will be pulling in, you may need a large amount of system RAM. There is memory clearing checks but that can only do so much.
   - Current usage observed at about 1MB/200 Items.

## Documentation
- [PnP.PowerShell](https://github.com/pnp/powershell)
- [Connect-PnPOnline](https://pnp.github.io/powershell/cmdlets/Connect-PnPOnline.html)
- [Get-PnPContext](https://pnp.github.io/powershell/cmdlets/Get-PnPContext.html)
- [Write-Progress](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/write-progress?view=powershell-7.4)
- [Convert-String | Microsoft](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/convert-string?view=powershell-5.1)
- [Installing PowerShell on Windows - PowerShell 7](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows?view=powershell-7.4#installing-the-msi-package)
- [About Classes - PowerShell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_classes?view=powershell-7.4)
- [Everything You Wanted to Know About Exceptions - PowerShell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-exceptions?view=powershell-7.4)
