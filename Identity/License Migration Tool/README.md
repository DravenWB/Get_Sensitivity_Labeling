# License Migration Tool
![image](https://github.com/DravenWB/Microsoft_PowerShell_Scripts/assets/46582061/44a3b645-9127-4c2a-ba2d-e4281a739130)

## Description
This tool is intended to assist tenants in migrating licenses for an entire tenant if they are not dynamically assigned. This utility allows you to select a license to remove and a license to add which will then replace one with the other for all users who have the license selected to remove. The script will also print out a CSV file containing all the changes made to the envirionment for documentation purposes. It is built into a menu based system in case there are any issues with saving the file.

Ex: You can remove all Office 365 E1 licenses from all users and then give them an Office 365 E3 License in the place of the previous license.

- This script does support and has been proven to work in GCC-High environments.<sup>1</sup>

## Documentation:
- [Install the Microsoft Graph PowerShell SDK | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0)
- [Using Microsoft Graph PowerShell authentication commands | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands?view=graph-powershell-1.0)
- [Get-MgSubscriberSku](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.identity.directorymanagement/get-mgsubscribedsku?view=graph-powershell-1.0)
- [Get-MgUser](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users/get-mguser?view=graph-powershell-1.0)
- [Set-MgUserLicense](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.users.actions/set-mguserlicense?view=graph-powershell-1.0)
- [Export-Csv](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-csv?view=powershell-7.3)

## Future Versions:
- A larger version of this script exists that requires further development and testing. The larger version supports a multitude of license based operations. Release date TBD.

### Footnotes
<sup>1. While this script supports GCC-High environments, the responsibility of running the script still remains with the operator. While you may ask questions, no official support exists for making the script work if it runs into issues within your envirionment.</sup>
