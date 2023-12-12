# Get All SharePoint Site Owners
![Site-Owner-1024x576-1665104080](https://github.com/DravenWB/Microsoft_PowerShell_Scripts/assets/46582061/44668fdc-b791-4945-856d-8f0b97d5721a)

## Description
This script is intended to gather ALL site owners for every SharePoint site in a Microsoft Online tenant. This was created as the current standard Get-SPOSite command only retrieves a single owner per site when used and is sometimes not detailed enough.

## Limitations
- Currently, this script only gets owners for default SharePoint groups. This will be added to the script in the future for better functionality.
- The script also gathers Online groups ending in _o as ID's and does not define them. Currently investigating method to display these more clearly.

## Documentation
- [Getting Started with the SharePoint Online Management Shell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
- [Get-SPOSite | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/get-sposite?view=sharepoint-ps)
- [Export-CSV | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.utility/export-csv?view=powershell-7.4)

## Future Planned Updates
- Expand group gathering to include non-default owner groups.
- Transcribe group ID's if possible to readable group names.
- Add operator check for acceptance of risk.
