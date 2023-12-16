# Notice
# Script requires re-validation by developer to ensure recent changes have not created issues for a normally working script. It is not recommended for use at this time.
## Re-validation includes the re-adding of users to original locations. Feature scope TBD.

# PUID Mismatches
### Beta

![Header](https://github.com/DravenWB/Microsoft_PowerShell_Scripts/assets/46582061/9df92d98-9a60-46ef-a686-08c039ca9164)

## Description
This script was created to handle situations in which a user ID is not matching, particularly for scenarios in which a user does not have access to a shared resource after being explicitly shared.

# Current Functions
- Remove a user from all sites in all locations to include OneDrives which will then allow for a fresh pull of the user ID when something is Shared.
- Print a report with times, dates, locations, errors (if any) and whether they were cleared from the location or not. 

## Disclaimer / Notice
- This script is what I would consider Beta. It has been proven and tested as working but may require minor polishing.
- Currently this script does not check for actual mismatches as that is more of an internal function to Microsoft. However, it may be possible to gather these ID's after further extensive testing for comparison operations.
- In order to check of the user is part of a site, you need to be the site collection administrator for that site. This is a requirement of the Get-SPOUser command.
- This script must be run in PowerShell 5. Currently, the SharePoint Online Management Shell is not supported in PowerShell 7.
- This script will help fix surface level ID mismatches where the sharing source is a SharePoint Site or OneDrive. Mismatches in which the user ID itself is incorrect are not currently supported.

## Planned Features
- Addition of a variety of functions to allow specification of sites, only target OneDrives, etc.
- If it is determined to be possible to get user internal ID's from azure and from SharePoint, I intend on adding actual mismatch checks.
- Implementation of internal level PUID mismatch handling in which the issue is the user ID and not the Site/OneDrive stored ID. (If Possible)

## Script Documentation
- [Get Started with the SharePoint Online Management Shell | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
- [Set-SPOUser](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/set-spouser?view=sharepoint-ps)
- [Remove-SPOUser](https://learn.microsoft.com/en-us/powershell/module/sharepoint-online/remove-spouser?view=sharepoint-ps)
- [about ForEach | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_foreach?view=powershell-7.4)
- [about Switch | Microsoft Learn](https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_switch?view=powershell-7.4)

## Related Issue Documentation
- [Fix site user ID mismatch in SharePoint or OneDrive | Microsoft Learn](https://learn.microsoft.com/en-us/sharepoint/troubleshoot/sharing-and-permissions/fix-site-user-id-mismatch)
- [Troubleshoot user profile removal issues in SharePoint](https://learn.microsoft.com/en-us/sharepoint/remove-users#site-by-site-in-sharepoint)
