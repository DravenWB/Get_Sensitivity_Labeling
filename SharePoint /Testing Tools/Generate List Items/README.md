# Generate List Items
<img width="1000" alt="large-1260358170" src="https://github.com/DravenWB/Microsoft_PowerShell_Scripts/assets/46582061/4ece54c7-82a2-46de-b62b-819df71ef3f9"> 

## Description
This script is intended to generate many list items in SharePoint for testing purposes and can do so on a large basis if required. It has currently been tested by myself for the generation of 9,000+ items to test the limitations of SharePoint. Do be cautious when running this script as I have seen a health score increase by 3 when doing so. <sup>1</sup>

## Limitations
- This script does require the use of PnP PowerShell as I have yet to identify a method of doing this via the standard SharePoint Management Shell.

## Documentation
- [PnP/PowerShell](https://github.com/pnp/powershell)
- [Getting Started with the SharePoint Online Management Shell](https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online)
- [SharePoint Limits](https://learn.microsoft.com/en-us/office365/servicedescriptions/sharepoint-online-service-description/sharepoint-online-limits)
- [X-SharePointHealthScore Header | Microsoft Learn](https://learn.microsoft.com/en-us/openspecs/sharepoint_protocols/ms-wsshp/c60ddeb6-4113-4a73-9e97-26b5c3907d33)

### Footnotes
<sup>1. A SharePoint health score of 10 indicates severe usage and throttling. An increase from 0 > 3 is heavily impacting for a script but still leaves the tenant in a health state. </sup>
