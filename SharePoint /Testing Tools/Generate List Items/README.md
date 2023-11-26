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

## PnP PowerShell Legal Disclaimer
Microsoft 365 Patterns and Practices (PnP)

The MIT License (MIT)

Copyright (c) Microsoft Corporation

All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

### Footnotes
<sup>1. A SharePoint health score of 10 indicates severe usage and throttling. An increase from 0 > 3 is heavily impacting for a script but still leaves the tenant in a health state. </sup>
