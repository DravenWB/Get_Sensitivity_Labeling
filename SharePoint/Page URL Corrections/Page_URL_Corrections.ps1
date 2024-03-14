$OperatorAcknowledgement = " "

Write-Host ""
Write-Host -ForegroundColor Yellow "Disclaimer: This script is not officially supported by Microsoft, its affiliates or partners."
Write-Host -ForegroundColor Yellow "This script is provided as is and the responsibility of understanding the script's functions and operations falls upon those that may choose to run it."
Write-Host -ForegroundColor Yellow "Positive or negative outcomes of this script may not receive future assistance as such."
Write-Host -ForegroundColor Yellow ""
Write-Host -ForegroundColor Yellow "To acknowledge the above terms and proceed with running the script, please enter the word > Accept < (Case Sensitive)."
Write-Host ""

$OperatorAcknowledgement = Read-Host "Acknowledgement"

if ($OperatorAcknowledgement -cne "Accept") #If operator acknowledgement check is not matched to "Accept", exit the script.
{
    Exit
}

#Checks to ensure that PowerShell 7 or greater is installed. If not, attempt installation
Write-Host "Now checking running PowerShell version..."
Start-Sleep -Seconds 1

#Get powershell version and set to string for check and/or output.
$InstalledVersion = ($PSVersionTable.PSVersion).ToString()

#If PowerShell version is greater than or equal to 7...
if ($InstalledVersion -ge '7')
    {
        #Inform the operator that the correct version required is installed.
        Write-Host ""
        Write-Host -ForegroundColor Green "Success! PowerShell version $InstalledVersion running."
        Start-Sleep -Seconds 1
    }

    else #Inform the operator that the correct version required is not installed and need to be run in PowerShell 7. Exit script upon completion.
        {
            Write-Host ""
            Write-Host -ForegroundColor Red "The currently running PowerShell version is $InstalledVersion."
            Write-Host -ForegroundColor Red "This PowerShell script requires PowerShell version 7 or greater."
            Write-Host -ForegroundColor Red "Please run in PowerShell 7 and try again."
            Start-Sleep -Seconds 3
            Exit
        }

#Checks to see if PnP.PowerShell is installed. If not, attempt to install it. If fails, exit script.
Write-Host -ForegroundColor Green "Now checking installed PnP.PowerShell version..."
Write-Host ""
Start-Sleep -Seconds 1
   
    #If module is installed...
    if (Get-Module -ListAvailable -Name "PnP.PowerShell")
        {
            #Inform the operator and continue.
            Write-Host -ForegroundColor Green "The PnP PowerShell Module is confirmed as installed!"
            Start-Sleep -Seconds 1
        }

            else #If module not found...
            {
                try #Inform the user and try to install the module.
                {
                    Write-Host -ForegroundColor Yellow "PnP PowerShell Module not found. Now attempting to install the module..."
                    Install-Module -Name PnP.PowerShell -Scope CurrentUser
                    Start-Sleep -Seconds 1
                    Import-Module -Name PnP.PowerShell -Scope Local

                    Write-Host -ForegroundColor Green "Success! PnP.PowerShell now installed and loaded!"
                }

                    catch #If installation fails, inform the user and exit.
                    {
                        Write-Host -ForegroundColor Red "Failed to install the PnP PowerShell Module due to error:" $_
                        Exit
                    }
            }

#Prompt the user to input required variables.
Write-Host ""
Write-Host "Please enter the URL of your SharePoint admin center:"
Write-Host "Example: https://contoso-admin.sharepoint.com"
Write-Host "Example: https://contoso-admin.sharepoint.us"
Write-Host "NOTE: If the name contains spaces, ensure you place a quotation mark at the beginning and end to ensure accurate input."

$AdminCenterURL = Read-Host "Admin Center URL"

#Removes quotations if present.
if ($AdminCenterURL.StartsWith('"') -or $AdminCenterURL.EndsWith('"'))
    {$AdminCenterURL = ($AdminCenterURL.Trim('"'))}

#Ensure the URL is ended with a /
if (-not $AdminCenterURL.EndsWith('/'))
    {$AdminCenterURL += '/'}

####################################################################################################################################################################################
#Connect to PnP Online and exit if it fails. If succeeds, get the context and proceed.
Write-Host ""
try {Connect-PnPOnline -Url $AdminCenterURL -UseWebLogin}

    catch
        {
            Write-Host -ForegroundColor Red "There was an error connecting to PnP Online: $_"
            Exit
        }

#Get the context.
$Context = Get-PnPContext

####################################################################################################################################################################################

#Connect to the SPO service and try geo-location locks if required.
try
{
    Connect-SPOService -Url $SPOAdminCenterURL

    Write-Host -ForegroundColor Green "Successfully connected to the SPO Service!" #Operator process information.
}

    catch
    {
        if($Endpoint -eq ".us")
        {
            try
            {
                Connect-SPOService -Url $SPOAdminCenterURL -Region ITAR

                Write-Host -ForegroundColor Green "Successfully connected to the SPO Service under ITAR restrictions!" #Operator process information.
            }

                catch
                {
                   Write-Host -ForegroundColor Red "Failed to connect to the GCC-High SPO Service with the following error:" $_
                   break;
                }
        }

        else
        {
            Write-Host -ForegroundColor Red "Failed to connect to the commercial SPO Service with the following error:" $_
            break;
        }
    }

####################################################################################################################################################################################

#Get the SharePoint sites for the entire tenant to conduct URL mismatch checks.
$SiteDirectory = Get-SPOSite -Limit All

#Process the site pages of every site to ensure the properties of that page matches the site it is located on.
foreach ($Site in $SiteDirectory)
{
    $Context.Load($Site) #Call the context to load the site.
    $Context.ExecuteQuery

    #Get dynamicly generated, relative URL for comparison with SharePoint Page property.
    $SiteRefURL = ("/sites/" + ((Get-SPOSite -Identity "https://spocotest.sharepoint.com/sites/Testsite65").Url -split("/"))[-1] + "/" + "SitePages")

    #Get all pages that do not start with the relative site URL.
    $CurrentSitePages = Get-PnPListItem -List "Site Pages" | Where {-Not $_.FieldValues.FileRef.StartsWith($SiteRefURL)}

    #For all pages that don't match and have broken navigation...
    foreach ($Page in $CurrentSitePages)
    {
        #Load the page context.
        $Context.Load($Page)
        $Context.ExecuteQuery

        #Set the page URL properties.
        [string]$FileLeafRef_New = $SiteRefURL + "/" + $Page.FileLeafRef

        #Set field values and trigger save under new version to allow for reversion if required.
        try
            {
                Set-PnPListItem -List $Page.FileLeafRef -Values @{"FileRef" = $FileLeafRef_New; "FileDirRef" = $SiteRefURL} -UpdateType Update
            }

            catch
                {
                    #Inform the user of page update failure.
                    Write-Host -ForegroundColor Red "An error has occurred in configuring the new page properties: $_"
                }
    }
}
