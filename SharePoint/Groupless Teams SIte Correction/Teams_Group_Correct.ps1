#Operator risk acknowledgemenet initialization to ensure blank variable in case script is cancelled after running the first time, then run again.
$OperatorAcknowledgement = " "

#Print disclaimer to the screen for the operator.
Write-Host -ForegroundColor DarkYellow "Disclaimer: This script is not officially supported or endorsed by Microsoft, its affiliates or partners"
Write-Host -ForegroundColor DarkYellow "This script is provided as is and the responsibility of understanding the scripts functions and operations falls upon those that may choose to run it."
Write-Host -ForegroundColor DarkYellow "Positive or negative outcomes of this script may not receive future assistance as such."
Write-Host -ForegroundColor DarkYellow ""
Write-Host -ForegroundColor DarkYellow "To acknowledge the above terms and proceed with running the script, please enter > Accept < (Case Sensitive)."

#Get operator confirmation.
$OperatorAcknowledgement = Read-Host "Acknowledgement"

#Check operator confirmation. If confirmation does not equal "Accept", print message to screen and exit the script.
if ($OperatorAcknowledgement -cne "Accept")
{
    Write-Host "Either the acknowledgement input does not match the word Accept or you have not agreed to accept the risk of running this script."
    Start-Sleep -Seconds 1
    Write-Host "The script will now exit. Have a nice day!"
    Exit
}

Write-Host " "
Write-Host -ForegroundColor Green "Acknowledgement accepted!"
Write-Host " "

###############################################################################################################################################################
#Run check against PowerShell version and verify that it is version 5 or greater. If not, inform the user and exit the script.

Write-Host "Now checking running PowerShell version..."
Start-Sleep -Seconds 1

$InstalledVersion = ($PSVersionTable.PSVersion).ToString()

if ($InstalledVersion -ge '5.2')
    {
        Write-Host -ForegroundColor Green "Success! PowerShell version $InstalledVersion running."
        Start-Sleep -Seconds 1
    }
        else
            {
                Write-Host -ForegroundColor Red "The currently running PowerShell version is $InstalledVersion."
                Write-Host -ForegroundColor Red "This PowerShell script requires PowerShell version 5.2 or greater."
                Write-Host -ForegroundColor Red "Please run in PowerShell 5.2 or greater and try again."
                Start-Sleep -Seconds 3
                Exit
            }

Write-Host " "

###############################################################################################################################################################

#Check if the SPO management module is installed and loaded.
Write-Host -ForegroundColor Green "Now checking for the SharePoint Online Management Shell installation status..."
Start-Sleep -Seconds 1

Write-Host " "

if (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")
    {
        Write-Host -ForegroundColor Green "The SharePoint Online Management shell is confirmed as installed!"
        Import-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope Local
        Start-Sleep -Seconds 1
    }

        else #If module not found, attempt to install the module.
        {
            try
            {
                Write-Host -ForegroundColor DarkYellow "SharePoint Online Management shell not found. Now attempting to install the module..."
                Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
                Start-Sleep -Seconds 3
                Import-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope Local
            }

                catch
                {
                    Write-Host " "
                    Write-Host -ForegroundColor Red "Failed to install the SharePoint Online Module due to error:" $_
                    Exit
                }
        }

Write-Host " "

###############################################################################################################################################################

Write-Host "Now gathering variables required to run the script..."
Start-Sleep -Seconds 1

Write-Host " "

#Get the admin center URL.
Write-Host "Please enter the URL for your SharePoint Admin Center for connecting."
Write-Host "Ex: https://contoso-admin.sharepoint.com"
Write-Host " "
$SharePointAdminURL = Read-Host "URL"

#Get admin UPN for permissions checks.
Write-Host "Please enter your SharePoint administrator email address:"
$AdminUPN = Read-Host "Admin UPN"

Write-Host " "
###############################################################################################################################################################

#Connect to online services required.
Write-Host "Now attempting to connect to SharePoint Online..."
Start-Sleep -Seconds 1

#Attempt to connect to the SharePoint Online service and exit if connection fails as it is required for the script.
try
    {Connect-SPOService -Url $SharePointAdminURL}

    catch
        {
            try
                {Connect-SPOService -Url $SharePointAdminURL -Region ITAR}

                catch
                    {
                        Write-Host -ForegroundColor Red "Failed to connect to SharePoint Online due to error:" $_
                        Exit
                    }
        }

###############################################################################################################################################################

Write-Host " "

Write-Host "Now gathering available tenant sites for processing..."
Start-Sleep -Seconds 1

#Gathers all sites in the tenant without an M365 Group ID.
try
    {
        $SiteDirectory = Get-SPOSite -Limit All | Where-Object {$_.GroupId -eq "00000000-0000-0000-0000-000000000000" -or $_.GroupId -eq ""}

        Write-Host " "
        Write-Host -ForegroundColor Green "Successfully gathered SharePoint site information!"
    }
        catch
            {
                Write-Host " "
                Write-Host -ForegroundColor Red "There was an error gathering the required site data: $_"
                Write-Host -ForegroundColor Red "This script will now exit."
                Exit
            }

###############################################################################################################################################################
#Ensure operator is a site collection admin and create a site M365 Private group for all items.
Write-Host -ForegroundColor Green "Now processing changes..."

Start-Sleep -Seconds 1

#Initialize processing counter.
$ProcessingCounter = 0
$TotalItems = $SiteDirectory.Count()
$ReversionIndex = @() #Variable to contain sites that the operator was not originally a site collection administrator of for later reversion.
$ErrorRecord = @() #Index to record errors for later review.

foreach($Site in $SiteDirectory)
    {
        $PercentComplete = ($ProcessingCounter/$TotalItems) * 100
        Write-Progress -Activity "Configuring private M365 Groups..." -Status "$ProcessingCounter out of $SiteDirectory completed." -PercentComplete $PercentComplete

        try 
            {
                Get-SPOUser -Site $Site.Url -LoginName $AdminUPN | Out-Null
            }
                catch 
                {
                    $ReversionIndex += $Site

                    Set-SPOUser -Site $Site.Url -LoginName $AdminUPN -IsSiteCollectionAdmin $true | Out-Null
                }

        try 
            {
                $Title = $Site.Title
                Set-SPOSiteOffice365Group -Site $Site.Url -DisplayName $Title -Description "Regenerated group for site collection $Title" -Alias $Title | Out-Null

                Write-Host -ForegroundColor Green "$Title successfully associated with M365 Private group."
            }
                catch 
                    {
                        Write-Host -ForegroundColor Yellow "An error occurred while creating the group. Now attempting to modify M365 group title."
                        Start-Sleep -Seconds 1

                        try 
                            {
                                $ModTitle = $Title += "-Mod"
                                Set-SPOSiteOffice365Group -Site $Site.Url -DisplayName $ModTitle -Description "Regenerated group for site collection $Title" -Alias $ModTitle
                            }
                            catch 
                                {
                                    $SiteURLCurrent = $Site.Url
                                    $XError = "Error: Second attempt to associate M365 Private group failed for $Title ($SiteURLCurrent)."
                                    Write-Host -ForegroundColor Red $XError
                                    $ErrorRecord += $XError
                                }
                        
                    }
        }

#Revert operator permissions back to previous state, print error report to local desktop and clear memory.

foreach ($Site in $ReversionIndex)
    {
        Set-SPOUser -Site $Site -LoginName $AdminUPN -IsSiteCollectionAdmin $false
    }

#If the error record is not empty, print to the local desktop.
if($ErrorRecord -ne $null)
    {
        $ErrorRecord | Out-File -FilePath ~/desktop/M365Group_Error_Record.txt
    }

Write-Host -ForegroundColor Green "All operations complete! You may find a record of errors on the local desktop if any were present."
