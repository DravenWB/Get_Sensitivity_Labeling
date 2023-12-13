####################################################################################################################################################################################
# Description: This script is designed to remove an administrator from all sites that they do not own. This is mostly helpful in scenarios in which an individual has become
#              the SharePoint site admin for a large list of sites.
#
# Development Date: December 13, 2023
#
####################################################################################################################################################################################

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

####################################################################################################################################################################################
#Get operator details.

Write-Host "Now gathering variables required to run the script..."
Start-Sleep -Seconds 1

Write-Host " "

#Get the admin center URL.
Write-Host "Please enter the URL for your SharePoint Admin Center for connecting."
Write-Host "Ex: https://contoso-admin.sharepoint.com"
Write-Host " "
$SharePointAdminURL = Read-Host "URL"

Write-Host " "

#Get the admin UPN.
Write-Host "Please enter your SharePoint Administrator email for connection and temporary permissions assignment."
Write-Host "Commercial Example: UPN@tenant.com"
Write-Host " "
$SharePointAdminUPN = Read-Host "Email"

Write-Host " "

####################################################################################################################################################################################

class OperationData
{
    [int]    $Index
    [string] $Date
    [string] $Time
    [string] $Location
    [bool]   $SiteCollectionAdmin
    [bool]   $IsSiteOwner
    [string] $RemovedAsAdmin
    [string] $XErrors  
}

####################################################################################################################################################################################

Write-Host "Now reviewing all SharePoint sites for administrator and ownership settings..."
Start-Sleep -Seconds 2

#Get all sites and assign to variable.
$SiteDirectory = Get-SPOSite -Limit All -IncludePersonalSite $true

#Initialize the index counter.
$IndexCounter = 0

#Filter through every site and check if the operator is an administrator / owner.
foreach ($Site in $SiteDirectory)
    {
        #Display progress bar for admin/owner checks.
        $ProgressPercent = ($IndexCounter / $SiteDirectory.Count) * 100
        $ProgressPercet = $ProgressPercent.ToString("#.##")
        Write-Progress -Activity "Setting Site Admin Permissions..." -Status "$ProgressPercent% Complete:" -PercentComplete $ProgressPercent

        #Check if user is part of the site.
        try
            {
                Get-SPOUser -Site $Site.Url -LoginName $SharePointAdminUPN
                $Run = $true
            }

            catch {$Run = $false}

        #Check if operator is part of the site.
        if ($Run)
            {
                #Get user data for site.
                $User = Get-SPOUser -Site $Site.Url -LoginName $SharePointAdminUPN

                #Get site owner group name(s).
                if ($User.Groups -notcontains (Get-SPOSiteGroup -Site $Site.Url | Where {$_.Roles -contains "Full Control"}).Title)
                    {
                        Set-SPOUser -Site $Site.Url -LoginName $SharePointAdminUPN -IsSiteCollectionAdmin $false
                    }
            }
    }
