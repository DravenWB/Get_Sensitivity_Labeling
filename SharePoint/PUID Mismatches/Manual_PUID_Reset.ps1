####################################################################################################################################################################################
# Description: This script is built to reset PUID mismatches for a user that is experiencing issues with sharing files, accessing sites they've been given permissions to, etc.
#
# How it Works: Effectively, this script will remove the user from the User Information List in a SharePoint site or OneDrive entirely by using the Remove-SPOUser command. This
#               will remove 100% of current permissions for that user from the location the script is run against. Data such as files, sites, etc. will remain in place but the
#               user will no longer have access to them until they are added back to the locations, re-shared previously shared documents from that location, etc.
#
# Dependencies: + SharePoint Online PowerShell Module (Version check and installation command included in script.)
#
# Before You Run: Important information to take into consideration before running this script! Please be sure to read these details in their entirety.
#
#               + The operator will need to be a site collection administrator for all locations it is run on. This is required by Get-SPOUser as a check is placed in the script
#                 to see if that user is part of the site/OneDrive before attempting to run the command to remove them.
#
#                 This script is currently configured to add you to any sites/OneDrives you do not currently have site collection permissions for!
#
#                 While this has no currently known major drawbacks, it may look suspicious if logged. Prior to the script finishing, any sites that you were not previously a site 
#                 collection admin of before running will be removed from your account as part of the cleanup process. This information is logged and stored in a text file saved 
#                 by this script for your records locally.
#
#               + For your ease of review, each section has been blocked out using the pound sign (#). Different operations take place within each section.
#
# Terms found within this script:
#
#               + UIL = User Information List (See README.md for more details)
#               + SPO = SharePoint Online
#               + UPN = User Principal Name
#               + PUID = Persistent Unique Identifier (See README.md for more details)
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
#Set error view and action for clean entry into the output file. Additionally gets the operator's current setting to change it back at script cleanup time. 

#Get operator current error output length and set to concise.
$OriginalErrorView = $ErrorView
$ErrorView = [System.Management.Automation.ActionPreference]::ConciseView

#Get operator current error action and set to stop.
$OriginalErrorAction = $ErrorActionPreference
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue

####################################################################################################################################################################################

#Run check against PowerShell version and verify that it is version 5 or greater. If not, inform the user and exit the script.

Write-Host "Now checking running PowerShell version..."
Start-Sleep -Seconds 1

$InstalledVersion = ($PSVersionTable.PSVersion).ToString()

if ($InstalledVersion -ge '5')
    {
        Write-Host -ForegroundColor Green "Success! PowerShell version $InstalledVersion running."
        Start-Sleep -Seconds 1
    }
        else
            {
                Write-Host -ForegroundColor Red "The currently running PowerShell version is $InstalledVersion."
                Write-Host -ForegroundColor Red "This PowerShell script requires PowerShell version 7 or greater."
                Write-Host -ForegroundColor Red "Please run in PowerShell 7 and try again."
                Start-Sleep -Seconds 3
                Exit
            }


####################################################################################################################################################################################

Write-Host " "

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

####################################################################################################################################################################################

Write-Host " "

Write-Host "Now gathering variables required to run the script..."
Start-Sleep -Seconds 1

Write-Host " "

#Get the admin center URL.
Write-Host "Please enter the URL for your SharePoint Admin Center for connecting."
Write-Host "Ex: https://contoso-admin.sharepoint.com"
$SharePointAdminURL = Read-Host "URL"

Write-Host " "

#Get the admin UPN.
Write-Host "Please enter your SharePoint Administrator email for connection and temporary permissions assignment."
Write-Host "Ex: first.last@tenant.com<mailto:first.last@tenant.com>"
$SharePointAdminUPN = Read-Host "Email"

Write-Host " "

#Get the UPN to run the UIL PUID mismatch for.
Write-Host "Please enter the email of the user to run a PUID mismatch for."
$UserUPN = Read-Host "User UPN"

####################################################################################################################################################################################

Write-Host " "

#Connect to online services required.
Write-Host "Now attempting to connect to SharePoint Online..."
Start-Sleep -Seconds 1

#Attempt to connect to the SharePoint Online service and exit if connection fails as it is required for the script.
try
    {
        Connect-SPOService -Url $SharePointAdminURL
    }

    catch
        {
            try
                {
                    Connect-SPOService -Url $SharePointAdminURL -Region ITAR
                }

                catch
                    {
                        Write-Host -ForegroundColor Red "Failed to connect to SharePoint Online due to error:" $_
                        Exit
                    }
        }

####################################################################################################################################################################################

Write-Host " "

Write-Host "Now gathering available tenant sites for processing..."
Start-Sleep -Seconds 1

#Gathers all sites in the tenant to include OneDrive accounts. Completed early to give operator site count in the next section's disclaimer readout.
try
    {
        $SiteDirectory = Get-SPOSite -Limit All -IncludePersonalSite $true

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

####################################################################################################################################################################################

Write-Host " "

#Inform operator how the operation works and provide important considerations.
Write-Host -ForegroundColor Red " << IMPORTANT >>"
Write-Host -ForegroundColor Yellow "+ PUID mismatch is completed by removing the user from the User Information List."
Write-Host -ForegroundColor Yellow "+ As such, all permissions for the particular user on every site will be removed."
Write-Host -ForegroundColor Yellow "+ This script does not make ID mismatch checks but runs manually for all sites."
Write-Host " "
Write-Host -ForegroundColor Yellow "+ The command Get-SPOUser requires that you are a sharepoint site administrator of every site you want to make changes for"
Write-Host -ForegroundColor Yellow "  to check for user presence."
Write-Host -ForegroundColor Yellow "+ This script checks if you are an active SharePoint site admin for the sites being processed before assigning permissions."
Write-Host -ForegroundColor Yellow "+ A check is also in place to restore original site collection administrator assignments for all sites processed at the end."
Write-Host " "
Write-Host "This operation will be run for the user " $UserUPN " on " $SiteDirectory.Count "sites and OneDrive locations combined."
Write-Host " "
Write-Host "To confirm that you would like to proceed, please enter the word > Confirm <."
Write-Host " "

#Get Confirmation that the operator is ready to proceed with the operation after being provided with details on currently configured functions.
do
{
    $DisclaimerTwo = Read-Host "Proceed?"

    if ($DisclaimerTwo -cne "Confirm")
        {
            Write-Host -ForegroundColor DarkYellow "Input did not match the word Confirm."
            Write-Host -ForegroundColor DarkYellow "Please try again or press Ctrl + C to exit the script."
        }
}
    until ($DisclaimerTwo -ceq "Confirm")

Write-Host " "

####################################################################################################################################################################################
#The following operations are combined into single block for easier data handling. Separated into sub-blocks for readability.

#Data storage class definition.
class OperationData
{
    [int]    $Index
    [string] $Date
    [string] $AdminCheckTime
    [string] $Location
    [bool]   $OriginallyAdmin
    [bool]   $AdminReverted
    [string] $UserCheckTime
    [string] $UserUPN
    [bool]   $UserRemoved
    [string] $AdminErrors
    [string] $UserErrors   
}

$LoggingIndex = @() #Index to store data for operational logging and file output.

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
$IndexCounter = 1 #Initialize counter for index numbering.

Write-Host "Now processing all site collection changes for $UserUPN."
Write-Host "This may take an extended period of time dependent on the size of your Tenant SharePoint..."
Start-Sleep -Seconds 3

#For each site in the tenant...
foreach ($Site in $SiteDirectory)
    {
        #Display progress bar for PUID mismatch resets.
        $ProgressPercent = ($IndexCounter / $SiteDirectory.Count) * 10
        Write-Progress -Activity "Manual PUID Mismatch Resets" -Status "$ProgressPercent% Complete:" -PercentComplete $ProgressPercent

        #Initialize errors to blank space to ensure proper listing + clearing of old data if available.
        $A_Error = " "
        $U_Error = " "

        #Check if the operator is currently a site collection admin for the current site being processed.
        $SiteAdminCheck = (Get-SPOUser -Site $Site.Url -LoginName $SharePointAdminUPN -ErrorAction SilentlyContinue).IsSiteAdmin

        if ($SiteAdminCheck)
            {
                $AdminTime = Get-Date -Format "%R %Z"
                $AdminCurrent = $true
            }

            #If operator is not a site collection administrator for the site being processed, add them in order to run changes for the user.
            else
                {
                    try
                        {
                            Set-SPOUser -Site $Site.Url -LoginName $SharePointAdminUPN -IsSiteCollectionAdmin $true
                        }

                        catch
                            {
                                $A_Error = $_ 
                            }

                    $AdminTime = Get-Date -Format "%R %Z"
                    $AdminCurrent = $false
                }
        
        #Check user presence on the site being processed. If they exist, remove them. If not, move to the next item.
        if (Get-SPOUser -Site $Site.Url -LoginName $UserUPN)
            {
                $UserTime = Get-Date -Format "%R %Z"

                try
                    {
                        Remove-SPOUser -Site $Site.Url -LoginName $UserUPN
                    }

                    catch
                        {
                            $U_Error = $_
                        }

                $UserWasRemoved = $true
            }

            else
                {
                    $UserTime = Get-Date -Format "%R %Z"
                    $UserWasRemoved = $false
                }

        #Write data to instantiated class object for temporary storage and file output.
        $DataTable = New-Object -TypeName OperationData -Property $([Ordered]@{
    
        Index = $IndexCounter
        Date = Get-Date -Format "%m/%d/%Y"
        AdminCheckTime = $AdminTime
        Location = $Site.Url
        OriginallyAdmin = $AdminCurrent
        UserCheckTime = $UserTime
        UserUPN = $UserUPN
        UserRemoved = $UserWasRemoved
        AdminErrors = $A_Error
        UserErrors = $U_Error
        })

        #Send the data table to the index.
        $LoggingIndex += $DataTable

        #Clear data table and increment counter for next site.
        $DataTable = $null
        $IndexCounter++
    }

####################################################################################################################################################################################

#Save script data to file.
$SaveModifier = Get-Date -Format "%m/%d/%Y"

do
    {
        Write-Host "This portion of the script has been placed into a loop in case saving the file fails."
        Write-Host "Save defaults to the file name PUID_Mismatch_Log.csv on the desktop. If it fails, it will try to append the date as a backup."
        Write-Host "Once the script completes, changes made that have been stored in memory will be lost."
        Write-Host "Please ensure you have the data you need before continuing."
        Write-Host " "
        Write-Host "1. Attempt to save file with default settings."
        Write-Host "2. Manually input location and filename for saving."
        Write-Host "3. Complete cleanup and Exit"

        $SaveSelection = Read-Host "Selection"

        switch($SaveSelection)
            {
                '1'
                {
                    try
                        {
                            $LoggingIndex | Export-Csv -Path ~\Desktop\PUID_Mismatch_Log.csv -NoClobber
                            Write-Host -ForegroundColor Green "The file has successfully been saved to the following location:"
                            Write-Host -ForegroundColor Green "> ~\Desktop\PUID_Mismatch_Log.csv <"
                        }

                        catch
                            {
                                Write-Host -ForegroundColor Yellow "Saving the file failed. Now attempting to add modifier and try again."
                                Start-Sleep -Seconds 3

                                $LoggingIndex | Export-Csv -Path ~\Desktop\PUID_Mismatch_Log$SaveModifier.csv -NoClobber 
                            }
                }
                
                '2'
                {
                    Write-Host "Please enter a location to save your file."
                    Write-Host "EX: ~\Documents\"
                    Write-Host "EX: C:\Users\Username\Documents\Folder\"

                    $SaveLocation = Read-Host "Save Location"

                    Write-Host " "

                    Write-Host "Please enter a name for your file."
                    Write-Host "EX: Site Changes"

                    $SaveName = Read-Host "Save File Name"

                    try
                        {
                            $LoggingIndex | Export-Csv -Path $SaveLocation+$SaveName.csv -NoClobber
                            Write-Host -ForegroundColor Green "The file has successfully been saved to the following location:"
                            Write-Host -ForegroundColor Green "> $SaveLocation+$SaveName.csv <"
                        }

                        catch
                            {
                                Write-Host -ForegroundColor Yellow "Saving the file failed. Now attempting to add modifier and try again."
                                Start-Sleep -Seconds 3

                                $LoggingIndex | Export-Csv -Path $SaveLocation+$SaveName+$SaveModifier.csv -NoClobber
                            }
                }

                '3'
                {
                    Write-Host -ForegroundColor Green "Now proceeding with admin credential correction and cleanup."
                    Write-Host -ForegroundColor Yellow "This may take some time depending on how many changes were required."
                    Start-Sleep -Seconds 3
                }
            }
    }

    until($SaveSelection -eq '3')

####################################################################################################################################################################################

#Script cleanup.

#Resore original admin site permissions.
foreach ($Entry in $LoggingIndex)
    {
        if (-not($Entry.OriginallyAdmin))
            {
                try
                    {
                        Set-SPOUser -Site $Entry.Location -LoginName $SharePointAdminUPN -IsSiteCollectionAdmin $false
                        $Entry.AdminReverted = $true
                    }

                    catch
                        {
                            $Entry.AdminReverted = $false
                            $Entry.AdminErrors = $Entry.AdminErrors + "Removal Error:" + $_
                        }
            }

            else
                {
                    $Error.AdminReverted = $false
                }
    }

#Release majority memory usage.
$LoggingIndex = $null

#Reset error view/length to operator original setting.
$ErrorView = [System.Management.Automation.ActionPreference]::$OriginalErrorView
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::$OriginalErrorAction

####################################################################################################################################################################################

Write-Host -ForegroundColor Green "Script now complete! Have a wonderful day! :)"
Start-Sleep -Seconds 3
Exit
