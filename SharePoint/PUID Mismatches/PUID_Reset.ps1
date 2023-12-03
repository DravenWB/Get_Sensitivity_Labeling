#Powershell script to reset PUID mismatches.

#  WARNING. SCRIPT IS NOT CURRENTLY COMPLETE. DO NOT USE.


#############################################################################################################################################################################################################

#Operator risk acknowledgemenet initialization.
$OperatorAcknowledgement = " "

Write-Host -ForegroundColor DarkYellow "Disclaimer: This script is not officially supported by Microsoft, its affiliates or partners"
Write-Host -ForegroundColor DarkYellow "This script is provided as is and the responsibility of understanding the scripts functions and operations falls upon those that may choose to run it."
Write-Host -ForegroundColor DarkYellow "Positive or negative outcomes of this script may not receive future assistance as such."
Write-Host -ForegroundColor DarkYellow ""
Write-Host -ForegroundColor DarkYellow "To acknowledge the above terms and proceed with running the script, please enter > Accept < (Case Sensitive)."

$OperatorAcknowledgement = Read-Host "Acknowledgement"

if ($OperatorAcknowledgement -ceq "Accept")
{
    #Get the admin center URL.
    Write-Host "Please enter the URL for your SharePoint Admin Center for connecting."
    Write-Host "Ex: https://contoso-admin.sharepoint.com"
    $SharePointAdminURL = Read-Host "URL"

    #Get the admin UPN.
    Write-Host "Please enter your SharePoint Admin email."
    Write-Host "Ex: first.last@tenant.com"
    $SharePointAdminUPN = Read-Host "Email"

    Write-Host "Please enter the UPN of the user to run a PUID mismatch for."
    $UserUPN = Read-Host "User UPN"

    #Check if the exchange online management module is installed and loaded.
    Write-Host -ForegroundColor Green "Now checking for the SharePoint Online Management Shell installation status."
    Start-Sleep -Seconds 1

    if (Get-Module -ListAvailable -Name "Microsoft.Online.SharePoint.PowerShell")
        {
            Write-Host -ForegroundColor Green "The SharePoint Online Management shell is confirmed as installed."
            Start-Sleep -Seconds 1
        }

         else #If module not found, attempt to install the module.
            {
                try
                {
                    Write-Host -ForegroundColor DarkYellow "SharePoint Online Management shell not found. Now attempting to install the module."
                    Install-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope CurrentUser
                    Start-Sleep -Seconds 1
                    Import-Module -Name Microsoft.Online.SharePoint.PowerShell -Scope Local
                }

                    catch
                    {
                        Write-Host -ForegroundColor Red "Failed to install the SharePoint Online Module due to error:" $_
                    }
            }

    #============================================================================================================================================================================#

    #Connect to services required.
    Write-Host -ForegroundColor Green "Now attempting to connect to ExchangeOnline and IPPSSession..."
    Start-Sleep -Seconds 1

    #Attempt to connect to the Exchange Online service and exit if connection fails as it is required for the script.
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

    #============================================================================================================================================================================#

    class SiteMatchIndex
        {
            [string] ${User}
            [string] ${SiteURL}
        }

    $SiteMatchIndex = @() #Variable to store sites that need to be corrected for the user.

    #Get all SharePoint sites, then store the site if an entry for the user exists.

    $SiteDirectory = Get-SPOSite -Limit All -IncludePersonalSite $true

    foreach ($Site in $SiteDirectory)
        {
            Set-SPOUser -Site $Site.Url -LoginName $SharePointAdminUPN -IsSiteCollectionAdmin $true

            $SiteUsers = Get-SPOUser -Limit All -Site $Site.Url 

            foreach ($User in $SiteUsers)
                {
                    if ($User.LoginName -eq $UserUPN)
                        {
                            $Object = New-Object PSObject -Property @{
                            User = $User.LoginName
                            SiteURL = $Site.Url
                            }

                            $SiteMatchIndex += $Object #Send object to data array.

                            $Object = $Null #Clear object for next use.
                        }
                }
        }

    class OldUserData
        {
            [string] ${SiteURL}
            [string] ${LoginName}
            [string] ${SiteAdmin}
            [string] ${OldGroups}
            [string] ${OldType}
        }

    $OldUserData = @()

    foreach ($Match in $SiteMatchIndex)
        {
            $OldData = Get-SPOUser -Site $Match.SiteURL -LoginName $Match.User

            $Object = New-Object PSObject -Property @{
            SiteURL = $Match.SiteURL
            LoginName = $OldData.LoginName
            SiteAdmin = $OldData.IsSiteAdmin
            OldGroups = $OldData.Groups
            OldType = $OldData.UserType
            }

            $OldUserData += $Object #Send object to data array.
            $Object = $Null #Clear object for next use.

            Remove-SPOUser -Site $Match.SiteURL -LoginName $Match.User
        }
    
    foreach ($Site in $OldUserData)
        {
            Add-SPOUser -Site "https://csstrailblazer.sharepoint.us/sites/WolfePack" -LoginName $Site.LoginName -Group "Wolfe Pack Members"
        }

    #If operator either does not accept or if the word Accept is not typed correctly acknowledging the entry disclaimer.
    Else
        {
            Write-Host "Either the acknowledgement input does not match the word Accept or you have not agreed to accept the risk of running this script."
            Start-Sleep -Seconds 1
            Write-Host "The script will now exit. Have a nice day!"
            Exit
        }
