##############################################################################
# - Script Definition: This script is designed to gather sensitivity labeling
#   information and surrounding details for analysis of issues surrounding
#   labels and label policies.
#
# - Development Date: November 19, 2023
#
# - Disclaimer: This script is not officially supported by Microsoft, its
# affiliates or partners. This script is provided as is and the responsibility
# of understanding the scripts functions and operations falls upon those that
# may choose to run it. Positive or negative outcomes of this script may not 
# receive future assistance as such.
##############################################################################

#Variable presets.
$ExchangeEnvironmentName = "O365USGovGCCHigh" #Set the environment connection type if required.

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
    #Start the PowerShell transcript, to be saved to a file on the desktop using the above name.
    Write-Host -ForegroundColor Green "Now starting PowerShell transcript."
    Start-Sleep -Seconds 1
    Start-Transcript

    #Get the operator UPN.
    Write-Host "Please enter your administrator account UPN for connecting."
    $AdministratorUPN = Read-Host "UPN"

    Write-Host "Please enter the UPN of the user to run diagnostics for."
    $UserUPN = Read-Host "User UPN"

    #Check if the exchange online management module is installed and loaded.
    Write-Host -ForegroundColor Green "Now checking for Exchange Online Management Shell installation status."
    Start-Sleep -Seconds 1

    if (Get-Module -ListAvailable -Name "ExchangeOnlineManagement")
        {
            Write-Host -ForegroundColor Green "The Exchange Online Management shell is confirmed as installed."
            Start-Sleep -Seconds 1
        }

         else #If module not found, attempt to install the module.
            {
                try
                {
                    Write-Host -ForegroundColor DarkYellow "Exchange Online Management shell not found. Now attempting to install the module."
                    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
                    Start-Sleep -Seconds 1
                    Import-Module -Name ExchangeOnlineManagement -Scope Local
                }

                    catch
                    {
                        Write-Host -ForegroundColor Red "Failed to install the Exchange Online Module due to error:" $_
                    }
            }

    #============================================================================================================================================================================#

    #Connect to services required.
    Write-Host -ForegroundColor Green "Now attempting to connect to ExchangeOnline and IPPSSession..."
    Start-Sleep -Seconds 1

    #Attempt to connect to the Exchange Online service and exit if connection fails as it is required for the script.
    try
        {
            Connect-ExchangeOnline -UserPrincipalName $AdministratorUPN -ExchangeEnvironmentName $ExchangeEnvironmentName
        }
            catch
                {
                    Write-Host -ForegroundColor Red "Failed to connect to Exchange Online due to error:" $_
                    Exit
                }

    #Attempt to connect to the Security and Compliance PowerShell and exit if fails as it is required for the script.
    try
        {
            Connect-IPPSSession -UserPrincipalName $AdministratorUPN
        }
            catch
                {
                    try #If standard connection fails, attempt to use GCCH Connection URI.
                        {
                           Connect-IPPSSession -UserPrincipalName $AdministratorUPN -ConnectionUri "https://ps.compliance.protection.office365.us/powershell-liveid/" 
                        }

                        catch
                            {
                                Write-Host -ForegroundColor Red "Failed to connect to the Security and Compliance PowerShell due to error:" $_
                                Exit
                            }
                }

#Get user current formatenumerationlimit for restoration upon script completion then set enumeration limit for script.
$InitEnumLimit = $formatenumerationlimit
$formatenumerationlimit=-1

#Get mailbox properties.
Get-Mailbox -Identity $UserUPN | Format-List -ErrorAction SilentlyContinue

#Get all OME configurations.
Get-OMEConfiguration | Format-List

#Get RMS policy templates.
Get-RMSTemplate -ResultSize Unlimited

#Get all labeling policies
Get-LabelPolicy | Format-List

#Get all sensitivity labels for the organization
Get-Label | Format-List

#Script cleanup.
$formatenumerationlimit = $InitEnumLimit

Write-Host -ForegroundColor Green "Script has completed all operations."
Write-Host -ForegroundColor Green "Have a great day! :)"
Stop-transcript
}

    #If operator either does not accept or if the word Accept is not typed correctly acknowledging the entry disclaimer.
    Else
        {
            Write-Host "Either the acknowledgement input does not match the word Accept or you have not agreed to accept the risk of running this script."
            Start-Sleep -Seconds 1
            Write-Host "The script will now exit. Have a nice day!"
            Exit
        }
