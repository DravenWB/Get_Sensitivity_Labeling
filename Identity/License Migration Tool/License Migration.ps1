##############################################################################
# - Script Definition: This script is designed to change the licenses of a
#   large amount of users as defined by the operator and then output all
#   changes to a CSV file to include users that received an error during license
#   change oeration.
#
# - Disclaimer: This script is not officially supported by Microsoft, its
#   affiliates or partners. This script is provided as is and the responsibility
#   of understanding the scripts, functions, and operations falls upon those that
#   may choose to run it. Responsibility of positive or negative outcomes of
#   this script may not receive future assistance as such.
##############################################################################

$LicenseToAdd = "ENTERPRISEPACK_USGOV_GCCHIGH"  #Variable that sets Office 365 E3 License for assignemnt.
$LicenseToRemove = "STANDARDPACK_USGOV_GCCHIGH"  #Variable that removes Office 365 E1 License.
$SavePathAndName = "~\Desktop\Tenant_License_Migration_Report.csv"
$LicenseChangeIndex = @() #Array to store data. Do not change.

### DO NOT CHANGE ###
#Set error action type to make error catching work.
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
#####################

Write-Host ""
Write-Host -ForegroundColor Green "Now checking installed Microsoft Graph PowerShell Module."
Write-Host ""

Start-Sleep -Seconds 1 #Gives operator a moment to read message.

    if(-not (Get-Module Microsoft.Graph -ListAvailable)) #Check for and install Microsoft Graph module if not installed.
    {
        Write-Host -ForegroundColor DarkYellow "Microsoft Graph (Microsoft.Graph) module not found. Now installing..."
        Start-Sleep -Seconds 1 #Gives operator a moment to read message.

        try
        {
            Install-Module Microsoft.Graph -Scope CurrentUser -Force
        }
            catch
            {
                Write-Host -ForegroundColor Red "There was an error installing the Microsoft.Graph Commandlet: $_"
                exit
            }

    }

Write-Host -ForegroundColor Green "Module check complete!"
Start-Sleep -Seconds 1 #Gives operator a moment to read message.

#Determine user's environment
Write-Host "Please select the tenant type for licenses being modified:"
Write-Host "1. Commercial (.com)"
Write-Host "2. GCC-High (.us)"

$TenantSelection = Read-Host "Selection:" #Get the user tenant selection.

switch($TenantSelection) #Assign user selection for later use.
{
    '1' {$TenantType = "Commercial"}
    '2' {$TenantType = "GCCH"}
}

#Connect to Graph services.
Write-Host -ForegroundColor Green "Now connecting to Graph services."

Start-Sleep -Seconds 1

if ($TenantType -like "Commercial") #Connect to the graph API for license modification.
    {Connect-MgGraph -Scopes User.ReadWrite.All, Organization.Read.All -NoWelcome} #Connects to the commercial graph service.

    else 
    {Connect-MGgraph -Scopes User.ReadWrite.All, Organization.Read.All -Environment USGov -NoWelcome} #Connects to the GCC-High graph service.

 #Define object oriented object variables for data storage.
    class LicenseChangeMatrix
    {
        [string] ${Date}
        [string] ${Time}
        [string] ${UserUPN}
        [string] ${OriginalLicense}
        [string] ${NewLicense}
        [string] ${Errors}
    }

#Convert license type to sku number
$RemoveLicenseSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq "$LicenseToRemove"
$AddLicenseSku = Get-MgSubscribedSku -All | Where SkuPartNumber -eq "$LicenseToAdd"

#Call users with selected license for removal and store in variable.
$Users = Get-MgUser -Filter "assignedLicenses/any(x:x/skuId eq $($RemoveLicenseSku.SkuId) )" -ConsistencyLevel eventual -All

do
{
    Write-Host "Please select an option:"
    Write-Host "1. Execute license migration changes."
    Write-Host "2. Write license changes to file."
    Write-Host "3. Exit."
    
    $MainMenu = Read-Host "Selection"

    switch($MainMenu)
    {
        '1'
        {
            Write-Host -ForegroundColor DarkYellow "The following amount of changes are about to be made:" $Users.count
            Write-Host -ForegroundColor DarkYellow "To make changes, please confirm by typing the word > Execute < (Case Sensitive):"
            $Confirmation = Read-Host "Confirmation:"

            If($Confirmation -ceq "Execute")
            {
                #Cycle through users storing users with licensing in variable
                foreach ($User in $Users)
                {
                    $Upn = $User.UserPrincipalName

                        #Write user converting from license to license
                        Write-Host -ForegroundColor Green "User $Upn will go from $LicenseToRemove to $LicenseToAdd"

                        $XError = " "

                        try
                        {
                        #Remove old license and add new license.
                        Set-MgUserLicense -UserId "$Upn" -AddLicenses @{SkuId = $AddLicenseSku.SkuId} -RemoveLicenses @($RemoveLicenseSku.SkuId)
                        }

                            catch #If removal of license fails.
                            {
                                #Assign error to variable for both host output and storage to csv file.
                                $XError = "Couldn't replace $Upn's $LicenseToRemove with $LicenseToAdd"
                                Write-Host -ForeGroundColor Red $XError #Print the error to the screen.
                            }

                        #Apply change data to object.
                        $Object = New-Object PSObject -Property $([ordered]@{
                    
                        Date = Get-Date -Format "MM/dd/yyyy"
                        Time = Get-Date -Format "HH:mm"
                        UserUPN = $Upn
                        OriginalLicense = $LicenseToRemove
                        NewLicense = $LicenseToAdd
                        Errors = $XError
                        })

                        $LicenseChangeIndex += $Object #Send object to global index.

                        $XError = $null #Clear error variable to ensure error is not duplicated on next object.
                        $Oject = $null #Clear error variable to ensure error is not duplicated on next object.
                        Write-Host "" #Spacer
                }
            }

                else #If confirmation fails. Operator must enter the word "Execute" precisely.
                {
                    Write-Host -ForegroundColor Red "Confirmation failed. Please try again."
                    Start-Sleep -Seconds 2
                }
        }

        '2'
        {
            #Combine path and filename for creation.
            try
            {
            #Export all items to a CSV file.
            $LicenseChangeIndex | Export-Csv -Path $SavePathAndName -NoClobber
            }
                catch
                {
                    Write-Host -ForegroundColor Red "There was an error outputting the gathered data to CSV. Attempting again with modifier..."

                    try
                    {
                        $TimeModifier = Get-Time "MM/dd/yyy - HH:mm"
                        $SavePathAndName = $SavePathAndName + $TimeModifier
                        $LicenseChangeIndex | Export-Csv -Path $SavePathAndName -NoClobber
                    }
                        catch
                        {
                            Write-Host -ForegroundColor Red "Failed to save file with modifer due to error: $_"
                        }
                }
        }
    }
}
until($MainMenu -eq '3')

Write-Host -ForegroundColor Green "All operations complete!"
