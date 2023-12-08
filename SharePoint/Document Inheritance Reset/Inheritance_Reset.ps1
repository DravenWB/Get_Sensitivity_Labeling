###################################################################################################################
# Purpose: This script is intended to reset the inheritance state of all items within a document library or a list.
#
# Development date: July 26, 2023
#
# Note: This version of the script resets the inheritance status of ALL items and is not yet modified to only
# affect items that have unique permissions.
#
# NOTICE: This script has been carefully written and thoroughly tested. However, it remains an experimental solution.
# By running this script, you accept the risks and responsibilities associated with running said code. Microsoft
# is not liable for any damages or resulting issues.
###################################################################################################################

#Operator risk acknowledgemenet initialization.
$OperatorAcknowledgement = " "

Write-Host -ForegroundColor DarkYellow "Disclaimer: This script is not officially supported by Microsoft, its affiliates or partners"
Write-Host -ForegroundColor DarkYellow "This script is provided as is and the responsibility of understanding the scripts functions and operations falls upon those that may choose to run it."
Write-Host -ForegroundColor DarkYellow "Positive or negative outcomes of this script may not receive future assistance as such."
Write-Host -ForegroundColor DarkYellow ""
Write-Host -ForegroundColor DarkYellow "To acknowledge the above terms and proceed with running the script, please enter > Accept < (Case Sensitive)."

$OperatorAcknowledgement = Read-Host "Acknowledgement"

if ($OperatorAcknowledgement -cne "Accept") #If operator acknowledgement check is not matched to "Accept", exit the script.
{
    Exit
}

####################################################################################################################################################################################

#Run check against PowerShell version and verify that it is version 7 or greater. If not, inform the user and exit the script.

Write-Host "Now checking running PowerShell version..."
Start-Sleep -Seconds 1

$InstalledVersion = ($PSVersionTable.PSVersion).ToString()

if ($InstalledVersion -ge '7')
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

#Complete modulecheck for PnP.
    if (Get-Module -ListAvailable -Name "PnP.PowerShell")
        {
            Write-Host -ForegroundColor Green "The PnP PowerShell Module is confirmed as installed."
            Start-Sleep -Seconds 1
        }

         else #If module not found, attempt to install the module.
            {
                try
                {
                    Write-Host -ForegroundColor DarkYellow "PnP PowerShell Module not found. Now attempting to install the module."
                    Install-Module -Name PnP.PowerShell -Scope CurrentUser
                    Start-Sleep -Seconds 1
                    Import-Module -Name PnP.PowerShell -Scope Local
                }

                    catch
                    {
                        Write-Host -ForegroundColor Red "Failed to install the PnP PowerShell Module due to error:" $_
                        Exit
                    }
            }

####################################################################################################################################################################################

#Set Variables via operator input.
Write-Host "Please enter the URL of the SharePoint site you are seeking to re-inherit items on:"
Write-Host "Ex: https://contoso.sharepoint.com/sites/SiteName"

$SiteURL = Read-Host "Site URL"

Write-Host "Please enter the site display name."
Write-Host "Example: Test Site"

$SiteName = Read-Host "Site Name"

Write-Host "Please enter the name of the document library or list you wish to re-inherit items in:"
Write-Host "The name of the document library or list should be the plain text display name. Example: Documents"

$ListName = Read-Host "List/Library"

Write-Host "Please enter the relative file path for the folder you are completing resets for. This is a logical path and not a literal one."
Write-Host "Ex: /sites/Sitename/lists/Documents/Folder 1/"
Write-Host "Note: The folder must start at /sites and include the lists variable."

$RelativeFolder = Read-Host "Target"

$RelativeFolder = $RelativeFolder + "*"

####################################################################################################################################################################################

#Connect to PnP Online and exit if it fails.

try
    {
        Connect-PnPOnline -Url $SiteURL -UseWebLogin
    }
        catch
            {
                Write-Host -ForegroundColor Red "There was an error connecting to PnP Online: $_"
                Exit
            }

####################################################################################################################################################################################

#Set error view and action for clean entry into the output file. Additionally gets the operator's current setting to change it back at script cleanup time. 

#Get operator current error output length and set to concise.
$OriginalErrorView = $ErrorView
$ErrorView = [System.Management.Automation.ActionPreference]::ConciseView

#Get operator current error action and set to stop.
$OriginalErrorAction = $ErrorActionPreference
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue

####################################################################################################################################################################################
#Preparation for primary operation.

#Get the context.
$Context = Get-PnPContext

#Get All Files from the selected location - In batches of 500
$ListItems = Get-PnPListItem -List $ListName -PageSize 500 | Where {$_.FileSystemObjectType -eq "File"}

$LoggingIndex = @() #Instantiate data index to store objects.

#Initialize the data class.
class InheritanceChange
{
    [string] $Date
    [string] $Time 
    [string] $FileName
    [string] $Location
    [string] $InheritanceReset
    [string] $Errors
}

####################################################################################################################################################################################

#Loop to cycle through each item and reset the context. Includes error handling.
foreach ($Item in $ListItems)
{
    if ($Item.FieldValues.FileRef -match $RelativeFolder) #Check if file has the same server relative file path as those targeted.
        {
            try
            {
                $Item.ResetRoleInheritance(); #Command to prime item inheritance reset.
                $Context.ExecuteQuery() #Command to execute item inheritance reset.

                $ItemName = $Item.FieldValues.FileLeafRef #Used to output name of item being processed.
                Write-Host -ForegroundColor Green "$ItemName Role inheritance reset"
            }
                catch #Error handling.
                {
                    Write-Host -ForegroundColor Red "There was an error with resetting the role inheritance of $ItemName" ":" $_
                    $XError = "$_"
                }
        }

        #Write data to instantiated class object for temporary storage and file output.
        $DataTable = New-Object -TypeName InheritanceChange -Property $([Ordered]@{
    
        Date = Get-Date -Format "MM/dd/yyy"
        Time = Get-Date -Format "HH:mm"
        FileName = $ItemName
        Location = $Item.FieldValues.FileRef
        Errors = $XError
        })

        #Send the data table to the index.
        $LoggingIndex += $DataTable

        #Clear data table and increment counter for next site.
        $DataTable = $null
} 

####################################################################################################################################################################################

#Save script data to file.
$SaveModifier = Get-Date -Format "MM/dd/yyy"

do
    {
        Write-Host "This portion of the script has been placed into a loop in case saving the file fails."
        Write-Host "Save defaults to the file name Inheritance_Reset_Log.csv on the desktop. If it fails, it will try to append the date as a backup."
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
                            $LoggingIndex | Export-Csv -Path "~\Desktop\Inheritance_Reset_Log.csv" -NoClobber
                            Write-Host -ForegroundColor Green "The file has successfully been saved to the following location:"
                            Write-Host -ForegroundColor Green "> ~\Desktop\Inheritance_Reset_Log.csv <"
                        }

                        catch
                            {
                                Write-Host -ForegroundColor Yellow "Saving the file failed. Now attempting to add modifier and try again."
                                Start-Sleep -Seconds 3

                                $LoggingIndex | Export-Csv -Path "~\Desktop\Inheritance_Reset_Log$SaveModifier.csv" -NoClobber 
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
#Clean up.

#Reset error view/length to operator original setting.
$ErrorView = [System.Management.Automation.ActionPreference]::$OriginalErrorView
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::$OriginalErrorAction

$LoggingIndex = $null

####################################################################################################################################################################################

Write-Host -ForegroundColor Green "Script has completed all operations! Have a wonderful day! :)"
