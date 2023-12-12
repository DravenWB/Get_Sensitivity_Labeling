###################################################################################################################
# Purpose: This script is intended to reset the inheritance state of all items within a document library or a list.
#
# Development date: December 12, 2023
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

####################################################################################################################################################################################
#Run check against PowerShell version and verify that it is version 7 or greater. If not, inform the user and exit the script.

Write-Host ""
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
        else #Inform the operator that the correct version required is not installed and need to be run in PowerShell 7.
            {
                Write-Host ""
                Write-Host -ForegroundColor Red "The currently running PowerShell version is $InstalledVersion."
                Write-Host -ForegroundColor Red "This PowerShell script requires PowerShell version 7 or greater."
                Write-Host -ForegroundColor Red "Please run in PowerShell 7 and try again."
                Start-Sleep -Seconds 3
                Exit
            }

####################################################################################################################################################################################
#Complete modulecheck for PnP.

#Inform the operator of module check.
Write-Host ""
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

####################################################################################################################################################################################
#Set Variables via operator input.

Write-Host ""
Write-Host "Please enter the URL of the SharePoint site you are seeking to re-inherit items on:"
Write-Host "Ex: https://contoso.sharepoint.com/sites/SiteName"
Write-Host "Ex: https://contoso.sharepoint.us/sites/SiteName"
Write-Host ""

$SiteURL = Read-Host "Site URL"

Write-Host ""
Write-Host "Please enter the name of the document library or list you wish to re-inherit items in:"
Write-Host "The name of the document library or list should be the plain text display name and not the one found in the URL."
Write-Host "Example: Documents"
Write-Host ""

$ListName = Read-Host "Library"

Write-Host ""
Write-Host "Please enter the file path for the folder you are completing resets for. This is a logical path and not a precise URL path."
Write-Host "Ex: /sites/Sitename/Shared Documents/Folder Name/"
Write-Host ""
Write-HOst "Requirements:"
Write-Host "  - It is important that you add a / at the end of the directory to avoid targetting the main library."
Write-Host "  - The folder must start at /sites."
Write-Host "  - The library name must be presented as it does in the URL with spaces in place of %20 as demonstrated above."
Write-Host ""

$RelativePath = Read-Host "Target"

####################################################################################################################################################################################

#Connect to PnP Online and exit if it fails.
Write-Host ""

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
#Preparation for primary operations.

#Get the context.
$Context = Get-PnPContext

#Inform the user.
Write-Host ""
Write-Host -ForegroundColor Green "Now getting document index..."
Start-Sleep -Seconds 1

#Get all files from the library - In batches of 500
$ListItems = Get-PnPListItem -List $ListName -PageSize 500 | Where {$_.FileSystemObjectType -eq "File" -or $_.FileSystemObjectType -eq "Folder"}

$ProcessingIndex = @() #Defines index to store filtered objects based on file path for processing.

####################################################################################################################################################################################
#Loop to cycle through each library item and sends filtered objects (of correct file path) to the processing index.

Write-Host ""
Write-Host -ForegroundColor Green "Now filtering index for selected directory..."
Start-Sleep -Seconds 1

$FilteringCounter = 0 #Initializes counter for filtering progress bar. 

foreach ($Item in $ListItems)
{
    #Update filtering progress bar.
    $ProgressPercent = ($FilteringCounter / $ListItems.Count) * 10 #Calculate display percentage.
    Write-Progress -Activity "Now Filtering SharePoint Files for Reset..." -Status "$ProgressPercent% Complete" -PercentComplete $ProgressPercent

    #Sets a variable to the relative path value of the currently processed item.
    $FilterObject = ($Item.FieldValues.FileRef).ToString()

    #Checks the relative path to confirm if it is a match (including sub-directories/files).
    if ($FilterObject.StartsWith($RelativePath))
        {
            #If the file path matches 
            $ProcessingIndex += $Item                 
        }
}

$ListItems = $null #Clear full list of documents in the document library to save memory as it is no longer needed.
$ProgressPercent = $null #Clear progress percent for next use.

####################################################################################################################################################################################
#Review currently filtered items and allow operator to review changes to be made prior to running.

Write-Host -ForegroundColor Yellow "##############################################################################################################"
Write-Host "" #Spacer
Write-Host                         "Now that item filtering has been completed, please review the below details to ensure you'd like to proceed:  "
Write-Host "" #Spacer
Write-Host                         "- Operation: Reset of file/folder permission inheritance to root directory. (Most Commonly: Document Library)"
Write-Host                         "- Number of files/folders:" $ProcessingIndex.Count
Write-Host "" #Spacer
Write-Host                         "Directories for processing:"
Write-Host                         "============================================"
                                   $ProcessingIndex.FieldValues.FileDirRef | Sort-Object | Get-Unique
Write-Host "" #Spacer
Write-Host -ForegroundColor Yellow "##############################################################################################################"

do
    {
        Write-Host "How would you like to proceed?"
        Write-Host ""
        Write-Host "1. Re-print operation summary details."
        Write-Host "2. Save all current details to CSV spreadsheet for highly detailed review."
        Write-Host "3. Proceed with file/folder inheritance reset."
        Write-Host "4. Stop and exit the script completely, clearing primary script memory, without making any changes."
        Write-Host ""
        Write-Host "Note: This menu will repeat until you select option 3 or exit the script entirely via option 4."

        $ReviewSelect = Read-Host "Selection"

        switch($ReviewSelect)
            {
                '1' #Re-print operation summary details.
                {
                    Write-Host -ForegroundColor Yellow "##############################################################################################################"
                    Write-Host "" #Spacer
                    Write-Host                         "Now that item filtering has been completed, please review the below details to ensure you'd like to proceed:  "
                    Write-Host "" #Spacer
                    Write-Host                         "- Operation: Reset of file/folder permission inheritance to root directory. (Most Commonly: Document Library)"
                    Write-Host                         "- Number of files/folders:" $ProcessingIndex.Count
                    Write-Host "" #Spacer
                    Write-Host                         "Directories for processing:"
                    Write-Host                         "============================================"
                                                       $ProcessingIndex.FieldValues.FileDirRef | Sort-Object | Get-Unique
                    Write-Host "" #Spacer
                    Write-Host -ForegroundColor Yellow "##############################################################################################################"
                    Write-Host "" #Spacer
                    Read-host -Prompt "Enter any key to continue" #Makes the script wait till the user is ready to continue.
                }

                '2' #Save all current details to CSV spreadsheet for a more detailed review.
                {
                    Write-Host ""
                    Write-Host -ForegroundColor Green "Now formatting data for output..."
                    Start-Sleep -Seconds 2 #Gives operator time to read message.

                    $ReviewIndex = @() #Initialize to save formatted objects.
                    $ReviewCounter = 0 #Initialize counter for progress bar and item indexing.

                    class ReviewData #Initialize class for data output.
                    {
                        [int]    $Index
                        [string] $Date
                        [string] $ItemType
                        [string] $Name
                        [string] $Path
                        [string] $Created
                        [string] $Modified 
                    }

                    #Cycle through each object filtered for processing, details for output in CSV and send to index.
                    foreach ($ReviewItem in $ProcessingIndex)
                        {
                            $ProgressPercent = ($ReviewCounter / $ProcessingIndex.Count) * 10 #Calculate display percentage.
                            Write-Progress -Activity "Now formatting review data..." -Status "$ProgressPercent% Complete" -PercentComplete $ProgressPercent

                            $DataTable = New-Object -TypeName ReviewData -Property $([Ordered]@{
    
                            Index = $ReviewCounter + 1
                            Date = Get-Date -Format "MM/dd/yyyy"
                            ItemType = $ReviewItem.FileSystemObjectType
                            Name = $ReviewItem.FieldValues.FileLeafRef
                            Path = $ReviewItem.FieldValues.FileDirRef
                            Created = $ReviewItem.FieldValues.Created
                            Modified = $ReviewItem.FieldValues.Modified
                            })

                            $ReviewIndex += $DataTable
                            $ReviewCounter++
                            $DataTable = $null #Clear current item for next use.
                        }

                    #Output indexed review data to CSV.

                    try
                        {
                            $ReviewIndex | Export-Csv -Path "~\Desktop\File_Reinheritance_Review.csv" -NoClobber
                            Write-Host -ForegroundColor Green "Your file was saved to the desktop as: > File_Reinheritance_Review.csv <"
                            Read-host -Prompt "Enter any key to continue" #Makes the script wait till the user is ready to continue.
                        }

                        catch #If file fails to save, append time for filename conflicts.
                            {

                                try
                                    {

                                        $ReviewIndex | Export-Csv -Path "~\Desktop\File_Reinheritance_Review_New.csv" -NoClobber
                                        Write-Host -ForegroundColor Yellow "A file with the same name was found on your desktop so the name was modified."
                                        Write-Host -ForegroundColor Green "The file was saved to the desktop as: > File_Reinheritance_Review_New.csv <"
                                        Read-host -Prompt "Enter any key to continue" #Makes the script wait till the user is ready to continue.
                                    }

                                    catch #If saving STILL fails, print error to operator.
                                        {
                                            Write-Host -ForegroundColor Red "There was an error saving the review data: $_"
                                            Read-host -Prompt "Enter any key to continue" #Makes the script wait till the user is ready to continue.
                                        }
                            }
                    
                    $ReviewIndex = $null #Clear review index memory upon completion.
                    $DataTable = $null #Clear data table memory for next use.
                    $ProgressPercent = $null #Clear value for next use.
                }

                '4' #Stop and exit the script completely without making any changes.
                {
                    Write-Host -ForegroundColor Green "Now exiting the script..."
                    Start-Sleep -Seconds 1
                    Write-Host -ForegroundColor Green "Clearing primary script used memory..."

                    $ProcessingIndex = $null
                    $ReviewIndex = $null

                    Write-Host -ForegroundColor Green "Have a wonderful day! :)"
                    Start-Sleep -Seconds 1
                    Exit
                }

            }
    }

    until($ReviewSelect -eq '3')

####################################################################################################################################################################################
#Run inheritance reset.

#Inform the operator.
Write-Host ""
Write-Host -ForegroundColor Green "Now proceeding to execution of file/folder inheritance within: $RelativePath"
Start-Sleep -Seconds 1

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

#Define the data class to store changes.
class InheritanceChange
{
    [int]    $Index 
    [string] $Date
    [string] $Time 
    [string] $FileName
    [string] $Location
    [bool]   $Inheritance_Reset
    [string] $Errors
}

$LoggingIndex = @() #Define data index to store changes for later output to log CSV file.
$ProcessingCounter = 0 #Initialize counter for change indexing in output file and progress bar.
$ProcessingDate = Get-Date -Format "MM/dd/yyyy" #Pre-assigned to get date once instead of potentially hundreds/thousands of times over.

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#
#Process the inheritance resets.

foreach($ProcessingItem in $ProcessingIndex)
    {
        #Update the progress bar.
        $ProgressPercent = ($ProcessingCounter / $ProcessingIndex.Count) * 10 #Calculate display percentage.
        Write-Progress -Activity "Now Processing File/Folder Inheritance Resets..." -Status "$ProgressPercent% Complete" -PercentComplete $ProgressPercent

        #Process the inheritance reset.
        try
        {
            $ProcessingItem.ResetRoleInheritance(); #Command to prime item inheritance reset.
            $Context.ExecuteQuery() #Command to execute item inheritance reset.

            $ItemName = $ProcessingItem.FieldValues.FileLeafRef #Used to output name of item to let the operator know it was processed.
            Write-Host ""
            Write-Host -ForegroundColor Green "$ItemName Role inheritance reset"

            $ResetCheck = $true #Sets variable for output if successful.
        }

            catch #Error handling.
            {
                Write-Host -ForegroundColor Red "Error recorded for the resetting the role inheritance of $ItemName" ":" $_
                $XError = "$_"

                $ResetCheck = $false #Sets variable for output if reset fails.
            }

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

        #Write data to instantiated class object for temporary storage and file output.
        $DataTable = New-Object -TypeName InheritanceChange -Property $([Ordered]@{
        
        Index = $ProcessingCounter + 1
        Date = $ProcessingDate
        Time = Get-Date -Format "HH:mm"
        FileName = $ItemName
        Location = $ProcessingItem.FieldValues.FileDirRef
        Inheritance_Reset = $ResetCheck
        Errors = $XError
        })

        #Send the data table to the index.
        $LoggingIndex += $DataTable

        #Increment counter for progress bar and indexing.
        $ProcessingCounter++

        #Clear variables for next use to ensure no duplicate values from previous items are used.
        $DataTable = $null
        $ItemName = $null
        $ResetCheck = $null
        $XError = $null
    }

####################################################################################################################################################################################

#Save script data to file.

do
    {
        Write-Host "This portion of the script has been placed into a loop in case saving the file fails."
        Write-Host ""
        Write-Host "Save defaults to the file name Inheritance_Reset_Log.csv on the desktop. If it fails, it will modify the name as a backup."
        Write-Host "Once the script completes, changes made that have been stored in memory will be lost."
        Write-Host ""
        Write-Host "Please ensure you have the data you need before exiting."
        Write-Host " "
        Write-Host "1. Attempt to save file."
        Write-Host "2. Complete cleanup and Exit."
        Write-Host ""

        $SaveSelection = Read-Host "Selection"

        switch($SaveSelection)
            {
                '1' #Attempt to save file with default settings.
                {
                    try
                        {
                            $LoggingIndex | Export-Csv -Path "~\Desktop\Inheritance_Reset_Log.csv" -NoClobber
                            Write-Host -ForegroundColor Green "The file has successfully been saved to the following location:"
                            Write-Host -ForegroundColor Green "> ~\Desktop\Inheritance_Reset_Log.csv <"
                            Read-host -Prompt "Enter any key to continue" #Makes the script wait till the user is ready to continue.
                        }

                        catch
                            {
                                Write-Host -ForegroundColor Yellow "Saving the file failed. Now attempting to add modifier and try again."
                                Start-Sleep -Seconds 2

                                $LoggingIndex | Export-Csv -Path "~\Desktop\Inheritance_Reset_Log_New.csv" -NoClobber
                                Write-Host -ForegroundColor Green "The file has successfully been saved to the following location:"
                                Write-Host -ForegroundColor Green "> ~\Desktop\Inheritance_Reset_Log_New.csv <"
                                Read-host -Prompt "Enter any key to continue" #Makes the script wait till the user is ready to continue.
                            }
                }
            }
    }

    until($SaveSelection -eq '2')

####################################################################################################################################################################################
#Clean up.

#Reset error view/length to operator original setting.
$ErrorView = [System.Management.Automation.ActionPreference]::$OriginalErrorView
$ErrorActionPreference = [System.Management.Automation.ActionPreference]::$OriginalErrorAction

$LoggingIndex = $null

####################################################################################################################################################################################

Write-Host -ForegroundColor Green "Script has completed all operations! Have a wonderful day! :)"
