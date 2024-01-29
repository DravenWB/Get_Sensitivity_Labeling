###################################################################################################################
# Purpose: This script is intended to reset the inheritance state of all items within a document library or a list.
#
# Development date: January 27, 2024
#
# Note: This version of the script resets the inheritance status of ALL items and is not yet modified to only
# affect items that have unique permissions.
#
# NOTICE: This script has been carefully written and thoroughly tested. However, it remains an experimental solution.
# By running this script, you accept the risks and responsibilities associated with running said code. Microsoft
# is not liable for any damages or resulting issues.
###################################################################################################################

#Get dependent function modules.
. .\Dependencies\Operator_Acknowledgement.ps1
. .\Dependencies\PS7_Dependency_Check.ps1
. .\Error_Handling_Config.ps1
. .\File_Output_Menu.ps1

#----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------#

# Initial Operations.
Write-Host ""
Operator_Acknowledgement_Check #Confirm user acknowledgement of operating script.
Write-Host ""
PS7_Version_Check #Check the version of PowerShell 7 and ensure it is the current version running.
Write-Host ""
PnP_Installation_Check #Check if PnP.PowerShell is installed.
Write-Host ""
Error_Action_Initialization #Copies user error action settings and configures required temporary settings for script.

####################################################################################################################################################################################
#Set Variables via operator input.

Write-Host ""
Write-Host "Please enter the full Logical URL of your target document library (or library folder)."
Write-Host "Be sure to include quotation marks at the start and end of your URL as spaces can be problematic."
Write-Host "This URL is part logical and is not a complete literal URL copied from the address bar."
Write-Host "Ex: https://contoso.sharepoint.com/sites/SiteName/Shared Documents/Folder/SubFolder/"
Write-Host "Ex: https://contoso.sharepoint.us/sites/SiteName/Shared Documents/"
Write-Host ""

$FullURL = Read-Host "Site Logical URL" #Get full logical URL from user.

$UrlDelimit = $FullURL -split '/' #Set delimiter and split the URL.

#Grab the tenant URL, Library and relative path.
$RawTenant = $UrlDelimit[2] #Get the tenant URL and append necessary characters for connection.
$ListName = $UrlDelimit[5] #Get the list/library name from the target relative URL.
$RelativePath = '/' + [string]::Join('/', $UrlDelimit[3..($UrlDelimit.Length - 2)]) #Get the relative path for filtering.

#Modify the raw tenant to a full URL for connection.
$SiteURL = "https://$RawTenant"

#Ensure that the relative URL ends with '/' to prevent whole library targeting.
if (-not $RelativePath.EndsWith('/')){$RelativePath += '/'}

####################################################################################################################################################################################
#Connect to PnP Online and exit if it fails. If succeeds, get the context and proceed.
Write-Host ""
try {Connect-PnPOnline -Url $SiteURL -UseWebLogin}

    catch 
        {
            Write-Host -ForegroundColor Red "There was an error connecting to PnP Online: $_"
            Exit
        }

#Get the context.
$Context = Get-PnPContext

####################################################################################################################################################################################
$ExitSwitch = $null #Initialization of property to control exit parameter.

do
    {
        Write-Host "Would you like operational log output? (Default: None if skipped.)"
        Write-Host "1. Enter logging configuration."
        Write-Host "2. Print current logging configuration."
        Write-Host "3. Delete logging configuration."
        Write-Host ""
        Write-Host "4. Continue with current configuration."
        Write-Host "5. Skip operational logging."

        $LoggingConfig = Read-Host "Selection"

        switch($LoggingConfig)
            
            '1'
            {
                #Get file save name.
                Write-Host "Enter a filename:"
                $LoggingFileName = Read-Host "Log Name"

                #Get file save path.
                Write-Host "Enter a save path:"
                Write-Host "Note: Recommend separate, dedicated directory."
                Write-Host "Note: If any spaces are in the path, start and finish the entry with quotation marks."
                $LoggingPath = Read-Host "Path"

                #Get total number of items to process.
                Write-Host "How many items would you like to save per .csv file?"
                Write-Host "Recommended: No more than 10,000 items."
                Write-Host "Input should be a number"

                $LoggingCount = Read-Host "Amount"

                #Ensure that the logging path has a / at the end.
                if (-not $LoggingPath.EndsWith('/'))
                    {
                        $LoggingPath += '/'
                    }

                #Ensure that the logging file name ends in .csv
                if (-not $LoggingFileName.EndsWith('.csv'))
                    {
                        $LoggingFileName += '.csv'
                    }

            }
            
            '2'
            {
                Write-Host "Logging File Name: $LoggingFileName"
                Write-Host "Save Directory: $LoggingPath"
                Write-Host "Log entries per file: $LoggingCount"
            }
            
            '3'
            {
                $LoggingFileName = $null
                $LoggingPath = $null
                $LoggingCount = $null

                Write-Host -ForegroundColor Green "Logging parameters cleared!"
            }
            
            '4'
            {
                Write-Host -ForegroundColor Green "Now continuing with current logging configuration..."
                Start-Sleep -Seconds 2
                $Exit = "Exit"
            }
            
            '5'
            {
                Write-Host -ForegroundColor Yellow "Now skipping operational logging..."

                $LoggingFileName = $null
                $LoggingPath = $null
                $LoggingCount = $null

                Start-Sleep -Seconds 2
                $Exit = $Exit
            }
    }

until($Exit -ceq "Exit")

####################################################################################################################################################################################
#Launch menu for pre-operation review.

do
{
    Write-Host "Enter a selection:"
    Write-Host ""
    Write-Host "1. Execute a CAML query to get the target location file count. (Fast)"
    Write-Host "2. Execute resets!"
    Write-Host "3. Exit script."

    $PreExecution = Read-Host "Selection"

    switch($PreExecution)

        '1'
        {
            # CAML Query to filter files and folders
            $camlQuery = "<View>
                            <Query>
                                <Where>
                                    <Or>
                                        <Eq>
                                            <FieldRef Name='FSObjType'/><Value Type='Integer'>0</Value>
                                        </Eq>

                                        <Eq>
                                            <FieldRef Name='FSObjType'/><Value Type='Integer'>1</Value>
                                        </Eq>
                                    </Or>
                                </Where>
                            </Query>
                            <RowLimit>
                                0
                            </RowLimit>
                        </View>"

            # Execute the query
            $items = Get-PnPListItem -List $ListName -Query $camlQuery

            # Output the number of items returned
            Write-Host "Number of items in : $($items.Count)"    
        }
        
        '3'
        {
            Exit
        }
}

until($PreExecution -eq 2)
####################################################################################################################################################################################

#If operator proceeds, define data classes and process changes.

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
$LoggingPathFull = $LoggingPath += $LoggingFileName

do 
    {
        $ProcessingIndex = Get-PnPListItem -List $ListName -PageSize 500 -Page $ProcessingCounter | Where {$_.FileSystemObjectType -eq "File" -or $_.FileSystemObjectType -eq "Folder"}

        foreach ($ProcessingItem in $ProcessingIndex) 
            {
                try
                    {
                        #Load the item and confirm up to date information.
                        $Context.Load($ProcessingItem)
                        $Context.ExecuteQuery()
            
                        #Process the inheritance reset.
                        $ProcessingItem.ResetRoleInheritance(); #Command to prime item inheritance reset.
                        $Context.ExecuteQuery() #Command to execute item inheritance reset.

                        $CurrentItemName = $ProcessingItem.FieldValues.FileLeafRef #Used to output name of item to let the operator know it was processed.
                        Write-Host ""
                        Write-Host -ForegroundColor Green "$CurrentItemName Role inheritance reset"

                        $ResetCheck = $true #Sets variable for output if successful.
                    }

                        catch #Error handling.
                        {
                            Write-Host -ForegroundColor Red "Error recorded for the resetting the role inheritance of $CurrentItemName" ":" $_
                            $XError = "$_"

                            $ResetCheck = $false #Sets variable for output if reset fails.
                        }

                if ($LoggingFileName -ne $null)
                    {
                        #Write data to instantiated class object for temporary storage and file output.
                        $DataTable = New-Object -TypeName InheritanceChange -Property $([Ordered]@{
        
                        Index = $ProcessingCounter + 1
                        Date = $ProcessingDate
                        Time = Get-Date -Format "HH:mm"
                        FileName = $CurrentItemName
                        Location = $ProcessingItem.FieldValues.FileDirRef
                        Inheritance_Reset = $ResetCheck
                        Errors = $XError
                        })

                        #Send the data table to the index.
                        $LoggingIndex += $DataTable

                        #If the index has processed the selected amount of items, output to file and clear for next file.
                        if ($ProcessingIndex -ge $LoggingCount)
                            {
                                Write-Host -ForegroundColor Green "Outputting current data to file..."

                                #Try to save the current data to file under the selected location.
                                try
                                    {
                                        Export-Csv -InputObject $LoggingIndex -Path $LoggingPathFull -NoClobber
                                    }
                                    
                                    #If saving fails, most commonly due to file name errors, rename the file and output again using the time to avoid duplicates a second time.
                                    catch
                                        {
                                            $Time = Get-Date -Format "mm"
                                            $LoggingPathMod = $LoggingPathFull += $Time += "_Mod"

                                            Export-Csv -InputObject $LoggingIndex -Path $LoggingPathMod -NoClobber

                                            $LoggingIndex = $null
                                        }
                            }

                        #Clear variables for next use to ensure no duplicate values from previous items are used.
                        $DataTable = $null
                        $CurrentItemName = $null
                        $ResetCheck = $null
                        $XError = $null
                    }
            }

        #Increment counter for progress bar, indexing and operational processing.
        $ProcessingCounter++
    } 

    while ($ProcessingIndex.Count -eq $ProcessingCounter)

####################################################################################################################################################################################
#Clean up.
Error_Action_Cleanup #Calls function to restore operator error action settings.

#Clear memory.
$LoggingIndex = $null
$ProcessingIndex = $null

####################################################################################################################################################################################

Write-Host -ForegroundColor Green "Script has completed all operations! Have a wonderful day! :)"
