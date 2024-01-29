function Write_Operation_Data
{
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

                                $LoggingIndex | Export-Csv -Path "~\Desktop\Inheritance_Reset_Log_Mod.csv" -NoClobber
                                Write-Host -ForegroundColor Green "The file has successfully been saved to the following location:"
                                Write-Host -ForegroundColor Green "> ~\Desktop\Inheritance_Reset_Log_New.csv <"
                                Read-host -Prompt "Enter any key to continue" #Makes the script wait till the user is ready to continue.
                            }
                }
            }
    }

    until($SaveSelection -eq '2')
}
