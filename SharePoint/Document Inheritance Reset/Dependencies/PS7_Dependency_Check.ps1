#Checks to ensure that PowerShell 7 or greater is installed. If not, attempt installation
function PS7_Version_Check
{
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

        else #Inform the operator that the correct version required is not installed and need to be run in PowerShell 7. Exit script upon completion.
            {
                Write-Host ""
                Write-Host -ForegroundColor Red "The currently running PowerShell version is $InstalledVersion."
                Write-Host -ForegroundColor Red "This PowerShell script requires PowerShell version 7 or greater."
                Write-Host -ForegroundColor Red "Please run in PowerShell 7 and try again."
                Start-Sleep -Seconds 3
                Exit
            }
}

#Checks to see if PnP.PowerShell is installed. If not, attempt to install it. If fails, exit script.
function PnP_Installation_Check
{
    #Inform the operator of module check.
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
}
