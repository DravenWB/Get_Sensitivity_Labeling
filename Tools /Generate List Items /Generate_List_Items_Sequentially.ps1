###################################################################################################################
# Purpose: This script is intended to generate many items in a list for testing purposes.
#
# WARNING: Carelessness with this script will cause degredation to performance of the tenant it is run on if mass
# amounts of objects are generated.
# 
# Development date: November 26, 2023
#
# Dependencies: Powershell v7 or greater, PnP.Powershell module and .Net Framework 4.8+. These are checked for 
# installation status at the start.
###################################################################################################################

#Check for module installation and correct versioning. If PowerShell version is less than version 7, inform the operator and exit.
Write-Host -ForegroundColor Green "Now checking installed dependencies to ensure they are available and are of the correct version..."
Start-Sleep -Seconds 1

Write-Host -ForegroundColor Green "Now checking installed PowerShell Version..."
Start-Sleep -Seconds 1

if ([string] $PSVersionTable.PSVersion -lt "7")
    {
        Write-Host -ForegroundColor Red "Your PowerShell instance is currently running version:" $PSVersionTable.PSVersion
        Write-Host -ForegroundColor Red "To run this script, please run in or install PowerShell version 7 or greater."
        Start-Sleep -Seconds 3
        Write-Host "Have a great day! :)"
        Exit
    }

    else
        {
            Write-Host -ForegroundColor Green "PowerShell v7+ already installed!"
        }

#Check for installed .net version and exit if version is found to be below version 4.8.
#Version not printed to user as the internal version is not listed as 4.8 (ex.) and is instead listed as a number like "533320" which isn't the most readable.
Write-Host -ForegroundColor Green "Now checking installed .net version..."
Start-Sleep -Seconds 1

[string] $DotNetVersion = Get-ItemPropertyValue -LiteralPath 'HKLM:SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' -Name Release
if ($DotNetVersion -lt "533320")
    {
        Write-Host -ForegroundColor Red "Your .Net framework installed is currently out of date."
        Write-Host -ForegroundColor Red "To run this script, please update your .Net version to 4.8 or greater."
        Start-Sleep -Seconds 3
        Write-Host "Have a great day! :)"
        Exit
    }

    else
        {
            Write-Host -ForegroundColor Green "Your .Net version is up to date!"
        }

#Check for installed PnP.PowerShell version and install if missing.
Write-Host -ForegroundColor Green "Now checking installed PnP.PowerShell Version..."
Start-Sleep -Seconds 1

if (-not(Get-Module -ListAvailable -Name "PnP.PowerShell"))
    {
        try
            {
                Write-Host -ForegroundColor DarkYellow "PnP.PowerShell not found. Now proceeding with installation for the current user only..."
                Install-Module -Name "PnP.PowerShell" -Scope CurrentUser
            }

            catch
                {
                    Write-Host -ForegroundColor Red "There was an error installing PnP.Powershell: $_"
                    Exit
                }
    }

    else
        {
            Write-Host -ForegroundColor Green "PnP.PowerShell found!"
        }

#Update the operator.
Write-Host -ForegroundColor Green "Dependency check complete!"
Start-Sleep -Seconds 1

#Get variables from the operator.

#Get the site to generate items for.
Write-Host "Please enter the full site URL you would like to generate items for:"
Write-Host "Ex: https://contoso.sharepoint.com/sites/SiteName/"
$SiteURL = Read-Host "Site"

#Get the list to generate items in.
Write-Host "Please enter the list you would like to generate items for:"
Write-Host "Ex: List Name"
$ListName = Read-Host "List"

#Get the amount of items to generate.
Write-Host "Please enter the amount of items you would like to generate:"
Write-Host "Ex: 1000"
[Int] $ItemCount = Read-Host "Amount"

#Get the name of all items to be generated.
Write-Host "Please enter the base name to be used for item generation:"
Write-Host "Ex: Test Item"
Write-Host "This will generate items such as Test Item 1, Test Item 2, Test Item 3, etc."
$ItemName = Read-Host "Name"

#Connect to the site via PnP.
try
    {
        Connect-PnPOnline -Url $SiteURL -UseWebLogin
    }

    catch
        {
            Write-Host -ForegroundColor Red "There was an error connecting to SharePoint via PnP PowerShell: $_"
        }

#Ask the user for what point they'd like to start the generation at:
Write-Host "What number would you like to start generating list items at? (Recommended: 1)"
[Int] $Counter = Read-Host "Starting Counter"

#Confirm operation prior to running:
Write-Host "Target Location:" ($SiteURL + $ListName)
Write-Host "Amount of items to be generated:" $ItemCount
Write-Host "Item Name:" $ItemName
Write-Host "Starting Count:" $Counter

Write-Host "To confirm this operation, please type the word > Accept < (Case Sensitive)."
$Validation = Read-Host "Confirmation"

if ($Validation -ceq "Accept")
    {
        #Loop to generate items.
        do
        {
            Add-PnPListItem -List $ListName -Values @{"Title" = "$ItemName $Counter"}
            $Counter++
        }

        #Argument determining how long the loop should run.
        while ($Counter -le $ItemCount)
    }

Write-Host -ForegroundColor Green "Item generation complete!"
Write-Host -ForegroundColor Green "Have a nice day!"
