##############################################################################
# - Script Definition: This script is designed to gather a comprehensive list
# of SharePoint site owners and save them to a .csv file for review. This
# list includes all sites within a tenant.
#
# - Development Date: October 08, 2023
#
# - Limitations:
#     + This script will only gather site owner data for default configured groups.
#     + Custom groups with owners permissions are not yet supported.
#     + Service Limitation: You must be the site collection administrator for all
#       sites you intend on gathering information for.
#     + Some sites may return "Group cannot be found" for sites with custom
#       group names or that which do not have an owners group.
#     + Some sites may return "Access Denied" for those that you do not have
#       permissions to or system level sites.
#     + Owners groups that contain an active directory or security group to apply
#       permissions will be currently listed as the group Object GUID and not the
#       name of the group. This may change in future iterations.
#     + Setting a file name for saving data may write over existing files with the
#       same name. Please check this prior to running. The default save location is
#       the desktop.
#
# - Disclaimer: This script is not officially supported by Microsoft, its
# affiliates or partners. This script is provided as is and the responsibility
# of understanding the scripts functions and operations falls upon those that
# may choose to run it. Responsibility of positive or negative outcomes of
# this script may not receive future assistance as such.
#
# - To accept this risk and arm this script for use, please un-comment the
# Connect-SPOService comand. This is in place to prevent damage via use by
# those who do not understand this scripts contents and functions.
##############################################################################

#Variables to be configured prior to the running of the script.

#Get the administrator email for connecting.
Write-Host "Please enter your administrator email to ensure permissions are available for data check."
$AdminUPN = Read-Host "AdminUPN"

#Get the filename the operator would like to set the output to.
Write-Host "Please enter the name of the file you would like to store the data in. Ex: Site Owners"
$FileName = Read-Host "Save File Name"

#Get the tenant name for connecting
Write-Host "Please enter the domain name of your tenant. Ex: contoso"
$TenantName = Read-Host "Tenant Domain Name"

#Get the tenant type for connecting.
Write-Host "What kind of tenant are you connecting to?"
Write-Host "1. Commercial (.com)"
Write-Host "2. GCC-High (.us)"

$Menu = Read-Host #Set variable for endpoint selection.

#Convert user selection variable to the correct string.
if ($Menu -eq 1)
    {
        $Endpoint = "com"
    }

    else
    {
        $Endpoint = "us"
    }

Write-Host -ForegroundColor Green "Now connecting to the SPO Service..." #Operator process information.

#Configure tenant admin center URL from user input.
$SPOAdminCenterURL = "https://$TenantName-admin.sharepoint.$Endpoint"

#Connect to the SPO service and try geo-location locks if required.
try
{
#    Connect-SPOService -Url $SPOAdminCenterURL

    Write-Host -ForegroundColor Green "Successfully connected to the SPO Service!" #Operator process information.
}

    catch
    {
        if($Endpoint -eq ".us")
        {
            try
            {
            Connect-SPOService -Url $SPOAdminCenterURL -Region ITAR

            Write-Host -ForegroundColor Green "Successfully connected to the SPO Service under ITAR restrictions!" #Operator process information.
            }

                catch
                {
                   Write-Host -ForegroundColor Red "Failed to connect to the GCC-High SPO Service with the following error:" $_
                   break; 
                }
        }

        else
        {
            Write-Host -ForegroundColor Red "Failed to connect to the commercial SPO Service with the following error:" $_
            break;
        }
    }

Write-Host -ForegroundColor Green "Now gathering tenant site data" #Operator process information.

#Gather all sites within the tenant.
$SiteIndex = Get-SPOSite -limit ALL

#Generate object class to store .csv data.
class TableData
{
    [string] ${SiteURL}
    [string] ${User}
}

#Generate array to store data.
$SiteIndexData= @()

#Loop to navigate each site and gather ALL owners assigned to each site.
foreach ($Site in $SiteIndex)
{
    #Set Site title variable for default groups.
    $SiteTitle = $Site.Title

    Write-Host "_______________________________________________"
    Write-Host "Now processing Site: $SiteTitle"
    Write-Host "_______________________________________________"

    #Get owners and add them to the site data index.
    $SiteOwners = Get-SPOUser -Site $Site.Url -Group "$SiteTitle Owners" #Gets all owners on the individual site.

    foreach($Item in $SiteOwners) #Loops through each owner on the site and adds the object
    {
        $Object = New-Object PSObject -Property @{
        SiteURL = $Site.Url
        User = $Item.LoginName
        }

        $SiteIndexData += $Object
    }

}

#Export site owner data to a file.
$SiteIndexData | Export-Csv ~/Desktop/$FileName.csv -Encoding utf8
Write-Host -ForegroundColor Green "Successfully saved $FileName.csv to the Desktop!"
