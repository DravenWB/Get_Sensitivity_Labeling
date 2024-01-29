function Operator_Acknowledgement_Check
{
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
}
