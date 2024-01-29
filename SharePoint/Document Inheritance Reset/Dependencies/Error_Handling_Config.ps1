function Error_Action_Initialization
{
	#Get operator current error output length and set to concise.
	$OriginalErrorView = $ErrorView
	$ErrorView = ConciseView

	#Get operator current error action and set to stop.
	$OriginalErrorAction = $ErrorActionPreference
	$ErrorActionPreference = Continue
}

function Error_Action_Cleanup
{
	#Reset error view/length to operator original setting.
	$ErrorView = $OriginalErrorView
	$ErrorActionPreference = $OriginalErrorAction
}
