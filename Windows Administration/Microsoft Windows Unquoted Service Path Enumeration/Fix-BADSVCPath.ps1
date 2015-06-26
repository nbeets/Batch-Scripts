#Fix-BADSVCPath.ps1
[cmdletbinding()]
	Param ( #Define a Mandatory input
	[Parameter(
	 ValueFromPipeline=$true,
	 ValueFromPipelinebyPropertyName=$true,
	 Position=0)] $obj
	) #End Param
 
Process
{ #Process Each object on Pipeline
	if ($obj.badkey -eq "Yes"){
		Write-Progress -Activity "Fixing $($obj.computername)\$($obj.key)" -Status "Working..."
		$regpath = $obj.Fixedkey
		$regpath = '"' + $regpath.replace('"', '\"') + '"' + ' /f'
		$obj.status = "Fixed"
		REG ADD "\\$($obj.computername)\$($obj.key)" /v ImagePath /t REG_EXPAND_SZ /d $regpath
		}
	Write-Output $obj
 
} #End Process