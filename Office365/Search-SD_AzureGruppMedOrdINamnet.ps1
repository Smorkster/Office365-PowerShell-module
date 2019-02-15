<#
.SYNOPSIS
	Sök på Office365-grupp genom sökningsord
.PARAMETER GruppNamn
	Ord som kan finnas med i gruppnamnet
.Example
	Search-SD_AzureGruppMedOrdINamnet -GruppNamn "Group1"
#>

function Search-SD_AzureGruppMedOrdINamnet
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $GruppNamn
	)

	try {
		$a = Get-MsolGroup -All | ? {$_.DisplayName -match $GruppNamn} | sort DisplayName
		if($a -eq $null)
		{
			Write-Host "`nIngen grupp med namn " -nonewline
			Write-Host $GruppNamn -ForegroundColor Red -nonewline
			Write-Host " finns i Exchange"
		} else {
			$a
		}
	} catch {
		Write-Host "Fel vid sökning på namn"
		Write-Host $GruppNamn -ForegroundColor Red
	}
}
