<#
.SYNOPSIS
	Hämtar ägare av distributionslista
.PARAMETER Distributionslista
	Namn eller mailadress för distributionslistan
.Example
	Get-SD_DistÄgare -Distributionslista "Distlista"
#>

function Get-SD_DistÄgare
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Distributionslista
	)

	try {
		( Get-DistributionGroup -Identity $Distributionslista -ErrorAction Stop ).ManagedBy | ? {$_ -notlike "*MIG-User-1-Farm-1*"}
	} catch {
		Write-Host "`nIngen distributionslista med namnet " -nonewline
		Write-Host $Distributionslista -ForegroundColor Red -nonewline
		Write-Host " funnen"
	}
}
