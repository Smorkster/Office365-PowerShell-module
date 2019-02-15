<#
.SYNOPSIS
	Lista alla medlemmar i en distributionslista
.Description
	HÃ¤mtar samtliga medlemmar av en distributionslista och listar dem sorterat i bokstavsordning efter namn
.PARAMETER Distributionslista
	Namn pÃ¥ distributionslistan
.Example
	Get-SD_DistMedlemmar -Distributionslista "Distlista"
#>

function Get-SD_DistMedlemmar
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Distributionslista
	)

	$list = @()
	try {
		Get-DistributionGroupMember -Identity $Distributionslista.Trim() -ErrorAction Stop | sort PrimarySMTPAddress | ft PrimarySMTPAddress, DisplayName
	} catch {
		Write-Host "`nDistributionslistan " -NoNewline
		Write-Host $Distributionslista -Foreground Red -NoNewline
		Write-Host " finns inte"
	}
}
