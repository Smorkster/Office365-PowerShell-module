<#
.SYNOPSIS
	Skapa en ny rumslista
.PARAMATER NamnPåRumslista
	Namnet för rumslistan. Ska följa namnstandarden "<Organisation> <Namn>"
.PARAMATER PrimarySMTPAddress
	Primär SMTP-adress för distributionslistan. Ska följa namnstandarden för distributionslistor
.PARAMATER Smtpadress
	Sekundär SMTP-adress
.SYNTAX
	New-DistributionGroup -NamnPåRumslista <Namn> -PrimarySMTPAddress <Mailadress> -Smtpadress <Mailadress>
#>
function New-SD_Rumslista
{
	param(
	[Parameter(Mandatory=$true)]
		[String] $NamnPåRumslista,
	[Parameter(Mandatory=$true)]
		[String] $PrimarySMTPAddress,
	[Parameter(Mandatory=$true)]
		[String] $Smtpadress
	)
	
	New-DistributionGroup -Name $NamnPåRumslista -Roomlist
	Set-DistributionGroup -Identity $NamnPåRumslista -PrimarySMTPAddress $PrimarySMTPAddress
	Set-DistributionGroup -Identity $NamnPåRumslista -EmailAddresses @{Add="smtp:$Smtpadress"}
	Set-DistributionGroup -Identity $NamnPåRumslista -Description "Now"
}
