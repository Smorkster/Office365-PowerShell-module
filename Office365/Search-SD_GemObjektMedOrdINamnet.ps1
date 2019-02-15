<#
.Synopsis
	Sök Outlook-objekt efter angivet sökord
.Description
	Sök samtliga Outlook-objekt, baserat på objekttyp, vars mailadress innehåller ett angivet sökord
	När skriptet körs, måste objekttyp anges
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Funk'
	Söker efter alla funktionsbrevlådor, vars mailadress innehåller 'test'
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Dist'
	Söker efter alla distributionslistor, vars mailadress innehåller 'test'
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Rum'
	Söker efter alla rum, vars mailadress innehåller 'test'
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Res'
	Söker efter alla resurser, vars mailadress innehåller 'test'
#>
function Search-SD_GemObjektMedOrdINamnet {
	param(
	[Parameter(Mandatory=$true)]
		[String] $SökOrd,
	[ValidateSet('Funk', 'Dist', 'Rum', 'Res')]
	[Parameter(Mandatory=$true)]
		[String] $Typ
	)

	if($Typ -eq "Dist")
	{
		Write-Host "Söker efter distributionslistor..." -Foreground Cyan
		$list = Get-DistributionGroup -Identity "*Dist*" -Filter "PrimarySMTPAddress -like '*$SökOrd*'" -ErrorAction SilentlyContinue
	} elseif ($Typ -eq "Funk") {
		Write-Host "Söker efter funktionsbrevlådor..." -Foreground Cyan
		$SökOrd = "*"+$SökOrd+"*"
		$list = Get-MailBox -Identity "*Funk*" -Filter "PrimarySMTPAddress -like '$SökOrd'" -ErrorAction SilentlyContinue
	} elseif ($Typ -eq "Rum") {
		Write-Host "Söker efter rum..." -Foreground Cyan
		$list = Get-MailBox -Identity "*Rum*" -Filter "PrimarySMTPAddress -like '*$SökOrd*'" -ErrorAction SilentlyContinue
	} elseif ($Typ -eq "Res") {
		Write-Host "Söker efter resurser..." -Foreground Cyan
		$list = Get-MailBox -Identity "*Res*" -Filter "PrimarySMTPAddress -like '*$SökOrd*'" -ErrorAction SilentlyContinue
	}

	if($list.Count -eq 0)
	{
		Write-Host "Ingen $Typ hittad, som har någon mailadress innehållandes " -NoNewLine
		Write-Host $SökOrd -Foreground Cyan
	} else {
		$list
	}
}
