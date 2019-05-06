<#
.Synopsis
	Sök Outlook-objekt efter angivet sökord
.Description
	Sök samtliga Outlook-objekt, baserat på objekttyp, vars mailadress innehåller ett angivet sökord
	När skriptet körs, måste objekttyp anges
.Parameter SökOrd
	Det sökord som ska finnas i objekts namn/mailadress
.Parameter Typ
	Vilken sorts objekt som hittas [Funktionsbrevlåda, Distributionslista, Rum, Resurs]
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Funktionsbrevlåda'
	Söker efter alla funktionsbrevlådor, vars mailadress eller namn innehåller 'test'
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Distributionslista'
	Söker efter alla distributionslistor, vars mailadress eller namn innehåller 'test'
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Rum'
	Söker efter alla rum, vars mailadress eller namn innehåller 'test'
.Example
	Search-SD_GemObjektMedOrdINamnet -SökOrd "test" -Typ 'Resurs'
	Söker efter alla resurser, vars mailadress eller namn innehåller 'test'
#>
function Search-SD_GemObjektMedOrdINamnet {
	param(
	[Parameter(Mandatory=$true)]
		[String] $SökOrd,
	[ValidateSet('Funktionsbrevlåda', 'Distributionslista', 'Rum', 'Resurs')]
	[Parameter(Mandatory=$true)]
		[String] $Typ
	)
	$SökOrd = "*"+$SökOrd+"*"

	if ($Typ -eq "Distributionslista")
	{
		Write-Host "Söker efter distributionslistor..." -Foreground Cyan
		$list = Get-DistributionGroup -Identity "$SökOrd" -ErrorAction SilentlyContinue
	} elseif ($Typ -eq "Funktionsbrevlåda") {
		Write-Host "Söker efter funktionsbrevlådor..." -Foreground Cyan
		$list = Get-MailBox -RecipientTypeDetails SharedMailbox -Identity "$SökOrd" -ErrorAction SilentlyContinue
	} elseif ($Typ -eq "Rum") {
		Write-Host "Söker efter rum..." -Foreground Cyan
		$list = Get-MailBox -RecipientTypeDetails RoomMailbox -Identity "$SökOrd" -ErrorAction SilentlyContinue
	} elseif ($Typ -eq "Resurs") {
		Write-Host "Söker efter resurser..." -Foreground Cyan
		$list = Get-Mailbox -RecipientTypeDetails EquipmentMailbox -Identity "$SökOrd" -ErrorAction SilentlyContinue
	} else {
		Write-Error "Felaktig söktyp angivet"
	}

	if($list.Count -eq 0)
	{
		Write-Host "Ingen $Typ hittad, med " -NoNewLine
		Write-Host $($SökOrd -replace "\*","") -Foreground Cyan -NoNewLine
		Write-Host " i namn eller mailadress"
	} else {
		$list | select DisplayName, PrimarySmtpAddress
	}
}
