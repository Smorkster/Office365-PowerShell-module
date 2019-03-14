<#
.SYNOPSIS
	Verifiera ifall en användare är medlem i en Azure-grupp
.PARAMETER MailAnvändare
	Användarens mailadress
.PARAMETER GruppNamn
	Namn på gruppen i fråga
.SYNTAX
	Confirm-SD_AnvändareMedlemIAzureGrupp -MailAnvändare <Mailadress> -GruppNamn <Namn på Azure-grupp>
.DESCRIPTION
	Söker användaren och kontrollerar om GruppNamn finns med i listan över medlemskap kopplat till användaren
#>

function Confirm-SD_AnvändareMedlemIAzureGrupp
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $MailAnvändare,
	[Parameter(Mandatory=$true)]
		[string] $GruppNamn
	)

	try {
		$user = Get-AzureADUser -SearchString $MailAnvändare
	} catch {
		Write-Host "Anvädnare hittades inte i Azure.`nAvslutar"
		return
	}

	try {
		$group = Get-AzureADGroup -SearchString $GruppNamn
	} catch {
		Write-Host "Grupp $GruppNamn hittades inte i Azure.`nAvslutar"
		return
	}

	$groups = @()
	try {
		$groups = $user | Get-AzureADUserMembership | ? {$_.DisplayName -like $GruppNamn}
		Write-Host $user.DisplayName -NoNewline -Foreground Cyan
		if ($groups)
		{
			Write-Host " är medlem i " -NoNewline
		} else {
			Write-Host " är inte medlem i " -NoNewline
		}
		Write-Host "'$GruppNamn'" -Foreground Cyan
	} catch {
		Write-Host "Ett fel uppstod"
	}
}
