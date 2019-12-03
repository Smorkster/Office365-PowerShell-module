<#
.Synopsis
	Är användare medlem i en Azure-grupp
.Description
	Hämtar användaren från Azure och kontrollerar om grupp finns med i listan över medlemskap kopplat till användaren
.Parameter MailAnvändare
	Användarens mailadress
.Parameter GruppNamn
	Namn på gruppen i fråga
.Example
	Confirm-SD_AnvändareMedlemIAzureGrupp -MailAnvändare <Mailadress> -GruppNamn <Namn på Azure-grupp>
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
		$user = Get-AzureADUser -ObjectId $MailAnvändare
	} catch {
		Write-Host "Användare hittades inte i Azure.`nAvslutar"
		return
	}

	if ( ( $group = Get-AzureADGroup -SearchString $GruppNamn ).Count -eq 0)
	{
		Write-Host "Grupp " -NoNewline
		Write-Host $GruppNamn -Foreground Cyan -NoNewline
		Write-Host " hittades inte i Azure.`nAvslutar"
		return
	}

	$groups = @()
	try {
		$groups = $user | Get-AzureADUserMembership -All $true | ? {$_.DisplayName -like $GruppNamn}
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
