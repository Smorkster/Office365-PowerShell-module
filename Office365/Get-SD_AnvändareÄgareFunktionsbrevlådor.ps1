<#
.Synopsis
	Hämta alla Funktionsbrevlådor en användare är markerad som ägare för
.Parameter id
	Användarens mailadress
.Example
	Get-SD_AnvändareÄgareFunktionsbrevlådor -id "ABCD"
#>

function Get-SD_AnvändareÄgareFunktionsbrevlådor
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id
	)

	$user = Get-ADUser -Identity $id -Properties *
	$funkar = Get-MailBox -Filter "CustomAttribute10 -like '*$user.Emailaddress*'"
	
	if($funkar.Count -gt 0)
	{
		Write-Host $MailAnvändare -NoNewline -Foreground Cyan
		Write-host " är ägare av"$funkar.Count"funktionsbrevlådor:"
		$funkar | ft DisplayName
	} else {
		Write-Host "Inga funktionsbrevlådor funna med"$Användare"som ägare"
	}
}
