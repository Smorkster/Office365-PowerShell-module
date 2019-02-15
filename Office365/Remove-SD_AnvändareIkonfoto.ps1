<#
.SYNOPSIS
	Tar bort en användares profilbild i Office365
.PARAMETER id
	id för användaren
.Example
	Remove-SD_AnvändareIkonfoto -id "ABCD"
#>

function Remove-SD_AnvändareIkonfoto
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id
	)
	try {
		$user = Get-Mailbox -Identity (Get-ADUser -Identity $id -Properties *).EmailAddress -ErrorAction Stop
	} catch {
		Write-Host "Ingen användare hittades"
	}
	if($user)
	{
		Remove-UserPhoto -Identity $id -Confirm
	} else {
		Write-Host "Inget konto hittat för " -NoNewline
		Write-Host $id
	}
}
