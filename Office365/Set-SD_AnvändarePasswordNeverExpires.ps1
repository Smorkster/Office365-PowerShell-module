<#
.SYNOPSIS
	Anger att lösenordet för kontot inte går ut
.PARAMETER id
	id för kontot
.DESCRIPTION
	Skriptet används för att sätta att lösenordet inte går ut.
.Example
	Set-SD_AnvändarePasswordNeverExpires -id "ABCD"
#>

function Set-SD_AnvändarePasswordNeverExpires
{
	param(
		[string] $id,
		[string] $Mailadress
	)

	if ($Mailadress -eq $null)
	{
		try {
			$Mailadress = Get-Mailbox -Identity (Get-ADUser -Identity $id -Properties *).EmailAddress
		} catch {
			Write-Host "Ingen användare med id $id hittades`nAvslutar"
			return
		}
	}
	Set-MsolUser -UserPrincipalName $Mailadress -PasswordNeverExpires $true
}
