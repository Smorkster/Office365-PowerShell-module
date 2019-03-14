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

	try {
		if ($Mailadress -eq $null)
		{
			$adUser = Get-ADUser -Identity $id -Properties * -ErrorAction Stop
			$Mailadress = Get-Mailbox -Identity $adUser.EmailAddress -ErrorAction Stop
		}
		Set-MsolUser -UserPrincipalName $Mailadress -PasswordNeverExpires $true -ErrorAction Stop
	} catch {
		if ($_.CategoryInfo.Reason -eq "ADIdentityNotFoundException")
		{
			Write-Host "Hittades inte i AD"
		} elseif ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") {
			Write-Host "Användare hittades inte i Exchange"
		} elseif ($_.FullyQualifiedErrorId -like "*UserNotFoundException*") {
			Write-Host "Användare hittades inte i Azure"
		} else {
			Write-Host "Fel uppstod i körningen:"
			$_
		}
	}
}
