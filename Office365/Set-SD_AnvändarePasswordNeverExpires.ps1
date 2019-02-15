<#
.SYNOPSIS
	Anger att lÃ¶senordet fÃ¶r kontot inte gÃ¥r ut
.PARAMETER id
	id fÃ¶r kontot
.DESCRIPTION
	Skriptet anvÃ¤nds fÃ¶r att sÃ¤tta att lÃ¶senordet inte gÃ¥r ut.
.Example
	Set-SD_AnvÃ¤ndarePasswordNeverExpires -id "ABCD"
#>

function Set-SD_AnvÃ¤ndarePasswordNeverExpires
{
	param(
	[Parameter(Mandatory=$true)]
		[String] $id
	)

	try {
		$obj = Get-Mailbox -Identity (Get-ADUser -Identity $id -Properties *).EmailAddress
	} catch {
		Write-Host "Ingen anvÃ¤ndare hittades"
	}
	Set-MsolUser -UserPrincipalName ($obj).UserPrincipalName -PasswordNeverExpires $true
}
