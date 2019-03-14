<#
.SYNOPSIS
	Kontrollera ifall msExchMailboxGuid är tomt
.PARAMETER id
	id för användaren
.SYNTAX
	Confirm-SD_AnvändaremsExchMailboxGuid -id <id>
.DESCRIPTION
	Kontrollera värdet på msExchMailboxGuid hos användaren.
	Detta används då det inte skapas någon låda för användaren, trots synk
#>

function Confirm-SD_AnvändaremsExchMailboxGuid
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id
	)

	try {
		$value = (Get-ADUser -Identity $id -Properties * -ErrorAction Stop).msExchMailboxGuid
		if($value -eq $null)
		{
			Write-Host "msExchMailboxGuid är tom"
		} else {
			Write-Host "msExchMailboxGuid är inte tom"
		}
	} catch {
		Write-Host "Ingen användare för $id hittades i AD"
	}
}
