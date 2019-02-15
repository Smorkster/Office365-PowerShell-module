<#
.SYNOPSIS
	Skapa behörighet för användaren att boka rum
.PARAMETER id
	id för den användare som ska ha behörigheten
.PARAMETER Rum
	Namn eller identitet på rummet som behörigheten ska skapas på
.DESCRIPTION
	Lägg in bokningsbehörighet för en användare till ett rum. Skriptet används när synken från ändring i Supportpanelen inte har gått över till Exchange.
.Example
	Add-SD_RumBokaEnAnvändare -id "ABCD" -Rum "RumA"
#>

function Add-SD_RumBokaEnAnvändare
{
	param(
	[Parameter(Mandatory=$true)]
		[String] $id,
	[Parameter(Mandatory=$true)]
		[String] $Rum
	)

	try {
		$RoomObject = Get-Mailbox -Identity $Rum
	} catch [System.Management.Automation.RemoteException] {
		Write-Host "Rum hittades inte"
		return
	}

	try {
		$User = (Get-ADUser -Identity $id -Properties *).Mailaddress
		if($User -eq $null) {
			Write-Host "Ingen mailadress registrerad i AD för användaren"
		}
	} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
		Write-Host "Användare hittades inte i AD"
		return
	}

	try {
		$UserAccount = Get-Mailbox -Identity $User
		$BookPolicy = (Get-CalendarProcessing -Identity $RoomObject).BookInPolicy += $User

		Set-CalendarProcessing -Identity $RoomObject -BookInPolicy $BookPolicy
	} catch [System.Management.Automation.RemoteException] {
		Write-Host "Rum $Rum hittades inte i Exchange.`nAvslutar"
	}
}
