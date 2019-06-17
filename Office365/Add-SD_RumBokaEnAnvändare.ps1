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
	Skapar behörighet för användare ABCD att skapa bokningar i rum RumA
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
		$User = Get-ADUser -Identity $id -Properties *
		if($User.Emailaddress -eq $null) {
			Write-Host "Ingen mailadress registrerad i AD för användaren.`nAvslutar"
			return
		}
	} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
		Write-Host "Användare hittades inte i AD"
		return
	}

	try {
		$UserAccount = Get-Mailbox -Identity $User.Emailaddress
		$BookPolicy = (Get-CalendarProcessing -Identity $RoomObject.Identity).BookInPolicy += $User.Emailaddress | select -Unique

		Set-CalendarProcessing -Identity $RoomObject.Identity -BookInPolicy $BookPolicy -AllBookInPolicy:$false
		
		Write-Host "$User.Name har nu behörighet att boka $RoomObject.DisplayName"
	} catch [System.Management.Automation.RemoteException] {
		Write-Host "Rum $Rum hittades inte i Exchange.`nAvslutar"
	}
}
