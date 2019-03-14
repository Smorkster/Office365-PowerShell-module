<#
.SYNOPSIS
	Lägger in en användare i gruppen O365-MigPilots
.PARAMETER id
	id för användaren
.DESCRIPTION
	Ifall en användare inte har blivit inlagd i gruppen O365-MigPilots kan det bli problem med att t.ex. logga in.
	Använd då detta skript för att lägga till användaren.
.Example Add-SD_AnvändareTillO365MigPilot -id "ABCD"
#>

function Add-SD_AnvändareTillO365MigPilot
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id
	)

	try {
		$group = Get-AzureADGroup -SearchString "O365-MigPilots" -ErrorAction Stop
		$User = Get-ADUser -Identity $id -Properties * -ErrorAction Stop

		if($User.EmailAddress -eq $null)
		{
			Write-Host "Ingen mailadress skapad för $id.`nAvslutar"
		} else {
			Write-Verbose "Lägger till användare"
			Add-DistributionGroupMember -Identity $group.ObjectId -Member $User.EmailAddress -BypassSecurityGroupManagerCheck -ErrorAction Stop
		}
	} catch {
		if ($_.CategoryInfo.Reason -eq "ADIdentityNotFoundException")
		{
			Write-Host "Användare hittades inte i AD"
		} elseif ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException") {
			Write-Host "Användare hittades inte i Azure"
		} else {
			Write-Host "Fel uppstod i körningen:"
			$_
		}
	}
}
