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

	$group = Get-AzureADGroup -SearchString "O365-MigPilots"
	$User = Get-ADUser -Identity $id -Properties *
	$x = @()

	if($User.EmailAddress -eq $null)
	{
		Write-Host "Ingen mailadress skapad för $id.`nAvslutar"
	} else {
		Write-Verbose "Lägger till användare"
		try {
			Add-DistributionGroupMember -Identity $group.ObjectId -Member $User.EmailAddress -BypassSecurityGroupManagerCheck -ErrorAction SilentlyContinue
		} catch {
			Write-Host "Fel vid försök att lägga till medlemskap`nTroligen finns $User.Name redan i gruppen"
		}
	}
}
