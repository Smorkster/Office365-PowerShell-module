<#
.SYNOPSIS
	Hämtar alla ombud för ett konto
.PARAMETER id_Ägare
	id för ägaren av Outlook-kontot
.Example
	Get-SD_Ombud -id_Ägare "ABCD"
	Hämtar alla ombud som skapats till mailkontot för användaren ABCD
#>

function Get-SD_AnvändareOmbud
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id_Ägare
	)
	
	try
	{
		$owner = Get-ADUser -Identity $id_Ägare -Properties *

		Write-Host "Behörigheter till Inkorgen:" -Foreground Green
		Get-MailboxFolderPermission -Identity $owner.EmailAddress -ErrorAction Stop | ? { $_.User -notmatch "Standard" -and $_.User -notmatch "Anonymous" } | % { Write-Host $_.User " -> " $_.AccessRights }

		Write-Host "`nBehörigheter till Kalender:" -Foreground Green
		try
		{
			Get-MailboxFolderPermission -Identity $owner":\Kalender" -ErrorAction Stop | ? { $_.User -notmatch "Standard" -and $_.User -notmatch "Anonymous" } | % { Write-Host $_.User " -> " $_.AccessRights }
		} catch {
			if ( $_.CategoryInfo -like "*ManagementObjectNotFoundException*" )
			{
				try
				{
					Get-MailboxFolderPermission -Identity $owner":\Calendar" -ErrorAction Stop | ? { $_.User -notmatch "Standard" -and $_.User -notmatch "Anonymous" } | % { Write-Host $_.User " -> " $_.AccessRights }
				} catch {
					Write-Host "Kunde inte hitta kalendern"
				}
			} else {
				Write-Host "Kunde inte hitta kalendern"
			}
		}
	} catch {
		if ($_.CategoryInfo.Reason -eq "ADIdentityNotFoundException")
		{
			Write-Host "Person med id "$id_Ägare" inte funnen i AD.`nAvslutar" -Foreground Red
			return
		} elseif ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException" -and $_.CategoryInfo.Activity -eq "Get-MailboxFolderPermission") {
			Write-Host "Mailkonto för $($owner.Name) inte funnet" -Foreground Red
		} else {
			Write-Host "Fel uppstod i körningen:"
			$_
		}
	}
}
