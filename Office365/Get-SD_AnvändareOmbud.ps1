<#
.SYNOPSIS
	Hämtar alla ombud för ett konto
.PARAMETER id_Ägare
	id för ägaren av Outlook-kontot
.Example
	Get-SD_Ombud -id_Ägare "ABCD"
#>

function Get-SD_AnvändareOmbud
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id_Ägare
	)
	
	try
	{
		$owner = (Get-ADUser -Identity $id_Ägare -Properties *).EmailAddress
		try
		{
			Write-Host "Behörigheter till Inkorgen:" -Foreground Green
			Get-MailboxFolderPermission -Identity $owner -ErrorAction Stop | ? { $_.User -notmatch "Standard" -and $_.User -notmatch "Anonymous" } | % { Write-Host $_.User " -> " $_.AccessRights }
			try
			{
			Write-Host "`nBehörigheter till Kalender:" -Foreground Green
			Get-MailboxFolderPermission -Identity $owner":\Kalender" -ErrorAction Stop | ? { $_.User -notmatch "Standard" -and $_.User -notmatch "Anonymous" } | % { Write-Host $_.User " -> " $_.AccessRights }
			} catch {
				if ( $Error[0].CategoryInfo -like "*ManagementObjectNotFoundException*" )
				{
					Write-Host "Du saknar behörighet för att se mappar/kalender`nLägg till dig själv som administratör i Exchange först."
				}
			}
		} catch {
			Write-Host "Mailkonto inte funnet" -Foreground Red
		}
	} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
		Write-Host "Person med id "$id_Ägare" inte funnen`nAvslutar" -Foreground Red
	}
}
