<#
.SYNOPSIS
	Uppdaterar användares behörighet till rum
.PARAMETER Rum
	Namn eller mailadress för rummet
.Parameter id
	id för användare som behöver få behörighet uppdaterad
.Example
	Update-SD_AnvändareBehörighetTillRum -Rum "Rum" -id "ABCD"
#>

function Update-SD_AnvändareBehörighetTillRum
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Rum,
	[Parameter(Mandatory=$true)]
		[string] $id
	)

	Write-Verbose "Hämtar mailbox i Exchange"
	$user = Get-Mailbox -Identity (Get-ADUser -Identity $id -Properties *).EmailAddress
	try {
		Write-Verbose "Hämtar mailbox för rummet"
		$room = Get-Mailbox	-Identity $Rum
		if($room -eq $null)
		{
			Write-Host "Hittade inte rummet"
			return
		} else {
			Write-Verbose "Kontrollerar om det redan finns behörighet för användaren i rummet"
			$roomMember = Get-MailboxFolderPermission -Identity $rum":\Kalender" | ? {$_.User -match $user.DisplayName}
			if($roomMember)
			{
				Write-Verbose "Behörighet fanns redan, tar bort behörigheten"
				Remove-MailboxFolderPermission -Identity $rum":\Kalender" -User $user.PrimarySmtpAddress -Confirm:$false
			}
			Write-Verbose "Lägger på behörigheten för användaren till rummet"
			Add-MailboxFolderPermission -Identity $rum":\Kalender" -User $user.PrimarySmtpAddress -AccessRights LimitedDetails > $null
		}
	} catch {
		Write-Host "Ingen funktionsbrevlåda med namnet " -nonewline
		Write-Host $Funktionsbrevlåda -ForegroundColor Magenta -nonewline
		Write-Host " funnen"
	}
}
