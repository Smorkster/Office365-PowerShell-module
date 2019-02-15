<#
.SYNOPSIS
	Verifiera ifall användare är medlem i en funktionsbrevlåda
.PARAMETER id
	Mailadress till användaren
.PARAMETER Funktionsbrevlåda
	Namn/mailadress för funktionsbrevlådan
.SYNTAX
	Confirm-SD_AnvändareMedlemIFunktionsbrevlåda -id <id> -Funktionsbrevlåda <Namn/mailadress>
.DESCRIPTION
	Söker funktionsbrevlåda och kontrollerar sedan om användaren finns med i medlemslistan
#>

function Confirm-SD_AnvändareMedlemIFunktionsbrevlåda
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id,
	[Parameter(Mandatory=$true)]
		[string] $Funktionsbrevlåda
	)

	$user = Get-ADUser -Identity $id -Properties *
	$member = Get-Mailbox -Identity $Funktionsbrevlåda | Get-MailboxPermission | ? {$_.User -eq $user.Emailaddress}

	Write-Host $MailAnvändare -ForegroundColor Cyan -NoNewline
	Write-Host " har " -NoNewline
	if($member)
	{
		Write-Host $member.AccessRights -NoNewline -ForegroundColor Cyan
	} else {
		Write-Host " ingen behörighet till " -nonewline
	}
	Write-Host $Funktionsbrevlåda -ForegroundColor Cyan
}
