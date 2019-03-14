<#
.SYNOPSIS
	Hämta lista över administratörer för en funktionsbrevlåda
.PARAMETER Funktionsbrevlåda
	Namn/mailadress för funktionsbrevlådan
.SYNTAX
	Get-SD_FunkAdmins -Funktionsbrevlåda <Namn/mailadress>
#>

function Get-SD_FunkAdmins
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Funktionsbrevlåda
	)

	$funk = Get-Mailbox -Identity $Funktionsbrevlåda.Trim() -ErrorAction SilentlyContinue
	if($funk -eq $null)
	{
		Write-Host "Funktionsbrevlåda " -NoNewline
		Write-Host $Funktionsbrevlåda -ForegroundColor Magenta -NoNewline
		Write-Host " finns inte"
	} else {
		Get-MailboxPermission -Identity $funk.Identity -ErrorAction Stop | ? {($_.AccessRights -match "FullAccess") -and ($_.User -match "@test.com")} | ft User
	}
}
