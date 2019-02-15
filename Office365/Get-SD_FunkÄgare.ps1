<#
.SYNOPSIS
	Hämta ägare av funktionsbrevlåda
.PARAMETER Funktionsbrevlåda
	Namn eller mailadress för funktionsbrevlådan
.Example
	Get-SD_FunkÄgare -Funktionsbrevlåda "Funklåda"
#>

function Get-SD_FunkÄgare
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Funktionsbrevlåda
	)

	try {
		$owner = ( Get-Mailbox -Identity $Funktionsbrevlåda.Trim() -ErrorAction Stop ).CustomAttribute10
		Get-Mailbox -Identity $owner.Substring(7) | select DisplayName, UserPrincipalName
	} catch {
		Write-Host "Ingen funktionsbrevlåda med namnet " -nonewline
		Write-Host $Funktionsbrevlåda -ForegroundColor Magenta -nonewline
		Write-Host " funnen"
	}
}
