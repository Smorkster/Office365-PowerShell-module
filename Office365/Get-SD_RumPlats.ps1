<#
.SYNOPSIS
	Hämtar angiven plats för ett rum, om det är angivet
.PARAMETER Rumsnamn
	Namn på rummet
.Example
	Get-SD_RumPlats -Rumsnamn "RumA"
#>

function Get-SD_RumPlats
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Rumsnamn
	)

	try {
		$a = (Get-Mailbox -Identity $Rumsnamn).Office
		if($a -eq $null -or $a -eq "")
		{
			Write-Host "Ingen plats är specificerad"
		} else {
			Write-Host "`n"$a"`n"
		}
	} catch {
		Write-Host "`nRum " -nonewline
		Write-Host $Rumsnamn -Foreground Red -NoNewline
		Write-Host " finns inte"
	}
}
