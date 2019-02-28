<#
.SYNOPSIS
	Anslut till tjänsterna
.Parameter reconnectExchange
	Återansluter till Exchange.
	Stänger nuvarande PSSession, skapar en ny och ansluter. Används när anslutningen till Exchange tappats vid inaktivitet
#>

function Connect-SD_ToServices
{
	param(
	 	[switch] $reconnectExchange
	)

	$ErrorActionPreference = "SilentlyContinue"
	$WarningPreference = "SilentlyContinue"
	#region Check PSSession
	if (!(Get-PSSession | Where {$_.ConfigurationName -eq "Microsoft.Exchange"}))
	{
		Write-Host "Du är inte ansluten till Exchange."

		Write-Host "Connecting to Exchange" -Foreground Cyan
		$EXOSession = New-ExoPSSession
		Import-PSSession $EXOSession -AllowClobber > $null
	} elseif ($reconnectExchange) {
		Get-PSSession | Remove-PSSession
		$EXOSession = New-ExoPSSession
		Import-PSSession $EXOSession -AllowClobber > $null
		Write-Host "Ansluten till Exchange" -ForegroundColor Green
	}
	#endregion Check PSSession

	#region Check MsolService
	Get-MsolDomain > $null
	if(-not $?)
	{
		Write-Host "Du är inte ansluten till MsolService."
		Write-Host "Ansluter." -Foreground Cyan
		Connect-MsolService
	}
	#endregion Check MsolService

	#region Check AzureAD
	Get-AzureADDomain > $null
	if(-not $?)
	{
		Write-Host "Du är inte ansluten till AzureAD."
		Write-Host "Ansluter." -Foreground Cyan
		Connect-AzureAD > $null
	}
	#endregion Check AzureAD

	$ErrorActionPreference = "Continue"
	$WarningPreference = "Continue"
	Write-Host "Connected to services"
}
