<#
.SYNOPSIS
	Anslut till tjänsterna
#>

function Connect-SD_ToServices
{
	$ErrorActionPreference = "SilentlyContinue"
	$WarningPreference = "SilentlyContinue"
	#region Check PSSession
	if($c -eq 1)
	{
		if (!(Get-PSSession | Where {$_.ConfigurationName -eq "Microsoft.Exchange"}))
		{
			Write-Host "Du är inte ansluten till Exchange."

			Write-Host "Connecting to Exchange" -Foreground Cyan
			$EXOSession = New-ExoPSSession
			Import-PSSession $EXOSession -AllowClobber > $null
		}
	}
	#endregion Check PSSession

	#region Check MsolService
	if($c -eq 2)
	{
		Get-MsolDomain > $null
		if(-not $?)
		{
			Write-Host "Du är inte ansluten till MsolService."
			Write-Host "Ansluter." -Foreground Cyan
			Connect-MsolService
		}
	}
	#endregion Check MsolService

	#region Check AzureAD
	if($c -eq 3)
	{
		Get-AzureADDomain > $null
		if(-not $?)
		{
			Write-Host "Du är inte ansluten till AzureAD."
			Write-Host "Ansluter." -Foreground Cyan
			Connect-AzureAD > $null
		}
	}
	#endregion Check AzureAD

	$ErrorActionPreference = "Continue"
	$WarningPreference = "Continue"
	Write-Host "Connected to services"
}
