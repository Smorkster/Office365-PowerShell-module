<#
.Synopsis
	Listar vilka användare som har behörighet till funktionsbrevlåda
.Parameter Funktionsbrevlåda
	Mailadress eller namn på funktionsbrevlådan
.Example
	Get-SD_FunkBehörigheter -Funktionsbrevlåda "test@test.com"
#>

function Get-SD_FunkBehörigheter
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Funktionsbrevlåda
	)

	Write-Verbose "Hämtar användare med full behörighet i Exchange"
	$fullExchange = Get-MailboxPermission -Identity $Funktionsbrevlåda.Trim() | ? {$_.AccessRights -like "*FullAccess*" -and $_.User -match "@test.com"}

	Write-Verbose "Hämtar användare med full behörighet i Azure"
	$AzureName = "MB-"+$Funktionsbrevlåda.Trim()+"-Full"
	$fullAzure = Get-AzureADGroupMember -ObjectId (Get-AzureADGroup -SearchString $AzureName).Objectid

	Write-Verbose "Hämtar användare med läsbehörighet i Exchange"
	$readExchange = Get-MailboxPermission $Funktionsbrevlåda.Trim() | ? {$_.AccessRights -like "ReadPermission" -and $_.User -match "@test.com"}

	Write-Verbose "Hämtar användare med läsbehörighet i Azure `n"
	$AzureName = "MB-"+$Funktionsbrevlåda.Trim()+"-Read"
	$readAzure = Get-AzureADGroupMember -ObjectId (Get-AzureADGroup -SearchString $AzureName).Objectid
	
	if ( $fullAzure.Count -eq 0 )
	{
		Write-Host "Inga användare har full behörighet i Azure" -Foreground Cyan
	} else {
		Write-Host "Dessa har full behörighet i Azure" -Foreground Cyan
	}
	$fullAzure | select -ExpandProperty UserPrincipalName

	Write "`n"

	if ( $fullExchange.Count -eq 0 )
	{
		Write-Host "Inga användare har full behörighet i Exchange" -Foreground Cyan
	} else {
		Write-Host "Dessa har full behörighet i Exchange" -Foreground Cyan
	}
	$fullExchange | select -ExpandProperty User

	Write "`n"

	if ( $readAzure.Count -eq 0 )
	{
		Write-Host "Inga användare har läsbehörighet i Azure" -Foreground Cyan
	} else {
		Write-Host "Dessa har läsbehörighet i Azure" -Foreground Cyan
	}
	$readAzure | select -ExpandProperty UserPrincipalName

	Write "`n"

	if ( $readExchange.Count -eq 0 )
	{
		Write-Host "Inga användare har läsbehörighet i Exchange" -Foreground Cyan
	} else {
		Write-Host "Dessa har läsbehörighet i Exchange" -Foreground Cyan
	}
	$readExchange | select -ExpandProperty User
}
