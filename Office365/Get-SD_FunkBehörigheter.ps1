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

	if ($Funktionsbrevlåda -match "@test.com")
	{
		$displayname = (Get-Mailbox -Identity $Funktionsbrevlåda).DisplayName
		$AzureNameFull = "MB-"+$displayname+"-Full"
		$AzureNameRead = "MB-"+$displayname+"-Read"
		$ExchangeName = $displayname
	} else {
		$AzureNameFull = "MB-"+$Funktionsbrevlåda.Trim()+"-Full"
		$AzureNameRead = "MB-"+$Funktionsbrevlåda.Trim()+"-Read"
		$ExchangeName = $Funktionsbrevlåda
	}

	Write-Verbose "Hämtar användare med full behörighet i Azure"
	$fullAzure = Get-AzureADGroupMember -ObjectId (Get-AzureADGroup -SearchString $AzureNameFull).Objectid

	Write-Verbose "Hämtar användare med läsbehörighet i Azure `n"
	$readAzure = Get-AzureADGroupMember -ObjectId (Get-AzureADGroup -SearchString $AzureNameRead).Objectid

	Write-Verbose "Hämtar användare med behörighet i Exchange"
	$ExchangeMembers = Get-MailboxPermission -Identity $ExchangeName | ? {$_.User -match "@test.com"}

	$fullExchange = $ExchangeMembers | ? {$_.AccessRights -like "*FullAccess*"}
	$readExchange = $ExchangeMembers | ? {$_.AccessRights -like "*ReadPermission*"}

	
	if ( $fullAzure.Count -eq 0 )
	{
		Write-Host "Inga användare har full behörighet i Azure" -Foreground Cyan
	} else {
		Write-Host "Dessa har full behörighet i Azure" -Foreground Cyan
		$fullAzure | sort UserPrincipalName | select -ExpandProperty UserPrincipalName
	}

	Write "`n"

	if ( $fullExchange.Count -eq 0 )
	{
		Write-Host "Inga användare har full behörighet i Exchange" -Foreground Cyan
	} else {
		Write-Host "Dessa har full behörighet i Exchange" -Foreground Cyan
	}
	$fullExchange | sort User | select -ExpandProperty User

	Write "`n"

	if ( $readAzure.Count -eq 0 )
	{
		Write-Host "Inga användare har läsbehörighet i Azure" -Foreground Cyan
	} else {
		Write-Host "Dessa har läsbehörighet i Azure" -Foreground Cyan
	}
	$readAzure | sort UserPrincipalName | select -ExpandProperty UserPrincipalName

	Write "`n"

	if ( $readExchange.Count -eq 0 )
	{
		Write-Host "Inga användare har läsbehörighet i Exchange" -Foreground Cyan
	} else {
		Write-Host "Dessa har läsbehörighet i Exchange" -Foreground Cyan
	}
	$readExchange | sort User | select -ExpandProperty User
}
