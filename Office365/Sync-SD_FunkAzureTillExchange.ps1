<#
.Synopsis
    Synka behörigheter för användare till funktionsbrevlåda, från Azure till Exchange
.Description
    Ifall användare har fått behörighet skapad i en Azure-grupp för funktionsbrevlåda, men detta inte har blivit översynkat till Exchange, lägger skriptet på behörigheterna manuellt.
.Parameter Funktionsbrevlåda
    Namn på funktionsbrevlådan
#>
function Sync-SD_FunkAzureTillExchange
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Funktionsbrevlåda
	)

	$azureGroupNameFull = "MB-"+$Funktionsbrevlåda.Trim()+"-Full"
	$azureGroupNameRead = "MB-"+$Funktionsbrevlåda.Trim()+"-Read"
	$ticker = 1

	try {
		$azureGroupFull = Get-AzureADGroup -SearchString $azureGroupNameFull -ErrorAction Stop
		$azureGroupRead = Get-AzureADGroup -SearchString $azureGroupNameRead -ErrorAction Stop
	} catch {
		Write-Host "Ingen grupp funnen i Azure för namn $($Funktionsbrevlåda.Trim()).`nAvslutar"
	}
	try {
		$exchange = Get-Mailbox -Identity $($Funktionsbrevlåda.Trim()) -ErrorAction Stop
	} catch {
		Write-Host "Ingen funktionsbrevlåda med namn $($Funktionsbrevlåda.Trim()) hittades i Exchange.`nAvslutar"
	}

	#region Sync Full
	$members = Get-AzureADGroupMember -ObjectId $azureGroupFull.ObjectId -All $true
	foreach($member in $members)
	{
		Write-Progress -Activity "Lägger till full behörighet för $($member.UserPrincipalName)" -PercentComplete (($ticker / $members.Count)*100)
		try {
			Add-MailboxPermission -Identity $exchange.Identity -User $member.UserPrincipalName -AccessRights FullAccess -AutoMapping:$true -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
		} catch {
			$_
		}
		$ticker++
	}
	#endregion

	#region Sync Read
	$ticker = 1
	$members = Get-AzureADGroupMember -ObjectId $azureGroupRead.ObjectId
	foreach($member in $members)
	{
		Write-Progress -Activity "Lägger till läsbehörighet för $($member.UserPrincipalName)" -PercentComplete (($ticker / $members.Count)*100)
		try {
			Add-MailboxPermission -Identity $exchange.Identity -User $member.UserPrincipalName -AccessRights ReadPermission -Confirm:$false -ErrorAction Stop -WarningAction SilentlyContinue | Out-Null
		} catch {
			$_
		}
		$ticker++
	}
	#endregion
}
