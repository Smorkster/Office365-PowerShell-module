<#
.Synopsis
    Synka behörigheter för användare till rum. Från Azure till Exchange
.Description
    Ifall en användare har fått behörighet skapad i en Azure-grupp för rum, men detta inte har blivit översynkat till Exchange, får vi lägga på behörigheten manuellt
    Parameter Rumslista kan ta ett eller flera rum och loopar då igenom varje rum enskilt
.Parameter id
    id för användare som ska synkroniseras
.Parameter Rum
    Ett eller flera rum som behörigheten ska läggas på
#>
function Sync-SD_RumAzureTillExchange
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id,
	[Parameter(Mandatory=$true)]
		[string[]] $Rum
	)

	$user = Get-Mailbox -Identity (Get-ADUser -Identity $id -Properties *).Emailaddress
	$ticker = 1

	foreach ($r in $Rum) {
		$eRum = "$r`:\Kalender"

		$bookpolicy = (Get-CalendarProcessing -Identity $r).BookInPolicy
		$bookpolicy += $user.LegacyExchangeDN
        Write-Progress -Activity $r -PercentComplete (($ticker/$Rum.Count)*100)
		Set-CalendarProcessing -Identity $r -AllBookInPolicy:$false -BookInPolicy $bookpolicy -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
		try {
			Add-MailboxFolderPermission -Identity $eRum -AccessRights LimitedDetails -Confirm:$false -User $user.PrimarySMTPAddress -ErrorAction Stop
		} catch {
			if ($_.CategoryInfo.Reason -like "*UserAlreadyExist*")
			{
				Write-Host "Behörighet finns redan"
			} elseif ($_.CategoryInfo.Reason -eq "ACLTooBigException") {
				Write-Host "För många medlemmar i Azure-gruppen. Kan inte synkronisera $($user.DisplayName) till Exchange.`nAvslutar."
				return
			} else {
				Write-Host "Problem vid skapande av behörighet i kalendern:"
				Write-Host $_.CategoryInfo.Reason
				Write-Host $_.Exception
			}
		}
		$ticker++
	}
}
