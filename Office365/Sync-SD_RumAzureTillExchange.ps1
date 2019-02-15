<#
.Synopsis
    Synka behörigheter för användare till rum. Från Azure till Exchange
.Description
    Ifall en användare har fått behörighet skapad i en Azure-grupp för rum, men detta inte har blivit översynkat till Exchange, får vi lägga på behörigheten manuellt
    Parameter Rumslista kan ta ett eller flera rum och loopar då igenom varje rum enskilt
.Parameter id
    id för användare som ska synkroniseras
.Parameter Rum
    
#>
function Sync-SD_RumAzureTillExchange
{
	param=(
	[Parameter(Mandatory=$true)]
		[string] $id,
	[Parameter(Mandatory=$true)]
		[string[]] $Rum
	)

	#$rumlista = @("SLSO Funk Bokning Torsplan plan4";"SLSO Funk Internservice";"SLSO Funk Receptionen Torsplan")
	$user = Get-Mailbox -Identity (Get-ADUser -Identity $id -Properties *).Emailaddress
	$ticker = 1

	foreach ($r in $Rum) {
		$ticker
		$ticker += 1
		$eRum = "$r`:\Kalender"

		$bookpolicy = (Get-CalendarProcessing -Identity $r).BookInPolicy
		$bookpolicy += $anv.LegacyExchangeDN
        Write-Progress -Activity $r -PercentComplete (($ticker/$Rumslista.Count)*100)
		Set-CalendarProcessing -Identity $r -AllBookInPolicy:$false -BookInPolicy $bookpolicy
		Add-MailboxFolderPermission -Identity $eRum -AccessRights LimitedDetails -Confirm:$false -User $anv.PrimarySMTPAddress
	}
}
