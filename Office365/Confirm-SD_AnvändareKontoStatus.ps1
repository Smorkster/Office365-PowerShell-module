<#
.Synopsis
	Kontrollera status på olika delar av ett mailkonto
.Parameter id
	id för användaren
.Parameter AllTests
	Switch för att köra alla kontroller
.Description
	Ange id för användaren och skriptet kommer kontrollera alla steg som utförs vid synkronisering. För varje kontroll skrivs en rapport i commandofönstret.
	Ifall något steg fallerar, kommer skriptet stoppas.
.Example
	Confirm-SD_AnvändareKontoStatus -id "ABCD"
#>
function Confirm-SD_AnvändareKontoStatus
{
	param(
	[Parameter(Mandatory=$true)]
		[String] $id,
		[switch] $AllTests
		)

	try
	{
		$user = Get-ADUser -Identity $id -Properties *
	} catch [Microsoft.ActiveDirectory.Management.ADIdentityNotFoundException] {
		Write-Host "AD-konto har inte skapats`nAlla tester avslutas."
		return
	}
	$fail = $false

	if($user.EmailAddress -eq $null)
	{
		Write-Host "Mailattributet ej synkroniserat till AD" -Foreground Red
		$fail = $true
		if(-not $AllTests)
		{
			return
		}
	} else {
		Write-Host "Mailattribut finns i AD" -Foreground Green
	}

	if(-not $user.Enabled)
	{
		Write-Host "AD-konto ej aktivt (Disabled)" -Foreground Red
		$fail = $true
		if(-not $AllTests)
		{
			return
		}
	} else {
		Write-Host "AD-konto är aktivt" -Foreground Green
	}

	if($user.LockedOut)
	{
		Write-Host "AD-konto är låst" -Foreground Red
		$fail = $true
		if(-not $AllTests)
		{
			return
		}
	} else {
		Write-Host "AD-konto är inte låst" -Foreground Green
	}

	if($user.msExchMailboxGuid -ne $null)
	{
		Write-Host "msExchMailboxGuid är inte tomt i AD" -Foreground Red
		$fail = $true
		if(-not $AllTests)
		{
			return
		}
	} else {
		Write-Host "msExchMailboxGuid är tomt i AD" -Foreground Green
	}

	try
	{
		$userAzure = Get-MsolUser -UserPrincipalName $user.EmailAddress
	} catch {
		Write-Host "O365-konto har inte skapats. Avbryter resten av testerna." -Foreground Red
		return
	}

	try
	{
		$haveLicens = $false
		$licenses = $userAzure.Licenses | select accountskuid | % {$_ -match "pack"}
		foreach($l in $licenses)
		{
			if($l -eq $true)
			{
				$haveLicens = $true
			}
		}
		if(-not $haveLicens)
		{
			Write-Host "E3-licens saknas" -Foreground Red
			$fail = $true
			if(-not $AllTests)
			{
				return
			}
		} else {Write-Host "E3-licens finns" -Foreground Green}

		$userGroups = Get-AzureADUser -SearchString $user.EmailAddress | Get-AzureADUserMembership | ? {$_.DisplayName -contains "o365-migpilots"}
		if($userGroups -eq $null)
		{
			Write-Host "Är inte medlem i O365-MigPilots" -Foreground Red
			$fail = $true
			if(-not $AllTests)
			{
				return
			}
		} else {Write-Host "Är medlem i O365-MigPilots" -Foreground Green}
		
		$userExchange = Get-Mailbox -Identity $user.EmailAddress -ErrorAction SilentlyContinue
		if($userExchange -eq $null)
		{
			Write-Host "Mailbox inte skapad i Exchange" -Foreground Red
			$fail = $true
			if(-not $AllTests)
			{
				return
			}
		} else {
			Write-Host "Mailbox skapad i Exchange" -Foreground Green
		}
	} catch [Microsoft.Online.Administration.Automation.MicrosoftOnlineException] {
		Write-Host "O365-konto har inte skapats. Avbryter resten av testerna." -Foreground Red
		$fail = $true
		if(-not $AllTests)
		{
			return
		}
	}
	if(-not $fail)
	{
		Write-Host "Allt ska fungera"
	}
}
