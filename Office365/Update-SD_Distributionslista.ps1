<#
.SYNOPSIS
	Lägger till och tar bort mailadresser till en distributionslista
.DESCRIPTION
	Läser in en CSV-fil innehållandes mailadresser, namn på distributionslista samt vad som ska göras.
	När skriptet startar, öppnas filen i Excel så att den kan editeras. När filen är sparad, fortsätter skriptet efter användaren har tryck på Enter.
	Om en extern mailadress inte finns i Exchange, skapas ett kontakt-objekt med adressen. Adressen läggs sedan till eller tas bort från distributionslistan. Om det är en intern mailadress som inte finns, hoppas adressen över.
	
	Varje rad i filen har följande struktur: Email	Group	Action
	Exempelvis: test@test.com	Distlista	Remove
	
	Kolumen Action får enbart innehålla:
		Add
		Remove
.Example
	Update-SD_Distributionslista
	Läser in mailadresser och action från Excel-filen och raderar respektive lägger till adressen i distributionslistan som anges i filen
#>

function Update-SD_Distributionslista
{
	#Variables
	Write-Host "Öppnar Excel-fil. Editera adresser som ska läggas till/tas bort. Stäng sedan Excel för att fortsätta." -Foreground Cyan
	$fil = "G:\\\Epost & Skype\Powershell\FilerTillSkripten\UpdateDistributionlist.csv"
	$Excel = New-Object -ComObject Excel.Application
	$Excel.Visible = $true
	$temp = $Excel.Workbooks.Open($fil)
	$ticker = 1
	Read-Host "Fortsätt genom att trycka Enter..."

	$data = Import-Csv -Delimiter ";" -Encoding UTF7 -Path $fil
	$numberOfEntries = ($data | measure).Count

	$data | foreach {
		Write-Host "Uppdatering $ticker av $numberOfEntries"
		$ticker = $ticker + 1
		$group = $_.Group
		try
		{
			$azureGroup = Get-DistributionGroup -Identity $group.Trim() -ErrorAction Stop
		} catch {
			Write-Host "Distributionslista '$group' finns inte i Exchange.`n"
			return
		}
		if ($_.Action.Trim() -eq "Add")
		{
		#region Add user
			$emailToAdd = $_.Email.Trim()
			#region Create contact object
			if ($emailToAdd -match "@test.com")
			{
				if(Get-Mailbox -Identity $emailToAdd -ErrorAction SilentlyContinue)
				{
					Write-Host "Maillåda för $emailToAdd finns." -Foreground Green
				} else {
					Write-Host "Ingen maillåda för $emailToAdd finns.`nHoppar över.`n" -Foreground Red
					return
				}
			} elseif (Get-MailContact -Identity $emailToAdd -ErrorAction SilentlyContinue) {
				Write-Host "Kontakt" $emailToAdd "finns i Exchange" -Foreground Green
			} else {
				Write-Host "Ingen kontakt för" $emailToAdd "hittades i Exchange, skapar" -Foreground Cyan
				New-MailContact -Name $emailToAdd -ExternalEmailAddress $emailToAdd | Out-Null
				Set-MailContact -Identity $emailToAdd -HiddenFromAddressListsEnabled $true | Out-Null
				Write-Host "Färdig skapa kontakt." -Foreground Green
			}
			#endregion Create contact object

			Write-Host "Lägger till" $emailToAdd "i grupp" $azureGroup.DisplayName -Foreground Cyan
			try { Add-DistributionGroupMember -Identity $azureGroup.Identity -Member $emailToAdd -ErrorAction Stop }
			catch {
				if( $_.CategoryInfo.Reason -eq "MemberAlreadyExistsException") {
					Write-Host $emailToAdd "finns redan i grupp" $azureGroup.DisplayName "`n"-Foreground Red
				} else {
					#Write-Host "Klar`n" -Foreground Green
					$_
				}
			}
		#endregion Add user
		} elseif ($_.Action.Trim() -eq "Remove") {
		#region Remove user
			Write-Host "Tar bort" $_.Email.Trim() "från" $group.Trim() -Foreground Yellow
			Remove-DistributionGroupMember -Identity $group.Trim() -Member $_.Email -Confirm:$false -ErrorAction SilentlyContinue
			Write-Host "Klar med borttag`n" -Foreground Green
		#endregion Remove user
		} else {
			Write-Host "`n`nAngiven action '$($_.Action.Trim())' följer inte standard och $($_.Email.Trim()) kommer inte hanteras.`n`n"
		}

		if($Error[0]) { $Error[0] = "" }
	}
	$Excel.Workbooks.Close()
	$Excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($temp) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}
