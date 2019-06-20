<#
.Synopsis
	Lägger till och tar bort mailadresser till en distributionslista
.Description
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
		$group = $_.Group
		if ($_.Action.Trim() -eq "Add")
		{
			Write-Host "($ticker av $numberOfEntries) Lägg till " -NoNewline
			Write-Host $_.Email.Trim() -Foreground Cyan -NoNewline
			Write-Host " i distributionslistan " -NoNewline
			Write-Host $group -Foreground Cyan
		} elseif ($_.Action.Trim() -eq "Remove") {
			Write-Host "($ticker av $numberOfEntries) Ta bort " -NoNewline
			Write-Host $_.Email.Trim() -Foreground Cyan -NoNewline
			Write-Host " från distributionslistan " -NoNewline
			Write-Host $group -Foreground Cyan
		} else {
			Write-Host "($ticker av $numberOfEntries) $($_.Action.Trim()) " -NoNewline
			Write-Host $_.Email.Trim() -Foreground Cyan -NoNewline
			Write-Host " på distributionslistan " -NoNewline
			Write-Host $group -Foreground Cyan
			Write-Host "`tAngiven action '$($_.Action.Trim())' följer inte standard och rad $($ticker+1) i Excel-filen kommer inte hanteras.`n"
			return
		}

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
			if((Get-Mailbox -Identity $emailToAdd -ErrorAction SilentlyContinue) -or (Get-Contact -Identity $emailToAdd -ErrorAction SilentlyContinue))
			{
				Write-Host "`tAdress finns i Exchange." -Foreground Green
			} else {
				Write-Host "`tInget kontaktobjekt hittades i Exchange, skapar" -Foreground Cyan
				New-MailContact -Name $emailToAdd -ExternalEmailAddress $emailToAdd | Out-Null
				Set-MailContact -Identity $emailToAdd -HiddenFromAddressListsEnabled $true | Out-Null
				Write-Host "`tKontaktobjekt skapat." -Foreground Green
			}
			#endregion Create contact object

			Write-Host "`tLägger till i distributionslista" -Foreground Cyan
			try
			{
				Add-DistributionGroupMember -Identity $azureGroup.Identity -Member $emailToAdd -ErrorAction Stop
				Write-Host "`tAdress tillagd"
			} catch {
				if( $_.CategoryInfo.Reason -eq "MemberAlreadyExistsException") {
					Write-Host "`tMailadress finns redan i distributionslistan`n" -Foreground Red
				} else {
					$_
				}
			}
		#endregion Add user
		} elseif ($_.Action.Trim() -eq "Remove") {
		#region Remove user
			Write-Host "`tTar bort adress från distributionslista" -Foreground Yellow
			Remove-DistributionGroupMember -Identity $group.Trim() -Member $_.Email -Confirm:$false -ErrorAction SilentlyContinue
			Write-Host "`tAdress borttagen`n" -Foreground Green
		#endregion Remove user
		}

		if($Error[0]) { $Error[0] = "" }
		$ticker = $ticker + 1
	}
	$Excel.Workbooks.Close()
	$Excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($temp) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}
