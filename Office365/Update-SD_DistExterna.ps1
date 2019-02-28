<#
.SYNOPSIS
	Lägger till och tar bort externamailadresser till en distributionslista
.DESCRIPTION
	Läser in en CSV-fil innehållandes externa mailadresser, namn på distributionslista samt vad som ska göras.
	När skriptet startar, öppnas filen i Excel så att den kan editeras. När det är klart fortsätter skriptet efter användaren har tryck på Enter.
	Om en adress inte finns i Exchange, skapas ett kontakt-objekt med adressen. Adressen läggs sedan till eller tas bort från distributionslistan.
	
	Varje rad i filen har följande struktur: Email	Group	Action
	Exempelvis: test@test.com	Distlista	Remove
	
	Kolumen Action får enbart innehålla:
		Add
		Remove
.Example
	Update-SD_DistExterna
#>

function Update-SD_DistExterna
{
	#Variables
	Write-Host "Öppnar Excel-fil. Editera adresser som ska läggas till/tas bort. Stäng sedan Excel för att fortsätta." -Foreground Cyan
	$fil = "G:\\\Epost & Skype\Powershell\FilerTillSkripten\ExternalContacts_Batch.csv"
	$Excel = New-Object -ComObject Excel.Application
	$Excel.Visible = $true
	$temp = $Excel.Workbooks.Open($fil)
	$ticker = 1
	Read-Host "Fortsätt genom att trycka Enter..."

	$data = Import-Csv -Delimiter ";" -Encoding UTF7 -Path $fil
	$numberOfEntries = $data.Count
	$errorLevel

	$data | foreach {
		#region Add user
		Write-Host "Uppdatering $ticker av $numberOfEntries"
		$ticker = $ticker + 1
		$group = $_.Group
		try
		{
			$errorLevel = 1
			$azureGroup = Get-DistributionGroup -Identity $group.Trim() -ErrorAction Stop
		} catch {
			Write-Host "Distributionslista '$group' finns inte i Exchange.`n"
			return
		}
		if ($_.Action.Trim() -eq "Add")
		{
			#region Create contact object
			if (Get-MailContact -Identity $_.Email.Trim() -ErrorAction SilentlyContinue)
			{
				Write-Host "Kontakt" $_.Email.Trim() "finns i Exchange" -Foreground Green
			} else {
				Write-Host "Ingen kontakt för" $_.Email.Trim() "hittades i Exchange, skapar" -Foreground Yellow
				New-MailContact -Name $_.Email.Trim() -ExternalEmailAddress $_.Email.Trim() | Out-Null
				Set-MailContact -Identity $_.Email.Trim() -HiddenFromAddressListsEnabled $true | Out-Null
				Write-Host "Färdig skapa kontakt." -Foreground Green
			}
			#endregion Create contact object

			Write-Host "Lägger till" $_.Email.Trim() "i grupp" $azureGroup.DisplayName -Foreground Yellow
			$errorLevel = 2
			Add-DistributionGroupMember -Identity $azureGroup.Identity -Member $_.Email.Trim()# -ErrorAction SilentlyContinue
			if($Error[0] -like "*is already a member of the group*")
			{
				Write-Host $_.Email.Trim() "finns redan i grupp" $azureGroup.DisplayName "`n"-Foreground Red
			} else {
				Write-Host "Klar`n" -Foreground Green
			}
		}
		#endregion Add user

		#region Remove user
		if ($_.Action.Trim() -eq "Remove")
		{
			Write-Host "Tar bort" $_.Email.Trim() "från" $group.Trim() -Foreground Yellow
			Remove-DistributionGroupMember -Identity $group.Trim() -Member $_.Email -Confirm:$false -ErrorAction SilentlyContinue
			Write-Host "Klar med borttag`n" -Foreground Green
		}
		#endregion Remove user

		if($Error[0]) { $Error[0] = "" }
	}
	$Excel.Workbooks.Close()
	$Excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($temp) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}
