<#
.SYNOPSIS
	Hämtar en lista över vilka som kan skapa bokningar i ett rum
.Parameter RumsNamn
	Namn på rummet som efterfrågas
.Parameter Sync
	Anger att gruppmedlemarna ska synkroniseras till Exchange
.Parameter Export
	Ange om användarnas mailadresser ska exporteras till en fil
.Example
	Get-SD_RumBokningsbehörighet -RumsNamn "RumA"
#>

function Get-SD_RumBokningsbehörighet
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $RumsNamn,
		[switch] $Sync,
		[switch] $Export
	)

	try {
		$RumsNamnExchange = $RumsNamn.Trim()+":\Kalender"
		$RumsNamnAzure = "Res-" + $RumsNamn.Trim() + "-Book"
		$AzureGroup = Get-AzureADGroup -SearchString $RumsNamnAzure
		$usersExchange = @()
		$usersAzure = @()
		$notSynced = @()

		if ($AzureGroup -eq $null)
		{
			Write-Host "Inget rum med namn " -NoNewline
			Write-Host $RumsNamn -ForegroundColor Cyan -NoNewline
			Write-Host " hittades.`nAvslutar"
			return
		}
		Get-AzureADGroupMember -ObjectId $AzureGroup.ObjectId -All $true -ErrorAction Stop | % {$usersAzure += $_.UserPrincipalName}
		if($usersAzure.Count -gt 0)
		{
			Get-MailboxFolderPermission -Identity $RumsNamnExchange -ErrorAction Stop | ? {$_.User -notlike "Standard" -and $_.User -notlike "Anonymous"} | % {$usersExchange += $_.User.ADRecipient.UserPrincipalName}

			$usersAzure | % {
				if ($usersExchange -notcontains $_){
					$notSynced += $_
				}
			}

			Write-Host "Dessa har behörighet att skapa bokningar i rum " -NoNewline
			Write-Host $RumsNamn -ForegroundColor Cyan
			$usersExchange | sort | % {Write-Host "`t "$_}
			if($notSynced.Count -gt 0)
			{
				if ($Sync)
				{
					Write-Host "`n$($notSynced.Count) har inte blivit synkade med bokningsbehörighet till Exchange"
					Write-Host "`nInitierar synkronisering från Azure till Exchange" -ForegroundColor Cyan
					Set-AzureADGroup -ObjectId (Get-AzureADGroup -SearchString $RumsNamnAzure).ObjectId -Description Now -ErrorAction Stop
					$ticker = 1
					foreach ($ns in $notSynced) {
						Write-Progress -Activity "Lägger på behörighet $ticker av $($notSynced.Count)" -PercentComplete (($ticker/$notSynced.Count)*100)
						try {
							Add-MailboxFolderPermission -Identity $RumsNamnExchange -AccessRights LimitedDetails -User $ns -Confirm:$false -ErrorAction Stop | Out-Null
						} catch {
							if ($_.CategoryInfo.Reason -eq "InvalidExternalUserIdException") {
								$address = ($_.Exception -split [char]0x22)[1]
								Write-Host "Adress $address finns inte. Personen har troligen slutat."
							} elseif ($_.CategoryInfo.Reason -eq "ACLTooBigException") {
								Write-Host "`nFör många medlemmar i Azure-gruppen. Kan inte synkronisera fler till Exchange.`n`nAvslutar därför hanteringen."
								return
							} else {
								$_
							}
						}
						$ticker++
					}
					$bp = (Get-CalendarProcessing -Identity $RumsNamn -ErrorAction Stop).BookInPolicy += $notSynced
					try {
						Set-CalendarProcessing -Identity $RumsNamn -BookInPolicy $bp -ErrorAction Stop
					} catch {
						if ($_.CategoryInfo.Reason -eq "CmdletProxyException")
						{
							Write-Host "En eller flera personer hade redan behörighet i kalendern"
						} elseif ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException" -and $_.CategoryInfo.Activity -eq "Set-CalendarProcessing") {
							Write-Host "Hittade inte rummets kalender för tilläggning av behörighet"
						}
					}
				} else {
					Write-Host "`nDessa $($notSynced.Count) har inte blivit synkade med bokningsbehörighet till Exchange"
					$notSynced = $notSynced | sort
					$notSynced

				}
			}
		} else {
			Write-Host "`nGruppen för bokningsbehörighet i Azure är tom.`nInga unika behörigheter har skapats, alla kan boka rummet."
		}

		if($Export)
		{
			$excel = New-Object -ComObject excel.application 
			$excel.visible = $false
			$excelWorkbook = $excel.Workbooks.Add()
			$excelTempsheet = $excelWorkbook.Worksheets.Add()
			$row = 1
			$excelTempsheet.Cells.Item($row, 1) = "Rumsamn"
			$excelTempsheet.Cells.Item($row, 1).Font.Bold = $true
			$excelTempsheet.Cells.Item($row, 2) = $RumsNamn
			$row = $row + 2
			$excelTempsheet.Cells.Item($row, 1) = "Bokningsbehörighet:"
			$excelTempsheet.Cells.Item($row, 1).Font.Bold = $true
			$row++

			foreach ($user in $usersExchange) {
				if ($user -notlike "")
				{
					$excelTempsheet.Cells.Item($row, 1) = $user
					$row++
				}
			}
			$excelRange = $excelTempsheet.UsedRange
			$excelRange.EntireColumn.AutoFit() | Out-Null
			$excelTempsheet.ListObjects.Add(1, $excelTempsheet.Range($excelTempsheet.Cells.Item(3,1),$excelTempsheet.Cells.Item($excelTempsheet.usedrange.rows.count, 1)), 0, 1) | Out-Null

			$excelWorkbook.SaveAs("H:\Kan boka i '$RumsNamn'.xlsx")
			$excelWorkbook.Close()
			Write-Host "Användarna exporterade till din H:`n" -NoNewline
			Write-Host "(H:\Kan boka i '$RumsNamn'.xlsx)" -ForegroundColor Cyan

			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelRange) | Out-Null
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelTempsheet) | Out-Null
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelWorkbook) | Out-Null
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()
			Remove-Variable excel
		}
	} catch [Microsoft.Open.AzureAD16.Client.ApiException] {
		Write-Host "Problem att hitta Azure-gruppen för behörigheter. Kontrollera att den finns och är korrekt kopplad."
	} catch {
		if ($_.CategoryInfo.Reason -eq "ManagementObjectNotFoundException" -and $_.CategoryInfo.Activity -eq "Get-MailboxFolderPermission") {
			Write-Host "Rummets maillåda gick inte att hitta"
		} else {
			Write-Host "Fel uppstod i körningen:"
			$_
		}
	}
}
