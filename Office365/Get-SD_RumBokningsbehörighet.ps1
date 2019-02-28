<#
.SYNOPSIS
	Hämtar en lista över vilka som kan skapa bokningar i ett rum
.PARAMETER RumsNamn
	Namn på rummet som efterfrågas
.PARAMETER Export
	Ange om användarnas mailadresser ska exporteras till en fil
.Example
	Get-SD_RumBokningsbehörighet -RumsNamn "RumA"
#>

function Get-SD_RumBokningsbehörighet
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $RumsNamn,
		[switch] $NoSync,
		[switch] $Export
	)

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
	try {
		Get-AzureADGroupMember -ObjectId $AzureGroup.ObjectId | % {$usersAzure += $_.UserPrincipalName}
		if($usersAzure.Count -gt 0)
		{
			Get-MailboxFolderPermission -Identity $RumsNamnExchange | ? {$_.User -notlike "Standard" -and $_.User -notlike "Anonymous"} | % {$usersExchange += $_.User.ADRecipient.UserPrincipalName}

			$usersAzure | % {
				if ($usersExchange -notcontains $_){
					$notSynced += $_
				}
			}

			Write-Host "Dessa har behörighet att skapa bokningar i rum " -NoNewline
			Write-Host $RumsNamn -ForegroundColor Cyan
			$usersExchange | sort | % {Write-Host "`t "$_}
			if (-not $NoSync)
			{
				if($notSynced.Count -gt 0)
				{
					Write-Host "`nDessa har inte blivit synkade med bokningsbehörighet till Exchange"
					$notSynced | % {write $_}
					Write-Host "`nInitierar synkronisering från Azure till Exchange" -ForegroundColor Cyan
					Set-AzureADGroup -ObjectId (Get-AzureADGroup -SearchString $RumsNamnAzure).ObjectId -Description Now
				}
			}
		} else {
			Write-Host "`nGruppen för bokningsbehörighet i Azure är tom.`nInga unik behörigheter har skapats, alla kan boka rummet."
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
	}
}
