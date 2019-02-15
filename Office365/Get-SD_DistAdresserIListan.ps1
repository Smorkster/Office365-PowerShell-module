<#
.SYNOPSIS
	Hämtar alla adresser i en distributionslista
.PARAMETER Distlista
	Namn på distributionslistan
.PARAMETER EndastExterna
	Används för att enbart hämta de externa adresserna i distributionslistan
.PARAMETER Exportera
	Anger ifall datan ska exporteras till en CSV-fil
.Example
	Get-SD_DistAdresserIListan -Distlista "Distlista"
.Example
	Get-SD_DistAdresserIListan -Distlista "Distlista" -EndastExterna
#>

function Get-SD_DistAdresserIListan
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $Distlista,
		[switch] $EndastExterna,
		[switch] $Exportera
	)

	$dl = Get-DistributionGroup -Identity $Distlista -ErrorAction SilentlyContinue
	if($dl -eq $null)
	{
		Write-Host "Ingen distributionslista med namn " -NoNewline
		Write-Host $Distlista -NoNewline -Foreground Cyan
		Write-Host " hittades"
	} else {
		if($EndastExterna)
		{
			$members = Get-DistributionGroupMember -Identity $Distlista.Trim() | Where-Object {$_.RecipientType -like "MailContact"}
		} else {
			$members = Get-DistributionGroupMember -Identity $Distlista.Trim() | sort Name
		}

		if($Exportera)
		{
			$export = @()
			$members | % {$export += [pscustomobject]@{"Namn"=$_.DisplayName; "Mail"=$_.PrimarySMTPAddress}}
			$ExcelObject=New-Object -ComObject Excel.Application
			$WorkBook=$ExcelObject.WorkBooks.Add(1)
			$WorkSheet=$WorkBook.WorkSheets.Item(1)
			$ExcelObject.Visible=$false
			$ExcelObject.DisplayAlerts = $false
			$row = 2
			
			$WorkSheet.Cells.Item(1, 'A').Value2 = "Namn"
			$WorkSheet.Cells.Item(1, 'B').Value2 = "Mail"
			foreach($a in $export)
			{
				$WorkSheet.Cells.Item($row, 'A').Value2 = $a.Namn
				$WorkSheet.Cells.Item($row, 'B').Value2 = $a.Mail
				$row++
			}
			$WorkSheet.UsedRange.EntireColumn.autofit() > $null
			$WorkSheet.ListObjects.Add(1, $WorkSheet.UsedRange, 0, 1) | Out-Null

			$WorkBook.SaveAs("H:\Medlemmar i distributionslista '$Distlista'.xlsx")
			$ExcelObject.WorkBooks.Close()
			$ExcelObject.Quit()
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkSheet) | Out-Null
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($WorkBook) | Out-Null
			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelObject) | Out-Null
			[System.GC]::Collect()
			[System.GC]::WaitForPendingFinalizers()

			Write-Host "$($row-2) medlemar från distributionslista '$dl', har exporterats till:"
			Write-Host "H:\Medlemmar i distributionslista '$Distlista'.xlsx" -Foreground Green
		} else {
			$members | ft DisplayName, PrimarySMTPAddress
		}
	}
}
