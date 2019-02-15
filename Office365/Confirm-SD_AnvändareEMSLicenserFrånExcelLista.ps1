<#
.SYNOPSIS
	Verifiera om användare har fått EMS-licens.
.PARAMETER SökvägTillFil
	Ange sökväg till filen med användarna
.SYNTAX
	Confirm-SD_AnvändareEMSLicenserFrånExcelLista -SökvägTillFil <Sökväg>
.DESCRIPTION
	Lista över användare hämtas från Excel-fil angiven av användare.
#>

function Confirm-SD_AnvändareEMSLicenserFrånExcelLista
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $SökvägTillFil
	)

	$donthave = New-Object System.Collections.ArrayList
	$notfound = New-Object System.Collections.ArrayList
	$users = New-Object System.Collections.ArrayList
	$testarrayid = @("hsa","id","id","id")
	$ticker = 0
	$need = 0

	#region Readfile
		Write-Host "Läser Excel-fil..." -Foreground Cyan -nonewline
		$column=1
		$objExcel = New-Object -ComObject Excel.Application
		$workbook = $objExcel.Workbooks.Open($SökvägTillFil)
		$sheet = $workbook.Worksheets.Item("Blad1")
		$objExcel.Visible = $false
		for(;$column -le $($sheet.UsedRange.Columns).Count-1;$column++)
		{
			if($testarrayid -contains $sheet.Cells.Item(1, $column).Text)
			{
				break
			}
		}

		for ($i=1; $i -le $($sheet.UsedRange.Rows).Count-1; $i++)
		{
			$users.Add($sheet.Cells.Item($i+1, $column).Text) > $null
		}
		#Close Excel-file
		$workbook.Close()
		$objExcel.quit()
	#endregion Readfile

	#region CheckUsers
		Write-Host "Startar test av id..." -Foreground Cyan
		foreach($id in $users)
		{
			$user = Get-ADUser -Identity $id
			$name = $user.GivenName + " " + $user.Surname

			try{
				$ticker = $ticker + 1
				$userinfo = Get-MsolUser -SearchString $name | where {$_.ImmutableId -match $id}
				if(-not ($userinfo.Licenses.AccountSku.SkuPartNumber).Contains('EMS'))
				{
					$donthave.Add($userinfo.DisplayName +" - " +$userinfo.UserPrincipalName) > $null
					$need = $need + 1
				}
			}catch{
				$notfound.Add($user.$column) > $null
			}
		}
	#endregion CheckUsers

	Write-Host "Dessa har inte EMS licens" -Foreground Cyan
	$donthave | fl
	Write-Host "`nDessa har inte skapats i Exchange" -Foreground Cyan
	$notfound | fl
	Write-Host "`nAntal kontrollerade:" $ticker "Antal som saknar EMS:" $need

	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
}
