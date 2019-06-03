<#
.SYNOPSIS
	Kopierar PowerShell-script till centrala mappen på G:
.DESCRIPTION
	Hämtar alla PowerShell-script från lokal och centrala mappen, gör en jämförelse om något har ändrats lokalt och uppdaterar i så fall.
.Example
	Copy-SD_KopieraTillG
#>

function Copy-SD_KopieraSkriptTillG
{
	$localDir = "H:\Programmering\Powershell"
	$centralDir = "G:\\\Epost & Skype\Powershell\"
	$local = Get-ChildItem $localDir -File -Recurse
	$central = Get-ChildItem $centralDir -File -Recurse
	$MD5 = New-Object -TypeName System.Security.Cryptography.MD5CryptoServiceProvider
	$oförändrade = 0
	$ticker = 1

	foreach ($fileA in $local) {
		Write-Progress -Activity "Hanterar fil ($ticker / $($local.Count))" -PercentComplete (($ticker/$local.Count)*100) -CurrentOperation $fileA.Name
		$fileB = $central | ? {$_.Name -eq $fileA.Name}
		if ($fileB -eq $null)
		{
			Write-Host "Ny fil: " -NoNewline
			Write-Host $fileA.Name -ForegroundColor Cyan
			if ( ($fileA.DirectoryName -split "\\" | select -Last 1) -eq "Powershell" )
			{
				$newFile = $centralDir + "\" + $fileA.Name
			} else {
				$newFile = $centralDir + ($fileA.DirectoryName -split "\\" | select -Last 1) + "\" + $fileA.Name
			}
			Copy-Item $fileA.FullName -Destination $newFile
		} else {
			$fileAHash = [System.BitConverter]::ToString($MD5.ComputeHash([System.IO.File]::ReadAllBytes($fileA.FullName)))
			$fileBHash = [System.BitConverter]::ToString($MD5.ComputeHash([System.IO.File]::ReadAllBytes($fileB.FullName)))
			if ($fileAHash -ne $fileBHash)
			{
				Write-Host "Uppdaterar " -NoNewline
				Write-Host $fileA.Name -ForegroundColor Cyan
				Copy-Item $fileA.FullName -Destination $fileB.FullName
			} else {
				$oförändrade += 1
			}
		}
		$fileB = $null
		$ticker++
	}
	Write-Host "$oförändrade filer oförändrade"
}
