function Show-SD_Meny
{
	$ticker = 0
	$CommandKeys = (Get-Module | ? {$_.Name -like "*Servicedesk*"}).ExportedCommands.Keys
	$items = $CommandKeys | % {$_ -split ("-") | select -First 1} | select -Unique
	$length = $CommandKeys | % {$_.Length} | sort | select -Last 1
	$row = New-Object PSObject

	foreach ($i in $items){
		Write-Host $i -ForegroundColor DarkBlue -BackgroundColor White
		foreach ( $name in $CommandKeys | ? { $_ -match $i } ){
			if( $ticker -lt 3 ){ Write-Host $name -NoNewLine; 1..( $length - $name.length ) | % { Write-Host " " -NoNewLine }; $ticker += 1 }
			else {Write-Host $name; $ticker = 0}
		}
		Write-Host "`n"
		$ticker = 0
	}
}

$ModuleFunctions = Get-ChildItem -Path "$PSScriptRoot\Servicedesk*" | ? {$_.Mode -match "d"} | % { Get-ChildItem $_ } | ? {$_.Name -like "*.ps1"}
$ToExport = $ModuleFunctions | Select-Object -ExpandProperty BaseName

foreach ($import in $ModuleFunctions)
{
	try
	{
		. $import.FullName
	} catch {
		Write-Error "Import misslyckad: $($import.FullName): $_"
	}
}

Export-ModuleMember -Function $ToExport
Export-ModuleMember -Function Show-SD_Meny
