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

function FillChangelog
{
	$global:changeloghash = @{} #Hashtable
	$global:changelog = @() #Array
	$change = $null
	$changeText = @()
	$changeSummary = ""
	foreach($changeTextRow in (Get-Content $PSScriptRoot\Changelog.txt))
	{
		if ($changeTextRow -match "\d{4}[-]\d{2}[-]\d{2}")
		{
			if ($change -eq $null)
			{
				$change = New-Object PSObject
				$change | Add-Member -MemberType NoteProperty -Name ChangeDatum -Value $changeTextRow
			}
		} elseif ($changeTextRow -eq "") {
			$change | Add-Member -MemberType NoteProperty -Name ChangeText -Value $changeText
			$change | Add-Member -MemberType NoteProperty -Name ChangeSummary -Value $changeSummary
			$global:changelog += $change
			$changeText = @()
			$changeSummary = ""
			$change = $null
		} else {
			if ($changeTextRow -notmatch "\*\*\*\*\*\*\*\*")
			{
                $split = $changeTextRow -split " - "
                $one = $split[0]
                $rest = $split[1..$split.Count]
				if ($one -match "-")
				{
					$changeSummary += "$one`n"
                    if($global:changeloghash.ContainsKey($one))
                    {
                        $global:changeloghash.$one += @($change.ChangeDatum; $rest)
                    } else {
                        $global:changeloghash.Add($one, @($change.ChangeDatum; $rest))
                    }
				}
			}
			if ($changeTextRow -notmatch "\*\*\*\*\*\*\*\*")
			{
				if (($changeTextRow -split " - ")[0] -match "-")
				{
					if ($changeSummary -eq "")
					{
						$changeSummary += "$(($changeTextRow -split " - ")[0])`n"
					} else {
						$changeSummary += "$(($changeTextRow -split " - ")[0])`n"
					}
				}
				$changeText += $changeTextRow
			}
		}
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
$version = Get-Content .\Changelog.txt -First 1
FillChangelog
Write-Host "Version $version är klar att användas" -ForegroundColor Green
