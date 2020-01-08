function FillChangelog
{
	$global:changeloghash = @{} #Hashtable
	$global:changelog = @() #Array
	$change = $null
	$changeText = @()
	foreach($changeTextRow in (Get-Content $PSScriptRoot\Changelog.txt))
	{
		# Row contains a date, create new changeobject
		if ($changeTextRow -match "\d{4}[-]\d{2}[-]\d{2}")
		{
			if ($change -eq $null)
			{
				$change = New-Object PSObject
				$change | Add-Member -MemberType NoteProperty -Name ChangeDatum -Value $changeTextRow
			}
		# Row is empty, no more changes for current date, add data to changelog
		} elseif ($changeTextRow -eq "") {
			$change | Add-Member -MemberType NoteProperty -Name ChangeText -Value $changeText
			$global:changelog += $change
			$changeText = @()
			$change = $null
		} else {
			if ($changeTextRow -notmatch "\*\*\*\*\*\*\*\*")
			{
				# Check if row describes changes to a script or general information
				if ($changeTextRow.Split(" ")[0] -match "-SD_" -or $changeTextRow.Split(" ")[0] -match "Starta")
				{
					$split = $changeTextRow -split " - "
					$scriptName = $split[0]
					$scriptChangeText = $split[1..$split.Count]
					if ($scriptName -match "-")
					{
						if($global:changeloghash.ContainsKey($scriptName))
						{
							$global:changeloghash.$scriptName += @($change.ChangeDatum; $scriptChangeText)
						} else {
							$global:changeloghash.Add($scriptName, @($change.ChangeDatum; $scriptChangeText))
						}
					}
				}

				$changeText += $changeTextRow
			}
		}
	}
}

function FillTypes
{
	$script:commandTypes = @()
	$script:commandTypes = $ToExport | % {$_ -split ("-") | select -First 1} | sort | select -Unique
}

$ModuleFunctions = Get-ChildItem -Path "$PSScriptRoot\Servicedesk*" -Directory | % { Get-ChildItem $_ } | ? { $_.Name -like "*.ps1" }
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
FillChangelog
FillTypes
Write-Host "Version $(($Global:changelog | select -First 1).ChangeDatum) är klar att användas" -ForegroundColor Green
