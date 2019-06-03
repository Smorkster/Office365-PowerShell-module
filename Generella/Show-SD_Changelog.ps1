<#
.Synopsis
	Visar ändringar av skript
.Description
	Visar ändringar som har gjorts i skript, baserat på namn eller datum. Parametrarna är dynamiska för att visa vilka faktiska alternativ som finns tillgängliga. T.ex. visar SkriptChanges namnet på alla skript som finns listade i changelog och visar då vilka ändringar som har gjorts för valt skript, listat per datum. Eftersom parametrarna är dynamiska, tar de med alla ändringar och datum som listas i changelog. Listorna (skapas i en hashtabell) läses om varje gång modulen blir inläst.
	Om ingen parameter anges, visas senaste changes gjorts vid senaste publicering
.Parameter SkriptChanges
	Listar namnet på alla skript som har någon change i changelog. Vid val av skript, listas alla changes gjorda per datum
.Parameter ChangeDatum
	Listar alla datum som ändringar har publicerats. Vid val av datum, listas alla changes gjorda det datumet
.Example
	Show-SD_Changelog -SkriptChanges Test-SD_Skript
	Hämtar alla ändringar som har gjorts i skript 'Test-SD_Skript'
.Example
	Show-SD_Changelog -ChangeDatum 1970-01-01
	Hämtar alla ändringar som gjordes vid 1970-01-01
#>

function Show-SD_Changelog
{
	[CmdletBinding()]

	param ()

	DynamicParam
	{
		$ParamAttrib = New-Object System.Management.Automation.ParameterAttribute
		$AttribColl = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$AttribColl.Add($ParamAttrib)
		$scriptNames = $global:changeloghash.Keys | sort
		$AttribColl.Add((New-Object  System.Management.Automation.ValidateSetAttribute($scriptNames)))
		$RuntimeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('SkriptChanges',  [string], $AttribColl)

		$ParamAttrib2 = New-Object System.Management.Automation.ParameterAttribute
		$AttribColl2 = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$AttribColl2.Add($ParamAttrib2)
		$changeDates = @()
		$global:changeloghash.Values.GetEnumerator() | % {$changeDates += $_[0]}
		$changeDates = $changeDates | select -Unique
		$AttribColl2.Add((New-Object  System.Management.Automation.ValidateSetAttribute($changeDates)))
		$RuntimeParam2 = New-Object System.Management.Automation.RuntimeDefinedParameter('ChangeDatum',  [string], $AttribColl2)

		$RuntimeParamDic = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
		$RuntimeParamDic.Add('SkriptChanges',  $RuntimeParam)
		$RuntimeParamDic.Add('ChangeDatum',  $RuntimeParam2)

		return  $RuntimeParamDic
	}

	process
	{
		if ($($PSBoundParameters.SkriptChanges)) {
			$stars = "*****************************"
			for($i = 1; $i -lt $($PSBoundParameters.SkriptChanges).Length; $i++) {$stars += "*"}
			Write-Host $stars -Foreground Green
			Write-Host "Ändringar gjorda för skript " -NoNewline
			Write-Host $($PSBoundParameters.SkriptChanges) -Foreground Cyan -NoNewline
			Write-Host "`n$stars`n" -Foreground Green
			$date = $true
			foreach ($text in $global:changeloghash.Item($($PSBoundParameters.SkriptChanges)))
			{
				if ($date)
				{
					Write-Host $text -Foreground Cyan
				} else {
					Write-Host $text
				}
				$date = -not $date
			}
		} elseif ($($PSBoundParameters.ChangeDatum)) {
			Write-Host "********************************" -Foreground Green
			Write-Host "Ändringar som gjordes " -NoNewline
			Write-Host $($PSBoundParameters.ChangeDatum) -Foreground Cyan -NoNewline
			Write-Host "`n********************************`n" -Foreground Green
			foreach ($changeItem in $global:changeloghash.GetEnumerator())
			{
				$a = $changeItem.Value
				if($a[0] -match $($PSBoundParameters.ChangeDatum))
				{
					Write-Host "$($changeItem.Name) " -Foreground Cyan -NoNewline
					$a[1]
				}
			}
		} else {
			$date = ($Global:changelog | select -First 1).ChangeDatum
			Write-Host "*************************************************************" -Foreground Green
			Write-Host "Senaste publiceringen (" -NoNewline
			Write-Host $date -Foreground Cyan -NoNewline
			Write-Host ") gjordes följande ändringar"
			Write-Host "*************************************************************`n" -Foreground Green
			foreach ($change in ($Global:changelog | select -First 1).ChangeText)
			{
				$name = ($change -split " - ")[0]
				$text = ($change -split " - ")[1]
				Write-Host $name " " -ForegroundColor Cyan -NoNewline
				Write-Host $text
			}
		}
	}
}
