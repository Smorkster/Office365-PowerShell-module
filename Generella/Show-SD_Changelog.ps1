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
		$AttribColl.Add((New-Object System.Management.Automation.ValidateSetAttribute($global:changeloghash.Keys)))
		$RuntimeParam = New-Object System.Management.Automation.RuntimeDefinedParameter('SkriptChanges', [string], $AttribColl)

		$ParamAttrib2 = New-Object System.Management.Automation.ParameterAttribute
		$AttribColl2 = New-Object System.Collections.ObjectModel.Collection[System.Attribute]
		$AttribColl2.Add($ParamAttrib2)
		$changeDates = @()
		$global:changelog | % {$_.ChangeDatum} | Sort-Object -Descending | % {$changeDates += $_}
		$AttribColl2.Add((New-Object System.Management.Automation.ValidateSetAttribute($changeDates)))
		$RuntimeParam2 = New-Object System.Management.Automation.RuntimeDefinedParameter('ChangeDatum', [string], $AttribColl2)

		$RuntimeParamDic = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
		$RuntimeParamDic.Add('SkriptChanges', $RuntimeParam)
		$RuntimeParamDic.Add('ChangeDatum', $RuntimeParam2)

		return  $RuntimeParamDic
	}

	process
	{
		if ($($PSBoundParameters.SkriptChanges)) {
			WriteTitle -Stars ($($PSBoundParameters.SkriptChanges).Length + 28) -Text $($PSBoundParameters.SkriptChanges)
			$date = $true
			foreach ($text in $global:changeloghash.Item($($PSBoundParameters.SkriptChanges)))
			{
				if ($date)
				{
					Write-Host "`n$($text)" -Foreground Cyan
				} else {
					writetext -text $text -name $null
				}
				$date = -not $date
			}
		} elseif ($($PSBoundParameters.ChangeDatum)) {
			WriteTitle -Stars 32 -Text $($PSBoundParameters.ChangeDatum)
			foreach ($changeItem in ($global:changelog | ? {$_.ChangeDatum -eq $($PSBoundParameters.ChangeDatum)}).ChangeText)
			{
				if ( ($split = ($changeItem -split " - ")).Count -gt 1)
				{
					writetext -text $split[1] -name $split[0]
				} else {
					writetext -text $changeItem
				}
			}
		} else {
			WriteTitle
			foreach ($change in ($Global:changelog | select -First 1).ChangeText)
			{
				$name = ($change -split " - ")[0]
				#Write-Host "$($name) " -ForegroundColor Cyan -NoNewline
				writetext -text ($change -split " - ")[1] -name $name
			}
		}
	}
}

function writetext
{
	param (
		[string] $text,
		[string] $name
	)

	$splittext = $text -split " / "

	if ($name)
	{
		Write-Host $name -Foreground Cyan
		foreach ($t in $($text -split "/ "))
		{
			Write-Host "`t$t"
		}
	} else {
		$splittext
	}
}

function WriteTitle
{
	param(
		[int] $Stars = 65,
		[string] $Text
	)

	for ($i = 0 ; $i -lt $Stars; $i++)
	{
		$starsText += "*"
	}

	Write-Host $starsText -Foreground Green
	if ($Text -match "\d\d\d\d-\d\d-\d\d")
	{
		Write-Host "Ändringar som gjordes " -NoNewline
		Write-Host $Text -Foreground Cyan
	} elseif ($Text -match "-SD_") {
		Write-Host "Ändringar gjorda för skript " -NoNewline
		Write-Host $Text -Foreground Cyan
		Write-Host "$($starsText)" -Foreground Green
		return
	} else {
		Write-Host "Vid senaste publiceringen (" -NoNewline
		Write-Host ($Global:changelog | select -First 1).ChangeDatum -Foreground Cyan -NoNewline
		Write-Host ") gjordes följande ändringar"
	}
	Write-Host "$($starsText)`n" -Foreground Green
}
