<#
.SYNOPSIS
	Lista grupper en dator Ã¤r medlem av
.PARAMETER DatorNamn
	Namn pÃ¥ dator att undersÃ¶ka
.DESCRIPTION
	HÃ¤mtar AD-objektet fÃ¶r datornamnet och listar de gruuper som har kopplats
.Example
	Get-SD_DatorsGruppmedlemskap -DatorNamn "Dat1"
#>

function Get-SD_DatorsGruppmedlemskap
{
	param(
	[Parameter(Mandatory=$true)]
		[String] $DatorNamn
	)

	try
	{
		(Get-ADComputer -Identity $DatorNamn -Properties *).MemberOf | % {($_ -split ',')[0]} | % {$_.Substring(3)}
	} catch {
		Write-Host "Dator med namn $DatorNamn finns inte i AD"
	}
}
