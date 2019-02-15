<#
.SYNOPSIS
	Lista alla grupper en användare är medlem i
.PARAMETER id
	Användarens id
.Example
	Get-SD_GrupperAnvändareÄrMedlemI -id "ABCD"
#>

function Get-SD_AnvändareGruppmedlemskap
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $id
	)

	$user = Get-ADUser -Identity $id -Properties *
	try {
		Write-Host $user.GivenName $user.Surname -Foreground Magenta -NoNewline
		Write-Host " är medlem i grupperna:"
		Get-AzureADUser -SearchString $user.EmailAddress | Get-AzureADUserMembership | sort DisplayName | ft DisplayName
	} catch {
		Write-Host "`nAnvändaren finns inte i Exchange"
	}
}
