<#
.Synopsis
    Exporterar alla användare i AD-grupp till CSV-fil
.Parameter Grupper
    En lista på grupper som ska exporteras
.Description
    Hämtar alla användare för varje given AD-grupp och exporterar dessa till en CSV-fil
.Example
    Get-SD_ExporteraADAnvändare -Grupper "Grupp1","Grupp2"
#>

function Get-SD_ExporteraADAnvändare
{
    param(
	[Parameter(Mandatory=$true)]
        [string[]] $Grupper
    )

    $members = @()

    foreach ($group in $Grupper) {
        foreach ($i in (Get-ADGroupMember -Identity $group)) {
			if($i.objectClass -eq "computer")
			{
				$member = Get-ADComputer -Identity $i
				$members += [pscustomobject]@{"Namn"=$member.Name}
			} elseif ($i.objectClass -eq "user") {
				$member = Get-ADUser $i -Properties *
				$members += [pscustomobject]@{"Namn"=$member.GivenName + " " + $member.Surname
				"Epost"=$member.EmailAddress
				"id"=$member.identity}
			}
        }
    }

	if ($Grupper.Count -gt 1)
	{$filename = "H:\Exporterade gruppbehörigheter från AD"}
	else {$filename = "H:\Gruppbehörigheter för $Grupper.csv"}
    $members | Export-csv -Path $filename -Encoding Unicode
    Write-Host "Användarna exporterade till $filename"
}
