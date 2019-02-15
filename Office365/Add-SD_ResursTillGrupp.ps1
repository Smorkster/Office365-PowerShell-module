<#
.SYNOPSIS
	Lägg in en/flera resurser eller rum i en administrations-/Azuregrupp
.PARAMETER ResursNamn
	Namn/mailadress för resursen/rummet
.Parameter GruppNamn
	Namn på gruppen som resurserna/rummen ska läggas in i
.DESCRIPTION
	Skriptet används för att lägg objekt till en administrationsgrupp i Azure.
	Använder ingen logik för att särskilja grupper, så om flera objekt har samma början i namnet, t.ex. samma organisation och lokalitet, kommer samtliga objekt placeras i administrationsgruppen.
.Example
	Add-SD_ResursTillGrupp -ResursNamn "RumA" -GruppNamn "GruppA"
	Rum och resurser flyttas till en grupp genom att söka på "RumA" och ange gruppnamn "GruppA"
#>

function Add-SD_ResursTillGrupp
{
	param(
	[Parameter(Mandatory=$true)]
		[string] $ResursNamn,
	[Parameter(Mandatory=$true)]
		[string] $GruppNamn
	)

	$rooms = Get-Mailbox -Identity "$ResursNamn*" -Filter {ResourceType -eq "Room"}
	$equipments = Get-Mailbox -Identity "$ResursNamn*" -Filter {ResourceType -eq "Equipment"}
	$groupObjectID = Get-MsolGroup -SearchString $GruppNamn

	try
	{
		$rooms | % {Get-MsolGroup -SearchString "res-$_-admins"} | % Add-MsolGroupMember -GroupObjectId $groupObjectID.ID -GroupMemberType Group -GroupMemberObjectId $_.ObjectID
	} catch {
		Write-Host "Rum finns redan i gruppen"
	}

	try
	{
		$equipments | % {Get-MsolGroup -SearchString "res-$_-admins"} | % Add-MsolGroupMember -GroupObjectId $groupObjectID.ID -GroupMemberType Group -GroupMemberObjectId $_.ObjectID
	} catch {
		Write-Host "Resurs finns redan i gruppen"
	}
}
