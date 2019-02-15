<#
.SYNOPSIS
	Listar samtliga skapade rumslistor
#>

function Get-SD_Rumslistor
{
	Get-DistributionGroup -Filter {RecipientTypeDetails -eq "RoomList"} | sort DisplayName | ft DisplayName, Identity
}
