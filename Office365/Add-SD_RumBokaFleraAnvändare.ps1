<#
.SYNOPSIS
	Lägger till/tar bort användare för behörighet att boka i ett eller flera rum.
.PARAMETER Remove
	Parameter för att ange ifall användarna ska tas bort.
.DESCRIPTION
	En batch-körning för att lägga in eller ta bort flera användare i flera rum. Användarna och rummen läses in från två filer som öppnas i början av skriptet.
	Varje användare från AddUserToRoomBookUsers.txt läggs in i varje rum AddUserToRoomBookGroups.txt
.Example
	Add-SD_RumBokaFleraAnvändare
	Lägger in varje användare i varje rum
.Example
	Add-SD_RumBokaFleraAnvändare -Remove
	Tar bort varje användare frän varje rum
#>

function Add-SD_RumBokaFleraAnvändare
{
	param(
		[switch] $Remove
	)

	#Variables
	$User = 0
	$Target = 0
	$Failed = @()
	$tickerTotal = 1
	$tickerUser = 1

	$fileUsers = "G:\\\Epost & Skype\Powershell\FilerTillSkripten\AddUserToRoomBookUsers.csv"
	$Excel = New-Object -ComObject Excel.Application
	$Excel.Visible = $true
	$temp = $Excel.Workbooks.Open($fileUsers)
	Read-Host "Editera filen för användare och tryck sedan Enter"
	$fileGroups = "G:\\\Epost & Skype\Powershell\FilerTillSkripten\AddUserToRoomBookGroups.csv"
	$Excel.Visible = $true
	$temp2 = $Excel.Workbooks.Open($fileGroups)
	Read-Host "Editera filen för grupper/rum och tryck sedan Enter"

	$InputUsers = Get-Content -Path $fileUsers
	$InputGroups = Get-Content -Path $fileGroups
	$numGroups = $InputGroups.Count
	$numUsers = $InputUsers.Count
	$numberOfEntries = $numGroups * $numUsers
	$Excel.Workbooks.Close()
	$Excel.Quit()

	$InputUsers | foreach {
		if ($User = (Get-MsolUser -UserPrincipalName $_.Trim() -ErrorAction SilentlyContinue).ObjectID)
		{
			Write-Host "("$tickerUser"/"$numUsers" )`tUser - " -Foreground Yellow -NoNewline
			Write-Host $_ -Foreground Cyan

			$InputGroups | foreach {
				$name = "res-"+$_.Trim()+"-book"
				if ($Target = (Get-MsolGroup -MaxResults 100000 -SearchString $name -ErrorAction SilentlyContinue).ObjectID){

					if ($Remove)
					{
					#region Remove user
						Write-Host "("$tickerTotal"/"$numberOfEntries") `tRemoving from..." -Foreground Yellow -NoNewline
						Write-Host $_ -Foreground Cyan -NoNewline
						try { Remove-MsolGroupMember -GroupObjectID $Target -GroupMemberType 'User' -GroupMemberObjectId $User -ErrorAction Stop } catch {}
						if($Error[0] -like "*The member you are trying to delete is not in this group*")
						{
							Write-Host " is not a member." -Foreground Red
						} else {
							Write-Host "... Done" -Foreground Green
						}
					#endregion
					} else {
					#region Add user
						Write-Host "("$tickerTotal"/"$numberOfEntries") `tAdding to " -Foreground Yellow -NoNewline
						Write-Host $_ -Foreground Cyan -NoNewline
						try { Add-MsolGroupMember -GroupObjectID $Target -GroupMemberType 'User' -GroupMemberObjectId $User -ErrorAction Stop	} catch {}
						if($Error[0] -like "*is already a member of this group*")
						{
							Write-Host " is already a member." -Foreground Red
						} else {
							Write-Host "... Done" -Foreground Green
						}
					#endregion
					}

					Set-MsolGroup -ObjectID $Target -Description "Now"
				} else {
					Write-Host "Group" $_ "not found..." -Foreground Red
				}
			$tickerTotal = $tickerTotal + 1
			}
		} else {
			$Failed += $_
			Write-Host "User $_ not found in Azure. Skipping." -Foreground Red
		}
		$tickerUser = $tickerUser + 1
		Write-Host "Next user`n..."
		try {$Error[0] = ""} catch {}
	}
	if($Failed.Count -gt 0)
	{
		Write-Host "Dessa kunde inte läggas till i någon grupp:"
		$Failed
	}
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($temp2) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($temp) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
}
