# Script to copy membership of one AD account and apply it to another


do{
    $Failed = $false
    try{
        # Get username to copy
        $copyUsername = Read-Host "Enter a username of the account to RETRIEVE the group memberships"
        # Get group memberships of the copyUsername
        $memberships = (Get-Aduser $copyUsername -Properties MemberOf | Select-Object MemberOf).MemberOf
    
    }

    catch {
        Write-Host "Could not find the user >>>$($copyUsername)<<<"
        Write-Host "Please enter a new one"
        $Failed = $true 
    }
} while ($Failed)

# Display group memberships of the user to be copied
Write-Host "`n---------- Existing user: $($copyUsername)'s Memberships ----------"
Write-Output $memberships | Format-List


do{
    $Failed = $false
    try{
        # Get username to apply to
        $pasteUsername = Read-Host "Enter a username of the account to APPLY the group memberships"

    }

    catch {
        Write-Host "Could not find the user >>>$($pasteUsername)<<<"
        Write-Host "Please enter a new one"
        $Failed = $true 
    }
} while ($Failed)





# Add pasteUsername to the same groups
# !!!NOTE!!! --> Un-Comment when ready to test and actually copy group memberships
$memberships | Add-ADGroupMember -Members $pasteUsername



# Get group memberships of the pasteUsername
$newMemberships = (Get-Aduser $pasteUsername -Properties MemberOf | Select-Object MemberOf).MemberOf

# Display RESULTING group memberships of the pasteUsername
Write-Host "`n`n---------- $($pasteUsername)'s resulting Memberships ----------"
Write-Output $newMemberships | Format-List


