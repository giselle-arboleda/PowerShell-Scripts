# Script to create an AD user account & optionally add the user
# to the same group memberships & permissions as an existing
# user. 

# Contributors: Matthew Castellanos & Giselle Arboleda




do {
    #Don't mess with these, they determine if the user's account already exists, if it will have a middle initial; and will run the program accordingly.
    $accountExists = $true
    # $uniqueAccount = $false

    #Prompts the user to enter the first name
    while ($accountExists -eq $true) {

        $firstName = Read-Host -Prompt "Enter the new user's first name "

        #First inital of name to apply to the SamAccountName (follows company policy)
        $firstInitial = $firstName[0]

        #Prompts the user to enter the last name
        $lastName = Read-Host -Prompt "Enter the new user's last name "

        #Checking if username already exists.
        $userName = "$firstInitial$lastName"
        $user = Get-ADUser $userName -ErrorAction SilentlyContinue
        
        if (-not $user) {
            Write-Host "This name can be used"
            $accountExists = $false
        } else {
            Write-Host "A user with this name already exists"
            $middleInitial = Read-Host "This name already exists, please input a middle initial (Capital Letter Only)" 
            $userName = "$firstInitial$middleInitial$lastName"         

            #This needs to be in the 'else' statement otherwise it won't work properly
            try {
                Get-ADUser $userName -ErrorAction Stop | Out-Null
                Write-Host "Error creating account. This account also exists. Please start from the beginning"
            } catch {
                Write-Host "This name can be used" 
                $accountExists = $false
            } 
        }      
    }

  
    #Testing if the inputs are correct
    Write-Host "The user's full name is"$firstName $lastName "and their username is"$userName

    #Saves the password and will be used during the account creation later in the script.  
    $password = Read-Host -Prompt "Enter the user's temporary password " -AsSecureString

    while ($true) {   
        $tempUser = Read-Host -Prompt "Is the user an intern or a temp? (Type 'intern' or 'temp'. Type 'no' if not an intern or temp)"
        
        #Creates the user account if intern or temp
        if ($tempUser -eq "intern" -or $tempUser -eq "temp" ) {
            New-ADuser -Name "$firstName $lastName" `
            -GivenName $firstName `
            -Surname $lastName `
            -DisplayName "$firstName $lastname-$tempUser" `
            -SamAccountName $userName"-"$tempUser `
            -UserPrincipalName $userName"-"$tempUser"@yourWorkDomain.org" `  # REPLACE WITH YOUR WORK'S EMAIL DOMAIN
            -EmailAddress $userName"-"$tempUser"@yourWorkDomain.org" `  # REPLACE WITH YOUR WORK'S EMAIL DOMAIN
            -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
            -ChangePasswordAtLogon $true `
            -Enabled $true
            break
        
        } elseif ($tempUser -eq "no") {
            New-ADUser -Name "$firstName $lastName" `
            -GivenName $firstName `
            -Surname $lastName `
            -DisplayName "$firstName $lastname" `
            -SamAccountName $userName `
            -UserPrincipalName $userName"@yourWorkDomain.org" `  # REPLACE WITH YOUR WORK'S EMAIL DOMAIN
            -EmailAddress $userName"@yourWorkDomain.org" `  # REPLACE WITH YOUR WORK'S EMAIL DOMAIN
            -AccountPassword (ConvertTo-SecureString -AsPlainText $password -Force) `
            -ChangePasswordAtLogon $true `
            -Enabled $true
            break
            
        } else {
            Write-Host "Please type 'intern', 'temp', or 'no' if the user is neither of these"
        }
    }
    Write-Host "Account created"
    $groupMembership = Read-Host "Do you want to copy group memberships from another account? (y / n)"
    
        if ($groupMembership -eq "y") {
            
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






            # Add username to the same groups
            $memberships | Add-ADGroupMember -Members $userName



            # Get group memberships of the username
            $newMemberships = (Get-Aduser $username -Properties MemberOf | Select-Object MemberOf).MemberOf

            # Display RESULTING group memberships of the username
            Write-Host "`n`n---------- $($username)'s resulting Memberships ----------"
            Write-Output $newMemberships | Format-List
        
            Write-Host "`n`nGroup memberships have been copied (Press ENTER)"
            Read-Host
        } else {
            Write-Host "`n`nNo group memberships have been copied (Press ENTER)"
            Read-Host
        } 
    Clear-Host
} until ($quit)