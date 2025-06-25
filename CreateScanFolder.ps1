# Create a network scan folder for a new user


do{
    $Failed = $false
    try{
        # Get username to copy
        $username = Read-Host "Please enter the first initial and last name for the folder"
        # Get the user, based on their "samAccountName"
        $user = Get-ADUser $username
    
    }

    catch {
        Write-Host "Could not find the user >>>$($username)<<<`n"
        $Failed = $true 
    }
} while ($Failed)



# Define the network path, get username input, and define folder name
$NetworkPath = "<\\network\Path>"  # Replace with your network path


# Combine the path and user folder name
$UserFolderPath = Join-Path -Path $NetworkPath -ChildPath $username

# Ensure the network path is accessible
if (Test-Path $NetworkPath) {
    # Run the script
    New-Item -Path "$UserFolderPath" -ItemType Directory
    Write-Host "Created the folder!"
} else {
    Write-Host "The network path does not exist: $NetworkPath" -ForegroundColor Red
}

$ScansFolder = "Scans"

# Combine the path and scan name
$ScansPath = Join-Path -Path $UserFolderPath -ChildPath $ScansFolder

# Ensure the network path is accessible
if (Test-Path $UserFolderPath) {
    # Run the script
    New-Item -Path "$ScansPath" -ItemType Directory
    Write-Host "Created the folder!"
} else {
    Write-Host "The user folder path does not exist: $UserFolderPath" -ForegroundColor Red
}



# Change the user's samAccountName as home directory
Set-ADUser -Identity $user.SamAccountName -HomeDirectory $UserFolderPath -HomeDrive Y:


$acl = Get-Acl $UserFolderPath
$control = "<DOMAIN>\$username","FullControl", "ContainerInherit,ObjectInherit", "None","Allow"  # Replace with your domain
$controlAccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule $control
$acl.AddAccessRule($controlAccessRule)



Set-Acl $UserFolderPath $acl
