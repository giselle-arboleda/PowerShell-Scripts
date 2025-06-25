

# Install ImportExcel module
# Install-Module -Name ImportExcel -Force

# Import the Excel file and store data in a variable
$filePath = "<PATH\to\\Excel\file>"  # Replace with path to Excel file
$excelData = Import-Excel -Path $filePath -WorksheetName "EMPLOYEE_MASTER"


Write-Host "Updating lists of employee emails and phones..."
Write-Host "-------------------"

# Create a list of combined first and last names
$nameList = @()
for ($row = 0; $row -lt $excelData.Count; $row++) {
    $fullName = $excelData[$row]."First Name"
   
    
    # Remove ", JR." from the last name if present
    $lastName = $excelData[$row]."Last Name".Split(',')[0]

    
#     # Combine first name and modified last name
#     $fullName = $fullName 
    $fullName = $excelData[$row]."First Name" + "*" + $lastName
    $fullName = $fullName -replace "'", "*"
    $fullName = $fullName -replace " ", ""
    $fullName = $fullName -replace "-", "*"

    Write-Host $fullName
    
    $nameList += $fullName
}

# # Export the result to a CSV file
$nameList | Out-File -FilePath "<Path\\To\OUT\CSV\NameFile>" -Encoding utf8 # Replace with path to resulting CSV NameFile



# Create a list of emails
$emailList = @()
# Create a list of phone numbers
$phoneList = @()

# $count = 0

# Connect to Outlook
$Outlook = New-Object -ComObject Outlook.Application
$Namespace = $Outlook.GetNamespace("MAPI")

# Get the Global Address List (GAL)
$GAL = $Namespace.AddressLists | Where-Object { $_.Name -eq "Offline Global Address List" }

#Begin iterating through names list
for ($row = 0; $row -lt $excelData.Count; $row++) {
    $name = $nameList[$row]
    if ($excelData[$row]."Location Description".contains("POLICE")){
        ##DOESN'T WORK YET
        $Contact = $GAL.AddressEntries | Where-Object { $_.Name -like "*$name*" }
        if ($Contact -ne $null) {
            $emailList += $Contact.GetExchangeUser().PrimarySmtpAddress
            $phoneList += $Contact.GetExchangeUser().BusinessTelephoneNumber
        }
        else {
            $emailList += "N/A"
            $phoneList += "N/A"
        }

    }
    elseif ([bool] !(Get-ADUser -Filter "Name -like '*$name*'")) {
        $emailList += "N/A"
        $phoneList += "N/A"
    }
    else {

            $user = Get-ADUser -Filter "Name -like '*$name*'" -Properties * | Select-Object SamAccountName
            $username = $user.SamAccountName
            $user = Get-ADUser -Identity $username -Properties TelephoneNumber, EmailAddress
         if ($user.EmailAddress -ne $null) {
                $emailList += $user.EmailAddress
            }
        else{
             $emailList += 'N/A'
        }
        if ($user.TelephoneNumber -ne $null) {
                $phoneList += $user.TelephoneNumber
            }
        else{
             $phoneList += 'N/A'
        }
    }

}

#Close
$Outlook.Quit()
Remove-Variable Outlook

# Export the result to a CSV file
$phoneList | Out-File -FilePath "<Path\\To\OUT\CSV\PhoneFile>" -Encoding utf8   # Replace with path to resulting CSV PhoneFile
# Export the result to a CSV file
$emailList | Out-File -FilePath "<Path\\To\OUT\CSV\EmailFile>" -Encoding utf8   # Replace with path to resulting CSV EmailFile

Write-Host "-------------------"
Write-Host "Updated list complete!!!"
Write-Host "-------------------"