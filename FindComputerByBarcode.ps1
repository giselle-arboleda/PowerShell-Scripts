# Find computer name by barcode number


# Create Excel COM object
$Excel = New-Object -ComObject Excel.Application 


# Open the workbook
$Workbook = $Excel.Workbooks.Open("Path\To\File")  # Replace with path to Excel workbook 


# Open the worksheet and display 
# the worksheet name
$workSheet = $Workbook.Sheets.Item(2)

$Range = $Worksheet.Range("B2").EntireColumn



do{
    $barcodeBool = $false

    # Get barcode
    $barcode = Read-Host "Enter barcode"

    # Find barcode
    $Search = $Range.find($barcode)

    # If barcode is found, get the value of
    # the "D" column and display the
    # computer name.
    if ($Search){
        $searchRow = $Search.row()

        $cellValue = $workSheet.Cells.Item($SearchRow, 4).Value()
        Write-Host "`nComputer name: $cellValue" -ForegroundColor "Cyan"
        Write-Host "Barcode: $barcode`n" -ForegroundColor "Cyan"
         
    }

    # Repeat until an existing barcode
    # is input.
    else {
        Write-Host "`nCouldn't find the barcode!`n`n" -ForegroundColor "Yellow"
        $barcodeBool = $true
    }   
    
} while ($barcodeBool)



# Close workbook when done
$Workbook.close($false)
