# Find computer name by barcode number

do{
# Create Excel COM object
$Excel = New-Object -ComObject Excel.Application 


# Open the workbook
$Workbook = $Excel.Workbooks.Open("Path/To/YOUR/ExcelWorkbook.xlsx")    # REPLACE WITH YOUR PATH


# Open the worksheet and display 
# the worksheet name
$workSheet = $Workbook.Sheets.Item(2)





do{
    $barcodeBool = $false

    # Get barcode
    $barcode = Read-Host "Enter barcode"

    # Find barcode
    $lastRow = $workSheet.UsedRange.Rows.Count
    $found = $false

    for ($i = 1; $i -le $lastRow; $i++) {
        $cellValue = $workSheet.Cells.Item($i, 2).Value2
        if ($cellValue -eq $barcode) {
            $found = $true
            $searchRow = $i
            break
        }
    }

    if ($found) {
        Write-Host "`n`nBarcode found!"
        $cellValue = $workSheet.Cells.Item($searchRow, 4).Value2
        Write-Host "Computer name: $cellValue`n`n" -ForegroundColor "Cyan"
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
} until ($quit)