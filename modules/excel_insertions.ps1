#EXCEL PORTION~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

# Create a new Excel application instance
$excel = New-Object -ComObject Excel.Application

# Make Excel visible (change to $false to run in background)
$excel.Visible = $true

# Open the target workbook (specify the path to your Excel file)
$workbook = $excel.Workbooks.Open($excelTemplate)

# Access the "QUOTE SHEET" worksheet
$worksheet = $workbook.Sheets.Item("QUOTE SHEET")

# Starting row index in Excel
$startRowIndex = 16

$numRows = $dtJoined3.Rows.Count

for ($i = 0; $i -lt $numRows; $i++) {
    $range = $worksheet.Range("A" + ($startRowIndex + $i) + ":B" + ($startRowIndex + $i))
    $range.EntireRow.Insert([Microsoft.Office.Interop.Excel.XlInsertShiftDirection]::xlShiftDown)
}



# Populate the newly created blank rows with the contents of the Datatable
for ($i = 0; $i -lt $numRows; $i++) {
    # Current row in the DataTable
    $row = $dtJoined3.Rows[$i]

    # Write data to Excel
    $worksheet.Cells.Item($startRowIndex + $i, 1) = $row[0].ToString()  # Column A
    $worksheet.Cells.Item($startRowIndex + $i, 2) = $row[1].ToString()  # Column B
}



$worksheet.Cells.Item(2, 7).Value2 = $customerName
$worksheet.Cells.Item(3, 7).Value2 = $customerStreet
$worksheet.Cells.Item(4, 7).Value2 = $customerCityStateZip
$worksheet.Cells.Item(5, 7).Value2 = $customerContactName
$worksheet.Cells.Item(6, 7).Value2 = $customerPhone

$worksheet.Cells.Item(8, 7).Value2 = $customerName
$worksheet.Cells.Item(9, 7).Value2 = $customerStreet
$worksheet.Cells.Item(10, 7).Value2 = $customerCityStateZip
$worksheet.Cells.Item(11, 7).Value2 = $customerContactName
$worksheet.Cells.Item(12, 7).Value2 = $customerPhone



$worksheet = $workbook.Sheets.Item("FP CALCULATOR")

$startRowIndex = 3
$startRowIndex_1 = 4

for ($i = 0; $i -lt $dtSKU2; $i++) {
    # Current row in the DataTable
    $row = $dtSKU2.Rows[$i]

    # Write data to Excel
    $worksheet.Cells.Item($startRowIndex_1 + $i, 2) = $row[0].ToString()  
    $worksheet.Cells.Item($startRowIndex_1 + $i, 4) = $row[2].ToString()
    $worksheet.Cells.Item($startRowIndex_1 + $i, 6) = $marginSelection_Show
    $worksheet.Cells.Item($startRowIndex_1 + $i, 8) = $months  
}




for ($i = 0; $i -lt $dtPrices1.Rows.Count; $i++) {
    # Current row in the DataTable
    $row = $dtPrices1.Rows[$i]

    # Write data to Excel
    $worksheet.Cells.Item($startRowIndex_1 + $i, 3) = $row[1].ToString()  
    
}




# Save the workbook as a new file
$newFilePath = $excelTemplateNew
$workbook.SaveAs($newFilePath)

# Close the workbook without saving changes (to keep the template unchanged)
$workbook.Close($false)

# Quit Excel
$excel.Quit()


# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()



# Create a new Excel application instance for the .XLS file
$excelXLS = New-Object -ComObject Excel.Application

# Make Excel visible (change to $false to run in background)
$excelXLS.Visible = $true

# Create a new .XLS workbook
$workbookXLS = $excelXLS.Workbooks.Add()
$worksheetXLS = $workbookXLS.Worksheets.Item(1)

# Insert siteStates DataTable starting at column A, row 1
#for ($i = 0; $i -lt $sitesStatesFinal.Rows.Count; $i++) {
#    $worksheetXLS.Cells.Item($i + 1, 1) = $sitesStatesFinal.Rows[$i][0].ToString()
#}

# Initialize an array to hold non-blank rows
$nonBlankRows = @()

# Filter out blank rows from the DataTable
for ($i = 0; $i -lt $sitesStatesFinal.Rows.Count; $i++) {
    if (-not [string]::IsNullOrWhiteSpace($sitesStatesFinal.Rows[$i][0].ToString())) {
        $nonBlankRows += $sitesStatesFinal.Rows[$i]
    }
}

# Adjust the indexArray1 considering the first row will be eventually skipped in insertion
$adjustedIndexArray = @()
foreach ($index in $indexArray1) {
    if ($index -lt $nonBlankRows.Count) {
        $adjustedIndexArray += $index # Adjusting for 0-based index and eventual first row removal
    }
}

# Remove rows based on the adjusted index array
$adjustedRows = $nonBlankRows | Where-Object { $nonBlankRows.IndexOf($_) -notin $adjustedIndexArray }

# Skip the first row of the adjusted results when inserting into Excel
for ($i = 1; $i -lt $adjustedRows.Count; $i++) {
    # Adjust for Excel's 1-based indexing
    $worksheetXLS.Cells.Item($i, 1) = $adjustedRows[$i][0].ToString()
}



# Insert dtPrices DataTable starting at column B, row 1
for ($i = 0; $i -lt $dtPrices1.Rows.Count; $i++) {
    $row = $dtPrices1.Rows[$i]
    $worksheetXLS.Cells.Item($i + 1, 2) = $row[0].ToString() # Column B
    $worksheetXLS.Cells.Item($i + 1, 3) = $row[1].ToString() # Column C
    $worksheetXLS.Cells.Item($i + 1, 4) = $row[2].ToString() # Column D
}

# Add additional column with 13.4343 starting at Column F, Row 1
for ($i = 1; $i -le $dtPrices1.Rows.Count; $i++) {
    $worksheetXLS.Cells.Item($i, 6) = 13.4343 # Column E
}

# Save the .XLS workbook
$newXLSFilePath = $XLSfilePath
$workbookXLS.SaveAs($newXLSFilePath, 56) # 56 is the file format for .xls

# Close the workbook
$workbookXLS.Close($true)

# Quit the Excel application
$excelXLS.Quit()


# Release COM objects for the .XLS file
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheetXLS) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbookXLS) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excelXLS) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
