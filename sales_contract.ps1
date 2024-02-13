Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data.DataSetExtensions
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName System.Data

# Set the absolute path to your config.ps1 script
$configPath = "D:\Programming\PowerShell\Sales Contract\config.ps1"
# Dot source the config script
. $configPath



# Show a message box with the desired message
[Windows.Forms.MessageBox]::Show('SELECT THE COVER PAGE', '')


$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    InitialDirectory = [Environment]::GetFolderPath('Desktop')
    Title = "Select the cover page"
    Filter = "Word Documents (*.docx;*.doc)|*.docx;*.doc"
}

$word = New-Object -ComObject Word.Application
$word.Visible = $true

if ($FileBrowser.ShowDialog() -eq 'OK') {
    $wordDocumentPath = $FileBrowser.FileName
    $doc = $word.Documents.Open($wordDocumentPath)

    if ($doc.Tables.Count -ge 1) {
        $table = $doc.Tables.Item(1)
        $coverPage_Table = New-Object System.Data.DataTable

        # Copy the table
        $table.Range.Copy()
        Write-Host "Table copied to clipboard."

        for ($rowIndex = 1; $rowIndex -le $table.Rows.Count; $rowIndex++) {
            $row = $table.Rows.Item($rowIndex)
            $dataRow = $coverPage_Table.NewRow()

            for ($colIndex = 1; $colIndex -le $row.Cells.Count; $colIndex++) {
                $cell = $row.Cells.Item($colIndex)
                $cellText = $cell.Range.Text.TrimEnd("`r", "`a")

                if ($rowIndex -eq 1) {
                    $coverPage_Table.Columns.Add($cellText)
                } else {
                    $dataRow[$colIndex - 1] = $cellText
                }
            }

            if ($rowIndex -gt 1) {
                $coverPage_Table.Rows.Add($dataRow)
            }
        }

        # Uncomment to print the DataTable to the console
        #foreach ($row in $coverPage_Table.Rows) {
        #    $row.ItemArray -join ", " | Write-Host
        #}
    }

    $doc.Close([Microsoft.Office.Interop.Word.WdSaveOptions]::wdDoNotSaveChanges)
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}




# Show a message box with the desired message
[Windows.Forms.MessageBox]::Show('SELECT THE PDF QUOTE', '')

function Select-PdfFile {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $fileDialog.Filter = 'PDF files (*.pdf)|*.pdf'
    $fileDialog.ShowDialog() | Out-Null
    return $fileDialog.FileName
}



$pdfToTextPath = "C:\Users\seanc\Downloads\xpdf-tools-win-4.04\xpdf-tools-win-4.04\bin64\pdftotext.exe"
$pdfFilePath = Select-PdfFile


if ([string]::IsNullOrWhiteSpace($pdfFilePath)) {
    Write-Host "No file selected."
} else {
    $outputTxtPath = [System.IO.Path]::ChangeExtension($pdfFilePath, '.txt')
    & $pdfToTextPath -table $pdfFilePath $outputTxtPath

    if (Test-Path $outputTxtPath) {
        $textContent = Get-Content $outputTxtPath -Raw

        Write-Host $textContent


     # Regex to extract the relevant section

     $regex0 = if ($textContent -match "Quotation[\s\S]+?Quoted") { $matches[0] } else { "" }

     $regex1 = if ($textContent -match "Item Description[\s\S]+?Final Quote") { $matches[0] } else { "" }

     # Removing pricing details to create $regex2
     $regex2 = $regex1 -replace "\d+\s*\d{1,3},\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\[[^\]]*\]|\$\s*\d+\s*\d*\.\d{2}|\d+\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}", ""
    }

}

$term = "Payment Tenure\s*:\s*(\d+)\s*Months"

if ($textContent -match $term) {
    $months = $matches[1]
    "Tenure: $months months"
} else {
    "Pattern not found."
}

$shipping = "Shipping\s*Cost\s*for\s*(\d{1,3}) Qty\s*\$\s*([\d\.]+)"

if ($textContent -match $shipping) {
    $quantity = $matches[1]
    $price = $matches[2]
    "Quantity: $quantity, Price: $price"
} else {
    "Pattern not found."
}



# Create the form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Please Select Margin"
$form.Size = New-Object System.Drawing.Size(300, 200)
$form.StartPosition = "CenterScreen"

# Add the 26% button
$button26 = New-Object System.Windows.Forms.Button
$button26.Location = New-Object System.Drawing.Point(30, 50)
$button26.Size = New-Object System.Drawing.Size(100, 23)
$button26.Text = "26%"
$button26.Add_Click({
    $script:marginSelection = 0.74
    $form.Close()
})
$form.Controls.Add($button26)

# Add the 35% button
$button35 = New-Object System.Windows.Forms.Button
$button35.Location = New-Object System.Drawing.Point(150, 50)
$button35.Size = New-Object System.Drawing.Size(100, 23)
$button35.Text = "35%"
$button35.Add_Click({
    $script:marginSelection = 0.65
    $form.Close()
})
$form.Controls.Add($button35)

# Show the form
$form.Add_Shown({$form.Activate()})
$form.ShowDialog() | Out-Null


# Line by line capture for descriptions, quantity, and prices


$customerInfo = [regex]::Matches($regex0, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value

$sitesStates = [regex]::Matches($regex1, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value
# Print the $sitesStates variable to the terminal
#Write-Host "Sites States:"
#$sitesStates | ForEach-Object { Write-Host $_ }

# Line by Line capture to build datatable with descriptions only
$sitesStatesRegex2 = [regex]::Matches($regex2, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value
# Print the $sitesStatesRegex2 variable to the terminal
#Write-Host "`nSites States Regex2:"
#$sitesStatesRegex2 | ForEach-Object { Write-Host $_ }


# Creating Data Tables

$customerInfoDT = New-Object System.Data.DataTable
$customerInfoDT.Columns.Add("Column1", [string])

$sitesStatesDT = New-Object System.Data.DataTable
$sitesStatesDT.Columns.Add("Column1", [string])

$sitesStatesFinal = New-Object System.Data.DataTable
$sitesStatesFinal.Columns.Add("Column1", [string])

$dtSKU = New-Object System.Data.DataTable
$dtSKU.Columns.Add("SKU Column", [string])

$dtPrices = New-Object System.Data.DataTable
$dtPrices.Columns.Add("Qty", [decimal])
$dtPrices.Columns.Add("List Price", [decimal])
$dtPrices.Columns.Add("Total Price", [decimal])
$dtPrices.Columns.Add("MRC", [decimal])


$dtJoined3 = New-Object System.Data.DataTable
$dtJoined3.Columns.Add("Description", [string])
$dtJoined3.Columns.Add("Item Cost", [string])
$dtJoined3.Columns.Add("Quantity", [string])
$dtJoined3.Columns.Add("MRC", [string])

$dtPrices2 = New-Object System.Data.DataTable
$dtPrices2.Columns.Add("MRC Unit Price", [string])
$dtPrices2.Columns.Add("Units", [double])
$dtPrices2.Columns.Add("MRC Total", [string])



#building customerInfoDT datatable
$rowCount = $customerInfo.Count
$counter = 0

while ($counter -le $rowCount - 1) {
    # Fetch the current line based on $counter
    $line = $customerInfo[$counter]
    
    # Check if the line is not just whitespace or empty and does not contain "Quotation" or "Billing"
    if (-not [string]::IsNullOrWhiteSpace($line) -and $line -notmatch "Quotation" -and $line -notmatch "Billing") {
        # If the line has content and does not contain "Quotation" or "Billing", add it to the DataTable
        $customerInfoDT.Rows.Add($line)
    }
    
    # Increment counter to move to the next line
    $counter++
}


# Loop through each row in the DataTable
foreach ($row in $customerInfoDT.Rows) {
    # Get the current value of Column1
    $currentText = $row["Column1"]
    
    # Define a regex pattern that matches the specified keywords (with possible multiple whitespaces between words) and any text following them
    $pattern = "(Quote\s+No.*|Quote\s+Date.*|Valid\s+Until.*|Payment\s+Term.*)|Quoted|$telcoSales"

    # Replace matched patterns with an empty string, effectively removing them
    $updatedText = $currentText -replace $pattern, ""

    # Update the row's text with the modified value
    $row["Column1"] = $updatedText.Trim() # .Trim() is used to remove any leading or trailing whitespace that might be left
}





#building sitesStatesDT
$rowCount = $sitesStatesRegex2.Count
$counter = 0

while ($counter -le $rowCount - 1) {
    $sitesStatesDT.Rows.Add($sitesStatesRegex2[$counter])
    $counter++
}


# Create a new DataTable for filtered results
$sitesStatesFiltered = New-Object System.Data.DataTable
$sitesStatesFiltered.Columns.Add("Column1", [string])

# Iterate through each row and apply the filtering logic
foreach ($row in $sitesStatesDT.Rows) {
    $columnValue = $row["Column1"]
    if (-not ([string]::IsNullOrWhiteSpace($columnValue)) -and
        -not ($columnValue.StartsWith("Total") -or
              $columnValue.Contains("Sub Total") -or
              $columnValue.Contains("Shipping") -or
              $columnValue.Contains("Item Description") -or
              $columnValue.Contains("Final") -or
              $columnValue.Contains("Price"))) {
        # Add the row to the filtered DataTable
        $filteredRow = $sitesStatesFiltered.NewRow()
        $filteredRow["Column1"] = $columnValue
        $sitesStatesFiltered.Rows.Add($filteredRow)
    }
}


for ($i = 0; $i -lt $sitesStatesFiltered.Rows.Count; $i++) {
    $currentCell = $sitesStatesFiltered.Rows[$i][0] # Assuming data is in the first column

    # Check if the current cell contains a newline character
    if ($currentCell -like "*SIM + Wifi*") {
        # Check if there is a next row
        if ($i -lt $sitesStatesFiltered.Rows.Count - 1) {
            $nextCell = $sitesStatesFiltered.Rows[$i + 1][0]
            
            # Merge the next cell content into the current cell, removing the newline character
            $sitesStatesFiltered.Rows[$i][0] = $currentCell.TrimEnd("`n") + $nextCell

            # Optionally, clear the content of the next cell instead of removing it
            $sitesStatesFiltered.Rows[$i + 1][0] = ""
        }
    }
}

# Remove blank rows at the end
for ($i = $sitesStatesFiltered.Rows.Count - 1; $i -ge 0; $i--) {
    $currentCell = $sitesStatesFiltered.Rows[$i][0]
    if (-not $currentCell -or $currentCell -eq "") {
        $sitesStatesFiltered.Rows.RemoveAt($i)
    }
}



# Trigger string
$strSplitTrigger = "SubTotal"

# Find indexes of rows containing the trigger string
$indexArray1 = @()
for ($i = 0; $i -lt $sitesStatesFiltered.Rows.Count; $i++) {
    if ($sitesStatesFiltered.Rows[$i][0].ToString().Contains($strSplitTrigger)) {
        $indexArray1 += $i
    }
}


# Create a new DataTable for the final result, with the same schema as sitesStatesFiltered
$sitesStatesFinal = New-Object System.Data.DataTable
$sitesStatesFiltered.Columns | ForEach-Object {
    $sitesStatesFinal.Columns.Add($_.ColumnName, $_.DataType)
}

# Iterate over each row in sitesStatesFiltered
foreach ($row in $sitesStatesFiltered.Rows) {
    # Create a new row for sitesStatesFinal
    $newRow = $sitesStatesFinal.NewRow()

    # Apply the regular expression replacement to the specific column (e.g., "Column1") and then trim the result
    $newRow["Column1"] = [System.Text.RegularExpressions.Regex]::Replace($row["Column1"].ToString(), "Bundle\s+SubTotal|\$|\d{1,2},?\d{3}\.\d{2}", "").Trim()

    # Add the modified row to sitesStatesFinal
    $sitesStatesFinal.Rows.Add($newRow)
}





# Assuming $regex1 is a string containing the input text for regex matches
$regexSKU = [System.Text.RegularExpressions.Regex]::new("\[[^\]]*\]")
$matchSKU = $regexSKU.Matches($regex1) | ForEach-Object { $_ }

$regexPricesPattern = "\d+\s*\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}"

$regexPrices = [System.Text.RegularExpressions.Regex]::new($regexPricesPattern)
$matchesPrice = $regexPrices.Matches($regex1) | ForEach-Object { $_ }



# Assuming dtPrices is an already initialized DataTable with the appropriate columns
# and $matchesPrice contains the price matches

foreach ($currentMatch in $matchesPrice) {
    # Replace multiple whitespace characters with a single space
    $currentMatchResults = $currentMatch.Value -replace "\s+", " "

    # Split the string into parts separated by spaces and then add as a new row to the DataTable
    $dtPrices.Rows.Add($currentMatchResults.Split(' '))
}


# Filter dtPrices directly
for ($i = $dtPrices.Rows.Count - 1; $i -ge 0; $i--) {
    $row = $dtPrices.Rows[$i]
    if ([string]::IsNullOrEmpty($row[0].ToString())) {
        $dtPrices.Rows.RemoveAt($i)
    }
}

# Create a new DataTable and copy the structure and data of dtPrices
$dtPrices1 = New-Object System.Data.DataTable
$dtPrices1 = $dtPrices.Clone()
$dtPrices1 = $dtPrices.Copy()



# Doing math for the Totals in Prices
foreach ($currentRow in $dtPrices1.Rows) {
    # Perform the calculation and division as per the original logic
    $value = [double]$currentRow[3] / $marginSelection 
    $quantity = [int]$currentRow[0]
    $result = [Math]::Round($value / $quantity, 2)

    # Prepare the new row data
    $newRowData = @(
        $result.ToString(),
        $currentRow[0].ToString(),
        [Math]::Round($value, 2).ToString()
    )

    # Add the new DataRow to dtPrices2
    $newRow = $dtPrices2.NewRow()
    $newRow.ItemArray = $newRowData
    $dtPrices2.Rows.Add($newRow)
}


# Inserting a blank row at the top of dtPrices2
$initialRow = $dtPrices2.NewRow() # Creates a new blank row
$dtPrices2.Rows.InsertAt($initialRow, 0) # Inserts the new blank row at the top (index 0)


# This inserts blank rows based on the indexing from the "Bundle Subtotal $" text
foreach ($currentItem in $indexArray1) {
    $newRow = $dtPrices2.NewRow() # Creates a new blank row
    $dtPrices2.Rows.InsertAt($newRow, $currentItem) # Inserts the new blank row at the specified index
    $dtPrices2.Rows.InsertAt($dtPrices2.NewRow(), $currentItem + 1) # Inserts another new blank row at the next index
}


# This joins the 'sitesStatesFinal' datatable with the 'dtPrices2' - and also makes their rows line up
# Assuming dtJoined3 is already set up correctly
foreach ($currentRow1 in $sitesStatesFinal.Rows) {
    foreach ($currentRow2 in $dtPrices2.Rows) {
        if ($sitesStatesFinal.Rows.IndexOf($currentRow1) -eq $dtPrices2.Rows.IndexOf($currentRow2)) {
            # Assuming both rows have the same schema and can be concatenated directly
            $joinedRow = $dtJoined3.NewRow()
            $joinedRow.ItemArray = $currentRow1.ItemArray + $currentRow2.ItemArray
            $dtJoined3.Rows.Add($joinedRow)
        }
    }
}


# Summing up the values in the "MRC" column
$mrcSUM = 0
foreach ($row in $dtJoined3.Rows) {
    $mrcValue = $row["MRC"]
    if ($null -ne $mrcValue -and $mrcValue -ne "") {
        try {
            $mrcSUM += [double]::Parse($mrcValue)
        } catch {
            Write-Host "Invalid MRC value: $mrcValue"
        }
    }
}
$mrcSUM = [Math]::Round($mrcSUM, 2).ToString()


# Adding a new row to 'dtJoined3'
$newRowForTotal = $dtJoined3.NewRow()
$newRowForTotal[3] = "Total MRC: " + '$' + $mrcSUM # Replace 3 with the actual index or column name
$dtJoined3.Rows.Add($newRowForTotal)

# Adding the first special row
$newRow = $dtJoined3.NewRow()
$newRow[3] = "Total MRC for MPP: N/A"
$dtJoined3.Rows.Add($newRow)

# Adding two entirely blank rows
for ($i = 0; $i -lt 2; $i++) {
    $blankRow = $dtJoined3.NewRow()
    $dtJoined3.Rows.Add($blankRow)
}


# Adding the row with specific data in cells [2] and [3]
$shippingRow = $dtJoined3.NewRow()
$shippingRow[2] = "Shipping Costs of AT&T Equipment, One Time Charge - (OTC)"
$shippingRow[3] = '$' + $price
$dtJoined3.Rows.Add($shippingRow)


# Replace the below line with your actual initial row data
$initialRowData_final = @(" ", " ", " ", " ") # Adjust as per your initial data requirements
$initialRow_final = $dtJoined3.NewRow()
$initialRow_final.ItemArray = $initialRowData_final
$dtJoined3.Rows.InsertAt($initialRow_final, 0)







## Building and modifying dtSKU DataTable
foreach ($currentMatch in $matchSKU) {
    $row = $dtSKU.NewRow()
    # Apply the regex replacement directly here
    $modifiedValue = [regex]::Replace($currentMatch.Value, "\d+\s*\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}", "")
    $row["SKU Column"] = $modifiedValue
    $dtSKU.Rows.Add($row)
}

$dtSKU.Rows.RemoveAt(0)


# Create a list to store the row indices to remove
$indicesToRemove = @()

# Remove rows after indices in $indexArray1
foreach ($index in $indexArray1) {
    if ($index + 1 -lt $sitesStatesFinal.Rows.Count) {
        $indicesToRemove += $index + 1
    }
}


# Find and remove blank rows (including rows with only spaces)
for ($i = 0; $i -lt $sitesStatesFinal.Rows.Count; $i++) {
    $row = $sitesStatesFinal.Rows[$i]
    $isEmpty = $true

    foreach ($cell in $row.ItemArray) {
        if (![string]::IsNullOrEmpty($cell.ToString().Trim())) {
            $isEmpty = $false
            break
        }
    }

    if ($isEmpty) {
        $indicesToRemove += $i
    }
}

# Remove the first row
$indicesToRemove += 0

# Create a new DataTable with filtered rows
$sitesStatesSKU = $sitesStatesFinal.Clone()
for ($i = 0; $i -lt $sitesStatesFinal.Rows.Count; $i++) {
    if (-not $indicesToRemove.Contains($i)) {
        $newRow = $sitesStatesSKU.NewRow()
        $newRow.ItemArray = $sitesStatesFinal.Rows[$i].ItemArray
        $sitesStatesSKU.Rows.Add($newRow)
    }
}





$dtSKU2 = New-Object System.Data.DataTable
$dtSKU2.Columns.Add("Product ID", [string])
$dtSKU2.Columns.Add("Size/Capacity & Other Details", [string])
$dtSKU2.Columns.Add("# Units", [int])
$dtSKU2.Columns.Add("Site Address", [string])

# Assuming the number of rows in each DataTable is the same
for ($i = 0; $i -lt $dtSKU.Rows.Count; $i++) {
    $newRow = $dtSKU2.NewRow()

    # Add "Product ID" from dtSKU
    $newRow["Product ID"] = $dtSKU.Rows[$i][0]

    # Add "Size/Capacity & Other Details" from sitesStatesFiltered
    if ($i -lt $sitesStatesSKU.Rows.Count) {
        $newRow["Size/Capacity & Other Details"] = $sitesStatesSKU.Rows[$i][0]
    }

    # Add "# Units" from dtPrices1 (assuming the middle column is at index 2)
    if ($i -lt $dtPrices1.Rows.Count) {
        $newRow["# Units"] = $dtPrices1.Rows[$i][0] -as [int]
    }

    # Leave "Site Address" blank for now
    $newRow["Site Address"] = ""

    $dtSKU2.Rows.Add($newRow)
}







#TESTING - DELETE THESE
# Assuming $dataTable is your DataTable
$rowCount = $sitesStatesDT.Rows.Count

# Display the row count
Write-Host "sitesStatesDT has $rowCount rows."




#REGEX FOR VARIABLES

# Customer Variables
$customerName = $coverPage_Table.Rows[0][0].ToString()

# Customer Address
$customerAddress = $coverPage_Table.Rows[2][0].ToString()

# Customer Street
$customerStreet_edit = [System.Text.RegularExpressions.Regex]::Match($customerAddress, "Address:(.*?)City").Value.Trim()
$customerStreet = [System.Text.RegularExpressions.Regex]::Replace($customerStreet_edit, "Address:|City", "")

# Customer City
$customerCity_edit = [System.Text.RegularExpressions.Regex]::Match($customerAddress, "City:(.*?)State").Value.Trim()
$customerCity = [System.Text.RegularExpressions.Regex]::Replace($customerCity_edit, "City:|State", "")

# Customer State
$customerState = [System.Text.RegularExpressions.Regex]::Match($customerAddress, "AL|AK|AZ|AR|CA|CO|CT|DE|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY").Value.Trim()

# Customer Zip Code
$customerZip = [System.Text.RegularExpressions.Regex]::Match($customerAddress, "\d{5}(-\d{4})?").Value.Trim()

# Customer Contact
$customerContact = $coverPage_Table.Rows[4][0].ToString()

# Customer Contact Name
$customerContactName_edit = [System.Text.RegularExpressions.Regex]::Match($customerContact, "Name:(.*?)Title").Value.Trim()
$customerContactName = [System.Text.RegularExpressions.Regex]::Replace($customerContactName_edit, "Name:|Title", "")

# Customer Title
$customerTitle_edit = [System.Text.RegularExpressions.Regex]::Match($customerContact, "Title:(.*?)Telephone").Value.Trim()
$customerTitle = [System.Text.RegularExpressions.Regex]::Replace($customerTitle_edit, "Title:|Telephone", "")

# Customer Phone
$customerPhone_edit = [System.Text.RegularExpressions.Regex]::Match($customerContact, "Telephone:(.*?)Fax").Value.Trim()
$customerPhone = [System.Text.RegularExpressions.Regex]::Replace($customerPhone_edit, "Telephone:|Fax", "")

# Customer Email
$customerEmail = [System.Text.RegularExpressions.Regex]::Match($customerContact, "(?("")(""[^""]+?""@)|(([0-9a-zA-Z]((\.(?!\.))|[-!#$%&'*+/=?^_`{|}~\w])*)(?<=[0-9a-zA-Z])@))" + "((\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-zA-Z][-0-9a-zA-Z]*[0-9a-zA-Z]*\.)+[a-zA-Z0-9][\-a-zA-Z0-9]{0,22}[a-zA-Z0-9]))").Value.Trim()

# ...and so on for the remaining variables




# Load Word COM object
$word = New-Object -ComObject Word.Application
$templateDoc = $word.Documents.Open($contractTemplate) # Update the path
$word.Visible = $true

# Placeholder text to find
$findText = "<<customer contact>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerContactName
}


# Placeholder text to find
$findText = "<<customer email>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerEmail
}


# Placeholder text to find
$findText = "<<customer title>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerTitle
}

# Convert the array of integers into a string
$indexArrayString = $indexArray1 -join ", "

# Placeholder text to find
$findText = "<<indexArray1>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the string representation of the variable content
    $textRange.Text = $indexArrayString
}




# Placeholder text
$findText = "<<pricingTable>>"
$find = $templateDoc.Content.Find
$find.ClearFormatting()

if ($find.Execute($findText)) {
    $dataTableRange = $find.Parent
    $dataTableRange.Select()

    $rowCount = $dtJoined3.Rows.Count
    $columnCount = $dtJoined3.Columns.Count
    $wordTable1 = $templateDoc.Tables.Add($dataTableRange, $rowCount + 1, $columnCount)

    # Center the entire table horizontally
    $wordTable1.Rows.Alignment = [Microsoft.Office.Interop.Word.WdRowAlignment]::wdAlignRowCenter

    # Set the cell alignment to center
    foreach ($cell in $wordTable1.Range.Cells) {
        $cell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
        $cell.VerticalAlignment = [Microsoft.Office.Interop.Word.WdCellVerticalAlignment]::wdCellAlignVerticalCenter

        # Set the font and size
        $cell.Range.Font.Name = "Arial"
        $cell.Range.Font.Size = 8

        # Adjust paragraph spacing
        $cell.Range.ParagraphFormat.SpaceBefore = 0
        $cell.Range.ParagraphFormat.SpaceAfter = 0

        # Set line spacing to single
        $cell.Range.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpaceSingle
        }
    }

# Add column headers
for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
    $headerText = [System.Convert]::ToString($dtJoined3.Columns[$columnIndex].ColumnName)
    $wordTable1.Cell(1, $columnIndex + 1).Range.Text = $headerText
}

# Add data rows
for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
    for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
        if ($wordTable1.Cell($rowIndex + 2, $columnIndex + 1)) {
            $cellData = $dtJoined3.Rows[$rowIndex][$columnIndex] -as [String]

            # Check if the column is a currency column and the cell contains a number
            if ($currencyColumnIndices -contains $columnIndex -and $cellData -match '^\d+(\.\d+)?$') {
                # Format as currency (with a dollar sign)
                $cellData = ('${0:N2}' -f [double]$cellData)
            }

            $wordTable1.Cell($rowIndex + 2, $columnIndex + 1).Range.Text = $cellData
        }
    }
}


    # Set the table's border style
    $borders = @(
        $wordTable1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft),
        $wordTable1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight),
        $wordTable1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderTop),
        $wordTable1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom),
        $wordTable1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderHorizontal),
        $wordTable1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical)
    )

    foreach ($border in $borders) {
        $border.LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleSingle
        $border.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlack
    }

    # Modifying the borders for the last 5 rows
    $startRow = $wordTable1.Rows.Count - 4

    for ($i = $startRow; $i -le $wordTable1.Rows.Count; $i++) {
    # Cells [0] and [1] in each of these rows
    $cell1 = $wordTable1.Cell($i, 1)
    $cell2 = $wordTable1.Cell($i, 2)

    # Removing the bottom and inner borders for these cells
    $cell1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    $cell1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    $cell2.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    $cell2.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
}



# Placeholder text for SKU table
$findTextSKU = "<<skuTable>>"
$findSKU = $templateDoc.Content.Find
$findSKU.ClearFormatting()

if ($findSKU.Execute($findTextSKU)) {
    $dataTableRangeSKU = $findSKU.Parent
    $dataTableRangeSKU.Select()

    $rowCountSKU = $dtSKU2.Rows.Count
    $columnCountSKU = $dtSKU2.Columns.Count
    $wordTableSKU = $templateDoc.Tables.Add($dataTableRangeSKU, $rowCountSKU + 1, $columnCountSKU)

    # Center the entire table horizontally
    $wordTableSKU.Rows.Alignment = [Microsoft.Office.Interop.Word.WdRowAlignment]::wdAlignRowCenter

    # Set the cell alignment to center
    foreach ($cellSKU in $wordTableSKU.Range.Cells) {
        $cellSKU.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
        $cellSKU.VerticalAlignment = [Microsoft.Office.Interop.Word.WdCellVerticalAlignment]::wdCellAlignVerticalCenter

        # Set the font and size
        $cellSKU.Range.Font.Name = "Arial"
        $cellSKU.Range.Font.Size = 8
    }

    # Add column headers for SKU table
    for ($columnIndexSKU = 0; $columnIndexSKU -lt $columnCountSKU; $columnIndexSKU++) {
        $headerTextSKU = [System.Convert]::ToString($dtSKU2.Columns[$columnIndexSKU].ColumnName)
        $wordTableSKU.Cell(1, $columnIndexSKU + 1).Range.Text = $headerTextSKU
    }

    # Add data rows for SKU table
    for ($rowIndexSKU = 0; $rowIndexSKU -lt $rowCountSKU; $rowIndexSKU++) {
        for ($columnIndexSKU = 0; $columnIndexSKU -lt $columnCountSKU; $columnIndexSKU++) {
            if ($wordTableSKU.Cell($rowIndexSKU + 2, $columnIndexSKU + 1)) {
                $cellDataSKU = $dtSKU2.Rows[$rowIndexSKU][$columnIndexSKU] -as [String]
                $wordTableSKU.Cell($rowIndexSKU + 2, $columnIndexSKU + 1).Range.Text = $cellDataSKU
            }
        }
    }

    # Set the table's border style
    $borders = @(
        $wordTableSKU.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft),
        $wordTableSKU.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight),
        $wordTableSKU.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderTop),
        $wordTableSKU.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom),
        $wordTableSKU.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderHorizontal),
        $wordTableSKU.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical)
    )

    foreach ($border in $borders) {
        $border.LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleSingle
        $border.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlack
    }
}





$findTextSKU1 = "<<customerInfoDT>>"
$findSKU1 = $templateDoc.Content.Find
$findSKU1.ClearFormatting()

if ($findSKU1.Execute($findTextSKU1)) {
    $dataTableRangeSKU1 = $findSKU1.Parent
    $dataTableRangeSKU1.Select()

    $rowCountSKU1 = $customerInfoDT.Rows.Count
    $columnCountSKU1 = $customerInfoDT.Columns.Count
    $wordTableSKU1 = $templateDoc.Tables.Add($dataTableRangeSKU1, $rowCountSKU1 + 1, $columnCountSKU1)

    # Center the entire table horizontally
    $wordTableSKU1.Rows.Alignment = [Microsoft.Office.Interop.Word.WdRowAlignment]::wdAlignRowCenter

    # Set the cell alignment to center
    foreach ($cellSKU1 in $wordTableSKU1.Range.Cells) {
        $cellSKU1.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
        $cellSKU1.VerticalAlignment = [Microsoft.Office.Interop.Word.WdCellVerticalAlignment]::wdCellAlignVerticalCenter

        # Set the font and size
        $cellSKU1.Range.Font.Name = "Arial"
        $cellSKU1.Range.Font.Size = 8
    }

    # Add column headers for SKU table
    for ($columnIndexSKU1 = 0; $columnIndexSKU1 -lt $columnCountSKU1; $columnIndexSKU1++) {
        $headerTextSKU1 = [System.Convert]::ToString($customerInfoDT.Columns[$columnIndexSKU1].ColumnName)
        $wordTableSKU1.Cell(1, $columnIndexSKU1 + 1).Range.Text = $headerTextSKU1
    }

    # Add data rows for SKU table
    for ($rowIndexSKU1 = 0; $rowIndexSKU1 -lt $rowCountSKU1; $rowIndexSKU1++) {
        for ($columnIndexSKU1 = 0; $columnIndexSKU1 -lt $columnCountSKU1; $columnIndexSKU1++) {
            if ($wordTableSKU1.Cell($rowIndexSKU1 + 2, $columnIndexSKU1 + 1)) {
                $cellDataSKU1 = $customerInfoDT.Rows[$rowIndexSKU1][$columnIndexSKU1] -as [String]
                $wordTableSKU1.Cell($rowIndexSKU1 + 2, $columnIndexSKU1 + 1).Range.Text = $cellDataSKU1
            }
        }
    }

    # Set the table's border style
    $borders = @(
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderTop),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderHorizontal),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical)
    )

    foreach ($border in $borders) {
        $border.LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleSingle
        $border.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlack
    }
}



#Placeholder text for SKU table

$findTextSKU1 = "<<sitesStatesSKU>>"
$findSKU1 = $templateDoc.Content.Find
$findSKU1.ClearFormatting()

if ($findSKU1.Execute($findTextSKU1)) {
    $dataTableRangeSKU1 = $findSKU1.Parent
    $dataTableRangeSKU1.Select()

    $rowCountSKU1 = $sitesStatesSKU.Rows.Count
    $columnCountSKU1 = $sitesStatesSKU.Columns.Count
    $wordTableSKU1 = $templateDoc.Tables.Add($dataTableRangeSKU1, $rowCountSKU1 + 1, $columnCountSKU1)

    # Center the entire table horizontally
    $wordTableSKU1.Rows.Alignment = [Microsoft.Office.Interop.Word.WdRowAlignment]::wdAlignRowCenter

    # Set the cell alignment to center
    foreach ($cellSKU1 in $wordTableSKU1.Range.Cells) {
        $cellSKU1.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
        $cellSKU1.VerticalAlignment = [Microsoft.Office.Interop.Word.WdCellVerticalAlignment]::wdCellAlignVerticalCenter

        # Set the font and size
        $cellSKU1.Range.Font.Name = "Arial"
        $cellSKU1.Range.Font.Size = 8
    }

    # Add column headers for SKU table
    for ($columnIndexSKU1 = 0; $columnIndexSKU1 -lt $columnCountSKU1; $columnIndexSKU1++) {
        $headerTextSKU1 = [System.Convert]::ToString($sitesStatesSKU.Columns[$columnIndexSKU1].ColumnName)
        $wordTableSKU1.Cell(1, $columnIndexSKU1 + 1).Range.Text = $headerTextSKU1
    }

    # Add data rows for SKU table
    for ($rowIndexSKU1 = 0; $rowIndexSKU1 -lt $rowCountSKU1; $rowIndexSKU1++) {
        for ($columnIndexSKU1 = 0; $columnIndexSKU1 -lt $columnCountSKU1; $columnIndexSKU1++) {
            if ($wordTableSKU1.Cell($rowIndexSKU1 + 2, $columnIndexSKU1 + 1)) {
                $cellDataSKU1 = $sitesStatesSKU.Rows[$rowIndexSKU1][$columnIndexSKU1] -as [String]
                $wordTableSKU1.Cell($rowIndexSKU1 + 2, $columnIndexSKU1 + 1).Range.Text = $cellDataSKU1
            }
        }
    }

    # Set the table's border style
    $borders = @(
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderTop),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderHorizontal),
        $wordTableSKU1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical)
    )

    foreach ($border in $borders) {
        $border.LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleSingle
        $border.Color = [Microsoft.Office.Interop.Word.WdColor]::wdColorBlack
    }
}






# Assuming $word is already a Word application object and $templateDoc is the document

#INSERTING THE COVER PAGE TABLE AND FORMATTING IT
$placeholderRange = $templateDoc.Content
$findText = "<<coverPage>>"

$find = $placeholderRange.Find
$find.ClearFormatting()

if ($find.Execute($findText)) {
    # Set the range to the end of the found text
    $placeholderRange = $find.Parent
    $placeholderRange.Collapse([Microsoft.Office.Interop.Word.WdCollapseDirection]::wdCollapseEnd)

    # Paste the table at the new range position
    $placeholderRange.Paste()

    $wordTable = $templateDoc.Tables[1]
    $placeholderRange.Font.Name = "Arial"
    $placeholderRange.Font.Size = 8
    $placeholderRange.HighlightColorIndex = [Microsoft.Office.Interop.Word.WdColorIndex]::wdNoHighlight

    # Specific cells to format
    $specificCells = @(
        [Tuple]::Create(2,1), [Tuple]::Create(2,2), [Tuple]::Create(2,3),
        [Tuple]::Create(4,1), [Tuple]::Create(4,2), [Tuple]::Create(4,3),
        [Tuple]::Create(6,1), [Tuple]::Create(6,2), [Tuple]::Create(6,3),
        [Tuple]::Create(8,1), [Tuple]::Create(8,2), [Tuple]::Create(8,3)
    )

    # State abbreviations
    $stateAbbreviations = @("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", "USA", "US")

    foreach ($row in $wordTable.Rows) {
        foreach ($cell in $row.Cells) {
            $cellTuple = [Tuple]::Create($row.Index, $cell.ColumnIndex)

            if ($specificCells -contains $cellTuple) {
                $text = $cell.Range.Text.Trim()
                $text = $text -replace ": ", ":"
                $text = $text -replace ":", ": "

                $words = $text -split ' '
                for ($i = 0; $i -lt $words.Length; $i++) {
                    $upperWord = $words[$i].ToUpper()

                    if ($stateAbbreviations -contains $upperWord) {
                        $words[$i] = $upperWord
                    }
                    elseif ($words[$i] -match "@") {
                        $words[$i] = $words[$i].ToLower()
                    }
                    else {
                        $words[$i] = [Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($words[$i].ToLower())
                    }
                }

                $newText = $words -join " "
                foreach ($abbreviation in $stateAbbreviations) {
                    $pattern = "\b$abbreviation\b"
                    $newText = [Regex]::Replace($newText, $pattern, $abbreviation.ToUpper(), [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                }

                $cell.Range.Text = $newText
            }
        }
    }
}





$newFilePath = $contractTemplateNew

# Save and close the document
$templateDoc.SaveAs([ref] $newFilePath)
$templateDoc.Close()

# Cleanup COM object
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($templateDoc) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()


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
$startRowIndex = 15

# Iterate through each row in the DataTable
for ($i = 0; $i -lt $dtJoined3.Rows.Count; $i++) {
    # Current row in the DataTable
    $row = $dtJoined3.Rows[$i]

    # Write data to Excel
    $worksheet.Cells.Item($startRowIndex + $i, 1) = $row[0].ToString()  # Column A
    $worksheet.Cells.Item($startRowIndex + $i, 2) = $row[1].ToString()  # Column B
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
for ($i = 0; $i -lt $sitesStatesFinal.Rows.Count; $i++) {
    $worksheetXLS.Cells.Item($i + 1, 1) = $sitesStatesFinal.Rows[$i][0].ToString()
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




