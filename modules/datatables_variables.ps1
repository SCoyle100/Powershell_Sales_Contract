#Refactor notes - for now, I would just declare the variables up here
#based off of the methods from the classes in forms, and then work from there. 
#I'm guessing the classes in this module will take the declared variables.

$regex0 = [RegexOperations]::ExtractQuotation($pdfText)
$regex1 = [RegexOperations]::ExtractItemDescription($pdfText)
$regex2 = [RegexOperations]::RemovePricingDetails($regex1)
$tenure = [RegexOperations]::ExtractPaymentTenure($pdfText)
$shippingInfo = [RegexOperations]::ExtractShippingCost($pdfText)








# Line by line capture for descriptions, quantity, and prices


$customerInfo = [regex]::Matches($regex0, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value

#$sitesStates = [regex]::Matches($regex1, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value
# Print the $sitesStates variable to the terminal
#Write-Host "Sites States:"
#$sitesStates | ForEach-Object { Write-Host $_ }

# Line by Line capture to build datatable with descriptions only
$sitesStatesRegex2 = [regex]::Matches($regex2, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value
# Print the $sitesStatesRegex2 variable to the terminal
#Write-Host "`nSites States Regex2:"
#$sitesStatesRegex2 | ForEach-Object { Write-Host $_ }

class DataTableManager {
    static [System.Data.DataTable] CreateCustomerInfoDT() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Column1", [string])
        return $dt
    }

    static [System.Data.DataTable] CreateSitesStatesDT() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Column1", [string])
        return $dt
    }

    static [System.Data.DataTable] CreatePricesDT() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Qty", [decimal])
        $dt.Columns.Add("List Price", [decimal])
        $dt.Columns.Add("Total Price", [decimal])
        $dt.Columns.Add("MRC", [decimal])
        return $dt
    }

    # Add more methods for other DataTables as needed
}



class DataTableOperations {
    static [void] AddRowToDataTable([System.Data.DataTable] $dt, [string] $data) {
        $dt.Rows.Add($data)
    }

    static [System.Data.DataTable] FilterDataTable([System.Data.DataTable] $dt, [string] $filterCriteria) {
        $filteredDT = $dt.Clone()
        foreach ($row in $dt.Rows) {
            if ($row["Column1"] -like $filterCriteria) {
                $filteredDT.ImportRow($row)
            }
        }
        return $filteredDT
    }

    # Add more methods for other operations as needed
}




# Creating Data Tables

#stores customer information from the sales quote 
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
$dtJoined3.Columns.Add("Site Number", [string])
$dtJoined3.Columns.Add("Monthly Recurring Charges (MRC) Per Unit", [string])
$dtJoined3.Columns.Add("Units", [string])
$dtJoined3.Columns.Add("Extended MRC", [string])

$dtPrices2 = New-Object System.Data.DataTable
#$dtPrices2.Columns.Add("Site Number", [string])
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
    $pattern = "(Quote\s+No.*|Quote\s+Date.*|Valid\s+Until.*|Payment\s+Term.*)|Quoted|$telcoSales|$telcoStreet|$telcoCity|$telcoStateZip|$telcoPhone|$telcoEmail|$telcoName"

    # Replace matched patterns with an empty string, effectively removing them
    $updatedText = $currentText -replace $pattern, ""

    # Update the row's text with the modified value
    $row["Column1"] = $updatedText.Trim() # .Trim() is used to remove any leading or trailing whitespace that might be left
}





#building sitesStatesDT

#TODO - Use simliar approach you used for customerInfoDT so we don't have to an extra step of datatable filtering
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

# Add a new column named "Site Number" to dtPrices2
$siteNumberColumn = New-Object System.Data.DataColumn "Site Number", ([string])
$dtPrices2.Columns.Add($siteNumberColumn)
$siteNumberColumn.SetOrdinal(0) # Move the "Site Number" column to be the first column

# Fill all cells in the "Site Number" column with the number "1"
foreach ($row in $dtPrices2.Rows) {
    $row["Site Number"] = "1"
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
    $mrcValue = $row["Extended MRC"]
    if ($null -ne $mrcValue -and $mrcValue -ne "") {
        try {
            $mrcSUM += [double]::Parse($mrcValue)
        } catch {
            Write-Host "Invalid MRC value: $mrcValue"
        }
    }
}
$mrcSUM = [Math]::Round($mrcSUM, 2).ToString()


# Adding bottom row for Total MRC to the pricing table (aka dtJoined3)
$newRowForTotal = $dtJoined3.NewRow()
$newRowForTotal[3] = "Total MRC:"
$newRowForTotal[4] = '$' + $mrcSUM
$dtJoined3.Rows.Add($newRowForTotal)

# Adding bottom row for Total MRC for MPP to the pricing table (aka dtJoined3)
$newRow = $dtJoined3.NewRow()
$newRow[3] = "Total MRC for MPP:"
$newRow[4] = "N/A"
$dtJoined3.Rows.Add($newRow)

# Adding two entirely blank rows for aesthetic
for ($i = 0; $i -lt 2; $i++) {
    $blankRow = $dtJoined3.NewRow()
    $dtJoined3.Rows.Add($blankRow)
}


# bottom row for shipping to the pricing table (aka dtJoined3)
$shippingRow = $dtJoined3.NewRow()
$shippingRow[3] = "Shipping Costs of AT&T Equipment, One Time Charge - (OTC)"
$shippingRow[4] = '$' + $price
$dtJoined3.Rows.Add($shippingRow)


# Replace the below line with your actual initial row data
$initialRowData_final = @(" ", " ", " ", " ") # Adjust as per your initial data requirements
$initialRow_final = $dtJoined3.NewRow()
$initialRow_final.ItemArray = $initialRowData_final
$dtJoined3.Rows.InsertAt($initialRow_final, 0)


foreach ($row in $dtJoined3.Rows) {
    if ($row[0] -like "*Total*") {
        $dtJoined3.Rows.Remove($row)
        
        break
    }
}

# Accept changes
$dtJoined3.AcceptChanges()










### Building and modifying dtSKU2 DataTable ###


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



#Variables from customer contact regex extraction

$customerName = $customerInfoDT.Rows[0][0]
$customerContactName = $customerInfoDT.Rows[3][0]
$customerEmail = $customerInfoDT.Rows[5][0]
$customerStreet = $customerInfoDT.Rows[1][0]
$customerCityStateZip = $customerInfoDT.Rows[2][0]
$customerPhone = $customerInfoDT.Rows[4][0]


$customerCity = $customerCityStateZip -replace '[A-Z]{2}|\d{5}(-\d{4})?|[A-Z]{2} \d{5}(-\d{4})?', ""


# Attempt to match state abbreviation
if ($customerCityStateZip -match '([A-Z]{2})') {
    $customerState = $matches[1] # Extract the state abbreviation
} else {
    $customerState = $null # No match found
}

# Attempt to match zip code (standard or ZIP+4)
if ($customerCityStateZip -match '(\d{5}(-\d{4})?)') {
    $customerZip = $matches[1] # Extract the zip code
} else {
    $customerZip = $null # No match found
}
