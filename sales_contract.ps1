Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data.DataSetExtensions
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName System.Data

# Set the absolute path to your config.ps1 script
$configPath = "D:\Programming\PowerShell\Sales Contract\config.ps1"
# Dot source the config script
. $configPath



# Show a message box with the desired message
[Windows.Forms.MessageBox]::Show('SELECT THE PDF QUOTE', '')

function Select-PdfFile {
    $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $fileDialog.InitialDirectory = [Environment]::GetFolderPath('Desktop')
    $fileDialog.Filter = 'PDF files (*.pdf)|*.pdf'
    $fileDialog.ShowDialog() | Out-Null
    return $fileDialog.FileName
}



$pdfToTextPath = "C:\Program Files\xpdf-tools-win-4.04\xpdf-tools-win-4.04\bin64\pdftotext.exe"
$pdfFilePath = Select-PdfFile


if ([string]::IsNullOrWhiteSpace($pdfFilePath)) {
    Write-Host "No file selected."
} else {
    $outputTxtPath = [System.IO.Path]::ChangeExtension($pdfFilePath, '.txt')
    & $pdfToTextPath -table $pdfFilePath $outputTxtPath

    if (Test-Path $outputTxtPath) {
        $textContent = Get-Content $outputTxtPath -Raw

        #Write-Host $textContent


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



# Add necessary assembly for Outlook and Windows Forms
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
Add-Type -AssemblyName System.Windows.Forms

# Initialize Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")

function Get-GALUserDetails {
    param([string]$emailAddress)

    $Recipient = $namespace.CreateRecipient($emailAddress)
    $Recipient.Resolve()

    if ($Recipient.Resolved -and $Recipient.AddressEntry.GetExchangeUser()) {
        $ExchangeUser = $Recipient.AddressEntry.GetExchangeUser()
        $Manager = $ExchangeUser.GetExchangeUserManager()

        if ($Manager) {
            $ManagerName = $Manager.Name
        }

        $Details = New-Object PSObject -Property @{
            Name = $ExchangeUser.Name
            JobTitle = $ExchangeUser.JobTitle
            BusinessAddress = $ExchangeUser.StreetAddress
            BusinessCity = $ExchangeUser.City
            BusinessState = $ExchangeUser.StateOrProvince
            BusinessZip = $ExchangeUser.PostalCode
            BusinessPhone = $ExchangeUser.BusinessTelephoneNumber
            ManagerName = $ManagerName
        }

        return $Details
    } else {
        Write-Warning "Could not resolve $emailAddress in Global Address List."
        return $null
    }
}

# Modified input dialog with Skip option
function Show-InputDialogWithSkip {
    param([string]$Message, [string]$WindowTitle = "Input")

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $WindowTitle
    $form.Size = New-Object System.Drawing.Size(300,200)
    $form.StartPosition = "CenterScreen"

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.Location = New-Object System.Drawing.Point(10,20)
    $label.Size = New-Object System.Drawing.Size(280,20)

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10,40)
    $textBox.Size = New-Object System.Drawing.Size(260,20)

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Location = New-Object System.Drawing.Point(10,70)
    $okButton.Size = New-Object System.Drawing.Size(75,23)
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $skipButton = New-Object System.Windows.Forms.Button
    $skipButton.Location = New-Object System.Drawing.Point(195,70)
    $skipButton.Size = New-Object System.Drawing.Size(75,23)
    $skipButton.Text = "Skip"
    $skipButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $form.AcceptButton = $okButton
    $form.CancelButton = $skipButton

    $form.Controls.Add($label)
    $form.Controls.Add($textBox)
    $form.Controls.Add($okButton)
    $form.Controls.Add($skipButton)

    $form.Topmost = $true

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $textBox.Text
    } elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel) {
        return $null
    }
}

# Main loop
do {
    $emailAddress = Show-InputDialogWithSkip -Message "Enter the User ID (Email Address) of the person:" -WindowTitle "User ID Input"
    if ($null -eq $emailAddress) {
        Write-Output "User skipped input."
        break
    }

    $UserDetails = Get-GALUserDetails -emailAddress $emailAddress

    if ($UserDetails) {
        [System.Windows.Forms.MessageBox]::Show("Name: $($UserDetails.Name)`nJob Title: $($UserDetails.JobTitle)`nBusiness Address: $($UserDetails.BusinessAddress)`nBusiness Phone: $($UserDetails.BusinessPhone)`nManager Name: $($UserDetails.ManagerName)", "User Details")
        $retry = $false
    } else {
        $retry = $true
        [System.Windows.Forms.MessageBox]::Show("No details found for $emailAddress. Would you like to retry?", "Error", [System.Windows.Forms.MessageBoxButtons]::RetryCancel) -eq [System.Windows.Forms.DialogResult]::Retry
    }
} while ($retry)


$salesName = $Details.Name
$salesJobTitle = $Details.JobTitle
$salesStreetAddress = $Details.BusinessAddress
$salesCity = $Details.BusinessCity
$salesState = $Details.BusinessState
$salesZip = $Details.BusinessZip
$salesPhone = $Details.BusinessPhone
$salesManagerName = $Details.ManagerName


function Convert-NameFormat {
    param([string]$name)

    if ($name -contains ',') {
        $parts = $name -split ','
        $formattedName = "$($parts[1].Trim()) $($parts[0].Trim())"
    } else {
        $formattedName = $name
    }

    return $formattedName
}

# Assuming you have $Name and $ManagerName variables already populated
$salesName = Convert-NameFormat -name $salesName
$salesManagerName = Convert-NameFormat -name $salesManagerName





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



#Variables from customer contact regex extraction

$customerName = $customerInfoDT.Rows[0][0]
$customerContactName = $customerInfoDT.Rows[3][0]
$customerEmail = $customerInfoDT.Rows[5][0]
$customerStreet = $customerInfoDT.Rows[1][0]
$customerCityStateZip = $customerInfoDT.Rows[2][0]
$customerPhone = $customerInfoDT.Rows[4][0]


$customerCity = $customerCityStateZip -replace ', [A-Z]{2} \d{5}(-\d{4})?|[A-Z]{2} \d{5}(-\d{4})?', ""


# Attempt to match state abbreviation
if ($customerCityStateZip -match ', ([A-Z]{2}) ') {
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




# Load Word COM object
$word = New-Object -ComObject Word.Application
$templateDoc = $word.Documents.Open($contractTemplate) # Update the path
$word.Visible = $true


# Placeholder text to find
$findText = "<<customer name>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerName
}

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
$findText = "<<customer phone>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerPhone
}

# Placeholder text to find
$findText = "<<customer street>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerStreet
}


# Placeholder text to find
$findText = "<<customer city>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerCity
}


# Placeholder text to find
$findText = "<<customer state>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerState
}


# Placeholder text to find
$findText = "<<customer zip>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $customerZip
}


# Placeholder text to find
$findText = "<<sales name>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $salesName
}

# Placeholder text to find
$findText = "<<sales street>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $salesStreetAddress
}

# Placeholder text to find
$findText = "<<sales city>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $salesCity
}

# Placeholder text to find
$findText = "<<sales state>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $salesState
}

# Placeholder text to find
$findText = "<<sales zip>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $salesZip
}

# Placeholder text to find
$findText = "<<sales manager>>"

# Access the Find object
$find = $templateDoc.Content.Find
$find.ClearFormatting()

# Check if the placeholder text is found in the document
if ($find.Execute($findText)) {
    # Get the range where the text was found
    $textRange = $find.Parent

    # Replace the found text with the variable content
    $textRange.Text = $salesManagerName
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

    $startRow = $wordTable1.Rows.Count - 4

for ($i = $startRow; $i -le $wordTable1.Rows.Count; $i++) {
    # Cells [0] and [1] in each of these rows
    $cell1 = $wordTable1.Cell($i, 1)
    $cell2 = $wordTable1.Cell($i, 2)

    # Removing the bottom border for these cells
    $cell1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    $cell2.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    
    # Correctly removing the "inner vertical border" between $cell1 and $cell2
    $cell1.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
}


    

# Remove internal vertical borders from the first row
for ($i = 2; $i -lt $wordTable1.Columns.Count; $i++) {
    $wordTable1.Cell(1, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
}

# Remove bottom borders from the first row, starting from the 2nd cell
for ($i = 2; $i -le $wordTable1.Columns.Count; $i++) {
    $wordTable1.Cell(2, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
}

# Remove internal vertical borders from the second row, starting at the 2nd cell
for ($i = 1; $i -lt $wordTable1.Columns.Count; $i++) {
    $wordTable1.Cell(2, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
}

# Remove internal vertical borders from the second row, starting at the 2nd cell
for ($i = 2; $i -lt $wordTable1.Columns.Count; $i++) {
    $wordTable1.Cell(3, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
}





foreach ($index in $indexArray1) {
    $rowIndex = $index + 3 # Adjusting each index as specified (+1)
    
    # Ensure the row index is within the table's bounds
    if ($rowIndex -le $wordTable1.Rows.Count) {
        # Loop through all but the last cell in the specified row to remove inner vertical borders
        for ($i = 1; $i -lt $wordTable1.Columns.Count; $i++) {
            $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        }
    }
}

foreach ($index in $indexArray1) {
    $rowIndex = $index + 3 # Adjusting each index as specified (+1)
    
    # Ensure the row index is within the table's bounds
    if ($rowIndex -le $wordTable1.Rows.Count) {
        # Loop through all but the last cell in the specified row to remove inner vertical borders
        for ($i = 2; $i -le $wordTable1.Columns.Count; $i++) {
            $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        }
    }
}

foreach ($index in $indexArray1) {
    $rowIndex = $index + 4 # Adjusting each index as specified (+1)
    
    # Ensure the row index is within the table's bounds
    if ($rowIndex -le $wordTable1.Rows.Count) {
        # Loop through all but the last cell in the specified row to remove inner vertical borders
        for ($i = 2; $i -lt $wordTable1.Columns.Count; $i++) {
            $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        }
    }
}


# Access the first cell in the second row
$cell = $wordTable1.Cell(3, 1)

# Make the text bold and set font size to 9
$cell.Range.Font.Bold = $true
$cell.Range.Font.Size = 9



# Iterate over each index in indexArray1 except for the last one
for ($j = 0; $j -lt $indexArray1.Count - 1; $j++) {
    $index = $indexArray1[$j]
    $rowIndex = $index + 4 # Adjusting each index as specified

    # Ensure the row index is within the table's bounds
    if ($rowIndex -le $wordTable1.Rows.Count) {
        # Access the first cell in the specified row
        $cell = $wordTable1.Cell($rowIndex, 1)
        
        # Make the text bold and set font size to 9
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Size = 9
    }
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


#Placeholder text for SKU table
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




# Formatting and beautifying the cover page
$wordTable = $templateDoc.Tables[1]

# Specific cells to format
$specificCells = @(
    [Tuple]::Create(2,1), [Tuple]::Create(2,2), 
    [Tuple]::Create(4,1), [Tuple]::Create(4,2), 
    [Tuple]::Create(6,1), [Tuple]::Create(6,2), 
    [Tuple]::Create(8,1), [Tuple]::Create(8,2)
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
            # Resetting cell formatting to remove any residual highlighting or font changes
            $cell.Range.Font.Name = "Arial"
            $cell.Range.Font.Size = 8
            $cell.Range.HighlightColorIndex = [Microsoft.Office.Interop.Word.WdColorIndex]::wdNoHighlight
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




