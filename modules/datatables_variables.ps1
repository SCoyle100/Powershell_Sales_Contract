#Refactor notes - for now, I would just declare the variables up here
#based off of the methods from the classes in forms, and then work from there. 
#I'm guessing the classes in this module will take the declared variables.


. "$PSScriptRoot\forms.ps1"


. "$PSScriptRoot\regex_operations.ps1"



#$pdfText = [PdfProcessor]::GetPdfText()





#$regex0 = [RegexOperations]::ExtractQuotation($pdfText)
#$regex1 = [RegexOperations]::ExtractItemDescription($pdfText)
#$regex2 = [RegexOperations]::RemovePricingDetails($regex1)
#$tenure = [RegexOperations]::ExtractPaymentTenure($pdfText)
#$shippingInfo = [RegexOperations]::ExtractShippingCost($pdfText)








# Line by line capture for descriptions, quantity, and prices


#$customerInfo = [regex]::Matches($regex0, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value

#$sitesStates = [regex]::Matches($regex1, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value
# Print the $sitesStates variable to the terminal
#Write-Host "Sites States:"
#$sitesStates | ForEach-Object { Write-Host $_ }

# Line by Line capture to build datatable with descriptions only
#$sitesStatesRegex2 = [regex]::Matches($regex2, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value
# Print the $sitesStatesRegex2 variable to the terminal
#Write-Host "`nSites States Regex2:"
#$sitesStatesRegex2 | ForEach-Object { Write-Host $_ }



#In PowerShell, when you declare a variable within a method, its scope is local to that method 
#unless explicitly defined otherwise. Each method in the DataTableManager class creates and 
#returns a new System.Data.DataTable object, and the $dt variable is local to each method. 
#This means that each method's $dt variable is separate and does not interfere with others, 
#even though they are named the same.

class DataTableManager {
    static [System.Data.DataTable] CreateCustomerInfoDT() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Column1", [string])
        return $dt
    }

    static [System.Data.DataTable] Create_sitesStatesDT() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Column1", [string])
        return $dt
    }

    static [System.Data.DataTable] Create_sitesStatesFinal() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Column1", [string])
        return $dt
    }

    static [System.Data.DataTable] Create_dtPrices() {
        $dtPrices = New-Object System.Data.DataTable
        $dtPrices.Columns.Add("Qty", [decimal])
        $dtPrices.Columns.Add("List Price", [decimal])
        $dtPrices.Columns.Add("Total Price", [decimal])
        $dtPrices.Columns.Add("MRC", [decimal])
        return $dtPrices
    }

    static [System.Data.DataTable] Create_dtPrices2() {
        $dtPrices2 = New-Object System.Data.DataTable
        $dtPrices2.Columns.Add("MRC Unit Price", [string])
        $dtPrices2.Columns.Add("Units", [double])
        $dtPrices2.Columns.Add("MRC Total", [string])
        return $dtPrices2
    }

    static [System.Data.DataTable] Create_dtJoined3() {
        $dtJoined3 = New-Object System.Data.DataTable
        $dtJoined3.Columns.Add("Description", [string])
        $dtJoined3.Columns.Add("Site Number", [string])
        $dtJoined3.Columns.Add("Monthly recurring charges (MRC) Per Unit", [string])
        $dtJoined3.Columns.Add("Units", [string])
        $dtJoined3.Columns.Add("Extended MRC", [string])
        return $dtJoined3
    }



    # Add more methods for other DataTables as needed
}







class DataTableOperations1 {

    #$pdfText = [PdfProcessor]::GetPdfText()

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

    static [System.Data.DataTable] Build_customerInfoDT_DT ([string] $regex0) { #before, this had input arguments of $data/$customerInfo and the $dt
        # Building the DataTable from the input data array

      

        $dt = [DataTableManager]::CreateCustomerInfoDT()

        $data = [regex]::Matches($regex0, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value



        $rowCount = $data.Count #this will be $customerInfo
        $counter = 0

        while ($counter -lt $rowCount) {
            $line = $data[$counter]
            if (-not [string]::IsNullOrWhiteSpace($line) -and $line -notmatch "Quotation" -and $line -notmatch "Billing") {
                $dt.Rows.Add($line)
            }
            $counter++
        }

        # Cleaning the rows in the DataTable
        $pattern = "(Quote\s+No.*|Quote\s+Date.*|Valid\s+Until.*|Payment\s+Term.*)|Quoted|SpecificName|SpecificStreet|SpecificCity|SpecificStateZip|SpecificPhone|SpecificEmail|SpecificName"
        foreach ($row in $dt.Rows) {
            $currentText = $row["Column1"]
            $updatedText = $currentText -replace $pattern, ""
            $row["Column1"] = $updatedText.Trim()
        }

        return $dt
    }

    static [System.Data.Datatable] Build_SitesStatesFiltered_DT ([string] $regex2) {

      
        $sitesStatesRegex2 = [regex]::Matches($regex2, "^.*", [System.Text.RegularExpressions.RegexOptions]::IgnoreCase -bor [System.Text.RegularExpressions.RegexOptions]::Multiline).Value


        $sitesStatesDT = New-Object System.Data.DataTable
        $sitesStatesDT.Columns.Add("Column1", [string])


        # Populate the DataTable with initial data
        $rowCount = $sitesStatesRegex2.Count
        $counter = 0
        while ($counter -le $rowCount - 1) {
            $sitesStatesDT.Rows.Add($sitesStatesRegex2[$counter])
            $counter++
        }

        # Create a new DataTable for filtered results
        $sitesStatesFiltered = New-Object System.Data.DataTable
        $sitesStatesFiltered.Columns.Add("Column1", [string])

        # Filter out unwanted entries from the initial data table
        foreach ($row in $sitesStatesDT.Rows) {
            $columnValue = $row["Column1"]
            if (-not ([string]::IsNullOrWhiteSpace($columnValue)) -and
                -not ($columnValue.StartsWith("Total") -or
                      $columnValue.Contains("Sub Total") -or
                      $columnValue.Contains("Shipping") -or
                      $columnValue.Contains("Item Description") -or
                      $columnValue.Contains("Final") -or
                      $columnValue.Contains("Price"))) {
                $filteredRow = $sitesStatesFiltered.NewRow()
                $filteredRow["Column1"] = $columnValue
                $sitesStatesFiltered.Rows.Add($filteredRow)
            }
        }

        # Merge cells as needed and clean up
        for ($i = 0; $i -lt $sitesStatesFiltered.Rows.Count; $i++) {
            $currentCell = $sitesStatesFiltered.Rows[$i][0]
            if ($currentCell -like "*SIM + Wifi*") {
                if ($i -lt $sitesStatesFiltered.Rows.Count - 1) {
                    $nextCell = $sitesStatesFiltered.Rows[$i + 1][0]
                    $sitesStatesFiltered.Rows[$i][0] = $currentCell.TrimEnd("`n") + $nextCell
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

            return $sitesStatesFiltered
        }

      # Static method to perform final filtering and create SitesStatesFinal DataTable
    static [System.Data.DataTable] Build_sitesStatesFinal_DT ([System.Data.DataTable] $sitesStatesFiltered) {
        $sitesStatesFinal = New-Object System.Data.DataTable
        $sitesStatesFiltered.Columns | ForEach-Object {
            $sitesStatesFinal.Columns.Add($_.ColumnName, $_.DataType)
        }
        foreach ($row in $sitesStatesFiltered.Rows) {
            $newRow = $sitesStatesFinal.NewRow()
            $newRow["Column1"] = [System.Text.RegularExpressions.Regex]::Replace($row["Column1"].ToString(), "Bundle\s+SubTotal|\$|\d{1,2},?\d{3}\.\d{2}", "").Trim()
            $sitesStatesFinal.Rows.Add($newRow)
        }
        return $sitesStatesFinal
    }
    
    static [int[]] FindIndexesOfTrigger([System.Data.DataTable] $dt, [string] $triggerString) {


        $indexArray = @()
        for ($i = 0; $i -lt $dt.Rows.Count; $i++) {
            if ($dt.Rows[$i][0].ToString().Contains($triggerString)) {
                $indexArray += $i
            }
        }
        return $indexArray #you should pass $sitesStatesFiltered as the datatable, and Bundle as the $triggerstring
    }


    static [System.Data.DataTable] Build_dtPrices1_DT([string] $regex1) {
        $regexPricesPattern = "\d+\s*\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}"
        $regexPrices = [System.Text.RegularExpressions.Regex]::new($regexPricesPattern)
        $matchesPrice = $regexPrices.Matches($regex1) | ForEach-Object { $_ }

        $dtPrices = [DataTableManager]::Create_dtPrices()

        foreach ($currentMatch in $matchesPrice) {
            $currentMatchResults = $currentMatch.Value -replace "\s+", " "
            $dtPrices.Rows.Add($currentMatchResults.Split(' '))
        }

        for ($i = $dtPrices.Rows.Count - 1; $i -ge 0; $i--) {
            $row = $dtPrices.Rows[$i]
            if ([string]::IsNullOrEmpty($row[0].ToString())) {
                $dtPrices.Rows.RemoveAt($i)
            }
        }

        $dtPrices1 = $dtPrices.Copy()

        return $dtPrices1

    }

    

        

        


    static [System.Data.DataTable] Build_dtPrices2_DT([string] $regex1, [System.Data.DataTable] $sitesStatesFinal, [System.Data.DataTable] $sitesStatesFiltered, [double] $marginSelection, [double] $price) {
        $regexPricesPattern = "\d+\s*\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}"
        $regexPrices = [System.Text.RegularExpressions.Regex]::new($regexPricesPattern)
        $matchesPrice = $regexPrices.Matches($regex1) | ForEach-Object { $_ }

        $dtPrices = [DataTableManager]::Create_dtPrices()

        foreach ($currentMatch in $matchesPrice) {
            $currentMatchResults = $currentMatch.Value -replace "\s+", " "
            $dtPrices.Rows.Add($currentMatchResults.Split(' '))
        }

        for ($i = $dtPrices.Rows.Count - 1; $i -ge 0; $i--) {
            $row = $dtPrices.Rows[$i]
            if ([string]::IsNullOrEmpty($row[0].ToString())) {
                $dtPrices.Rows.RemoveAt($i)
            }
        }

        $dtPrices1 = $dtPrices.Copy()

        $dtPrices2 = [DataTableManager]::Create_dtPrices2()


        foreach ($currentRow in $dtPrices1.Rows) {
            $value = [double]$currentRow[2] / $marginSelection 
            $quantity = [int]$currentRow[0]
            $result = [Math]::Round($value / $quantity, 2)
            $newRowData = @(
                $result.ToString(),
                $currentRow[0].ToString(),
                [Math]::Round($value, 2).ToString()
            )
            $newRow = $dtPrices2.NewRow()
            $newRow.ItemArray = $newRowData
            $dtPrices2.Rows.Add($newRow)
        }

        $siteNumberColumn = New-Object System.Data.DataColumn "Site Number", ([string])
        $dtPrices2.Columns.Add($siteNumberColumn)
        $siteNumberColumn.SetOrdinal(0)
        foreach ($row in $dtPrices2.Rows) {
            $row["Site Number"] = "1"
        }

        $initialRow = $dtPrices2.NewRow()
        $dtPrices2.Rows.InsertAt($initialRow, 0)

        $indexArray1 = [DataTableOperations1]::FindIndexesOfTrigger($sitesStatesFiltered, "Bundle Subtotal $")
        foreach ($currentItem in $indexArray1) {
            $newRow = $dtPrices2.NewRow()
            $dtPrices2.Rows.InsertAt($newRow, $currentItem)
            $dtPrices2.Rows.InsertAt($dtPrices2.NewRow(), $currentItem + 1)
        }

        return $dtPrices2

    }


    static [System.Data.DataTable] Build_dtJoined3_DT([string] $regex1, [System.Data.DataTable] $sitesStatesFinal, [double] $marginSelection, [double] $price) {
        $regexPricesPattern = "\d+\s*\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}"
        $regexPrices = [System.Text.RegularExpressions.Regex]::new($regexPricesPattern)
        $matchesPrice = $regexPrices.Matches($regex1) | ForEach-Object { $_ }

        $dtPrices = [DataTableManager]::Create_dtPrices()

        foreach ($currentMatch in $matchesPrice) {
            $currentMatchResults = $currentMatch.Value -replace "\s+", " "
            $dtPrices.Rows.Add($currentMatchResults.Split(' '))
        }

        for ($i = $dtPrices.Rows.Count - 1; $i -ge 0; $i--) {
            $row = $dtPrices.Rows[$i]
            if ([string]::IsNullOrEmpty($row[0].ToString())) {
                $dtPrices.Rows.RemoveAt($i)
            }
        }

        $dtPrices1 = $dtPrices.Copy()

        $dtPrices2 = [DataTableManager]::Create_dtPrices2()


        foreach ($currentRow in $dtPrices1.Rows) {
            $value = [double]$currentRow[2] / $marginSelection 
            $quantity = [int]$currentRow[0]
            $result = [Math]::Round($value / $quantity, 2)
            $newRowData = @(
                $result.ToString(),
                $currentRow[0].ToString(),
                [Math]::Round($value, 2).ToString()
            )
            $newRow = $dtPrices2.NewRow()
            $newRow.ItemArray = $newRowData
            $dtPrices2.Rows.Add($newRow)
        }

        $siteNumberColumn = New-Object System.Data.DataColumn "Site Number", ([string])
        $dtPrices2.Columns.Add($siteNumberColumn)
        $siteNumberColumn.SetOrdinal(0)
        foreach ($row in $dtPrices2.Rows) {
            $row["Site Number"] = "1"
        }

        $initialRow = $dtPrices2.NewRow()
        $dtPrices2.Rows.InsertAt($initialRow, 0)

        $indexArray1 = [DataTableOperations1]::FindIndexesOfTrigger($dtPrices2, "Bundle Subtotal $")
        foreach ($currentItem in $indexArray1) {
            $newRow = $dtPrices2.NewRow()
            $dtPrices2.Rows.InsertAt($newRow, $currentItem)
            $dtPrices2.Rows.InsertAt($dtPrices2.NewRow(), $currentItem + 1)
        }

        $dtJoined3 = [DataTableManager]::Create_dtJoined3()

        foreach ($currentRow1 in $sitesStatesFinal.Rows) {
            foreach ($currentRow2 in $dtPrices2.Rows) {
                if ($sitesStatesFinal.Rows.IndexOf($currentRow1) -eq $dtPrices2.Rows.IndexOf($currentRow2)) {
                    $joinedRow = $dtJoined3.NewRow()
                    $joinedRow.ItemArray = $currentRow1.ItemArray + $currentRow2.ItemArray
                    $dtJoined3.Rows.Add($joinedRow)
                }
            }
        }

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

        $newRowForTotal = $dtJoined3.NewRow()
        $newRowForTotal[3] = "Total MRC:"
        $newRowForTotal[4] = '$' + $mrcSUM
        $dtJoined3.Rows.Add($newRowForTotal)

        $newRow = $dtJoined3.NewRow()
        $newRow[3] = "Total MRC for MPP:"
        $newRow[4] = "N/A"
        $dtJoined3.Rows.Add($newRow)

        for ($i = 0; $i -lt 2; $i++) {
            $blankRow = $dtJoined3.NewRow()
            $dtJoined3.Rows.Add($blankRow)
        }

        $shippingRow = $dtJoined3.NewRow()
        $shippingRow[3] = "Shipping Costs of AT&T Equipment, One Time Charge - (OTC)"
        $shippingRow[4] = '$' + $price #variable for getting shipping cost
        $dtJoined3.Rows.Add($shippingRow)

        $initialRowData_final = @(" ", " ", " ", " ")
        $initialRow_final = $dtJoined3.NewRow()
        $initialRow_final.ItemArray = $initialRowData_final
        $dtJoined3.Rows.InsertAt($initialRow_final, 0)

        foreach ($row in $dtJoined3.Rows) {
            if ($row[0] -like "*Total*") {
                $dtJoined3.Rows.Remove($row)
                break
            }
        }

        $dtJoined3.AcceptChanges()

        return $dtJoined3

    }

}





































