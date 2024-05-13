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
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Qty", [decimal])
        $dt.Columns.Add("List Price", [decimal])
        $dt.Columns.Add("Total Price", [decimal])
        $dt.Columns.Add("MRC", [decimal])
        return $dt
    }

    static [System.Data.DataTable] Create_dtPrices2() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("MRC Unit Price", [string])
        $dt.Columns.Add("Units", [double])
        $dt.Columns.Add("MRC Total", [string])
        return $dt
    }

    static [System.Data.DataTable] Create_dtJoined3() {
        $dt = New-Object System.Data.DataTable
        $dt.Columns.Add("Description", [string])
        $dt.Columns.Add("Site Number", [string])
        $dt.Columns.Add("Monthly recurring charges (MRC) Per Unit", [string])
        $dt.Columns.Add("Units", [string])
        $dt.Columns.Add("Extended MRC", [string])
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

        #$pdfProcessor = [PdfProcessor]::new()
        #$pdfProcessor.ConvertToText($pdfFilePath)
        #$pdfText = [PdfProcessor]::GetPdfText()

        $dt = [DataTableManager]::CreateCustomerInfoDT()

        #$regex0 = [RegexOperations]::ExtractQuotation($pdfText)

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

        #$sitesStatesFiltered = New-Object System.Data.DataTable
        #$sitesStatesFinal = New-Object System.Data.Datatable
    
    
        #$pdfText = [PdfProcessor]::GetPdfText()

        #$regex1 = [RegexOperations]::ExtractItemDescription($pdfText)
        #$regex2 = [RegexOperations]::RemovePricingDetails($regex1)
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
        return $indexArray
    }

}





































