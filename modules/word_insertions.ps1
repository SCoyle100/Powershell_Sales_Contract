# Load Word COM object
Add-Type -AssemblyName Microsoft.Office.Interop.Word




class WordAutomation {




    [object]$Word
    [object]$TemplateDoc

    WordAutomation([string]$contractTemplate) {
        $this.Word = New-Object -ComObject Word.Application
        $this.TemplateDoc = $this.Word.Documents.Open($contractTemplate)
        $this.Word.Visible = $true
    }

    [void] FindPlaceholderText([string]$FindText, [string]$ReplaceText) {
        $find = $this.TemplateDoc.Content.Find
        $find.ClearFormatting()
        while ($find.Execute($FindText)) {
            $textRange = $find.Parent
            $textRange.Text = $ReplaceText
            $find.Wrap = 1 # wdFindContinue
        }
    }

    [void] AddTable([System.Data.DataTable]$dtJoined3, [array]$currencyColumnIndices) {
        $findText = "<<pricingTable>>"
        $find = $this.TemplateDoc.Content.Find
        $find.ClearFormatting()

        if ($find.Execute($findText)) {
            $dataTableRange = $find.Parent
            $dataTableRange.Select()

            $rowCount = $dtJoined3.Rows.Count
            $columnCount = $dtJoined3.Columns.Count
            $wordTable1 = $this.TemplateDoc.Tables.Add($dataTableRange, $rowCount + 1, $columnCount)

            $wordTable1.Rows.Alignment = [Microsoft.Office.Interop.Word.WdRowAlignment]::wdAlignRowCenter

            foreach ($cell in $wordTable1.Range.Cells) {
                $cell.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
                $cell.VerticalAlignment = [Microsoft.Office.Interop.Word.WdCellVerticalAlignment]::wdCellAlignVerticalCenter
                $cell.Range.Font.Name = "Arial"
                $cell.Range.Font.Size = 8
                $cell.Range.ParagraphFormat.SpaceBefore = 0
                $cell.Range.ParagraphFormat.SpaceAfter = 0
                $cell.Range.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.WdLineSpacing]::wdLineSpaceSingle
            }

            for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
                $headerText = [System.Convert]::ToString($dtJoined3.Columns[$columnIndex].ColumnName)
                $wordTable1.Cell(1, $columnIndex + 1).Range.Text = $headerText
            }

            for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
                for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
                    if ($wordTable1.Cell($rowIndex + 2, $columnIndex + 1)) {
                        $cellData = $dtJoined3.Rows[$rowIndex][$columnIndex] -as [String]
                        if ($currencyColumnIndices -contains $columnIndex -and $cellData -match '^\d+(\.\d+)?$') {
                            $cellData = ('${0:N2}' -f [double]$cellData)
                        }
                        $wordTable1.Cell($rowIndex + 2, $columnIndex + 1).Range.Text = $cellData
                    }
                }
            }

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

            $this.RemoveBordersForSpecialRows($wordTable1)
            $this.MakeSpecificCellsBold($wordTable1, $indexArray1)
        }
    }

    [void] AddSKUTable([System.Data.DataTable] $dtSKU2) {
        $findTextSKU = "<<skuTable>>"
        $findSKU = $this.TemplateDoc.Content.Find
        $findSKU.ClearFormatting()

        if ($findSKU.Execute($findTextSKU)) {
            $dataTableRangeSKU = $findSKU.Parent
            $dataTableRangeSKU.Select()

            $rowCountSKU = $dtSKU2.Rows.Count
            $columnCountSKU = $dtSKU2.Columns.Count
            $wordTableSKU = $this.TemplateDoc.Tables.Add($dataTableRangeSKU, $rowCountSKU + 1, $columnCountSKU)

            $wordTableSKU.Rows.Alignment = [Microsoft.Office.Interop.Word.WdRowAlignment]::wdAlignRowCenter

            foreach ($cellSKU in $wordTableSKU.Range.Cells) {
                $cellSKU.Range.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.WdParagraphAlignment]::wdAlignParagraphCenter
                $cellSKU.VerticalAlignment = [Microsoft.Office.Interop.Word.WdCellVerticalAlignment]::wdCellAlignVerticalCenter
                $cellSKU.Range.Font.Name = "Arial"
                $cellSKU.Range.Font.Size = 8
            }

            for ($columnIndexSKU = 0; $columnIndexSKU -lt $columnCountSKU; $columnIndexSKU++) {
                $headerTextSKU = [System.Convert]::ToString($dtSKU2.Columns[$columnIndexSKU].ColumnName)
                $wordTableSKU.Cell(1, $columnIndexSKU + 1).Range.Text = $headerTextSKU
            }

            for ($rowIndexSKU = 0; $rowIndexSKU -lt $rowCountSKU; $rowIndexSKU++) {
                for ($columnIndexSKU = 0; $columnIndexSKU -lt $columnCountSKU; $columnIndexSKU++) {
                    if ($wordTableSKU.Cell($rowIndexSKU + 2, $columnIndexSKU + 1)) {
                        $cellDataSKU = $dtSKU2.Rows[$rowIndexSKU][$columnIndexSKU] -as [String]
                        $wordTableSKU.Cell($rowIndexSKU + 2, $columnIndexSKU + 1).Range.Text = $cellDataSKU
                    }
                }
            }

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
    }

    [void] RemoveBordersForSpecialRows([object]$wordTable1, [array] $indexArray1) {
        $startRow = $wordTable1.Rows.Count - 5
        for ($rowIndex = $startRow; $rowIndex -le $wordTable1.Rows.Count; $rowIndex++) {
            for ($colIndex = 1; $colIndex -le 3; $colIndex++) {
                $cell = $wordTable1.Cell($rowIndex, $colIndex)
                $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
                $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
                $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
            }
        }

        foreach ($index in $indexArray1) {
            $rowIndex = $index + 3
            if ($rowIndex -le $wordTable1.Rows.Count) {
                for ($i = 1; $i -lt $wordTable1.Columns.Count; $i++) {
                    $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
                }
            }
        }

        for ($j = 0; $j -lt $indexArray1.Count - 1; $j++) {
            $index = $indexArray1[$j]
            $rowIndex = $index + 3
            if ($rowIndex -le $wordTable1.Rows.Count) {
                for ($i = 2; $i -le $wordTable1.Columns.Count; $i++) {
                    $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
                }
            }
        }

        $secondToLastRow = $wordTable1.Rows.Count - 1
        $thirdToLastRow = $wordTable1.Rows.Count - 2
        $columnsToModify = 4, 5
        $rowsToModify = @($thirdToLastRow, $secondToLastRow)

        foreach ($rowIndex in $rowsToModify) {
            foreach ($colIndex in $columnsToModify) {
                $cell = $wordTable1.Cell($rowIndex, $colIndex)
                $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
                $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
            }
        }

        foreach ($colIndex in $columnsToRemoveBottomBorder) {
            $cell = $wordTable1.Cell($thirdToLastRow, $colIndex)
            $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        }
    }

    [void] MakeSpecificCellsBold([object]$wordTable1, [array]$indexArray1) {
        $cell = $wordTable1.Cell(3, 1)
        $cell.Range.Font.Bold = $true
        $cell.Range.Font.Size = 9

        for ($j = 0; $j -lt $indexArray1.Count - 1; $j++) {
            $index = $indexArray1[$j]
            $rowIndex = $index + 4
            if ($rowIndex -le $wordTable1.Rows.Count) {
                $cell = $wordTable1.Cell($rowIndex, 1)
                $cell.Range.Font.Bold = $true
                $cell.Range.Font.Size = 9
            }
        }
    }

    [void] FormatCoverPage([array]$specificCells, [array]$stateAbbreviations) {
        $wordTable = $this.TemplateDoc.Tables[1]
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
                        } elseif ($words[$i] -match "@") {
                            $words[$i] = $words[$i].ToLower()
                        } else {
                            $words[$i] = [Globalization.CultureInfo]::CurrentCulture.TextInfo.ToTitleCase($words[$i].ToLower())
                        }
                    }

                    $newText = $words -join " "
                    foreach ($abbreviation in $stateAbbreviations) {
                        $pattern = "\b$abbreviation\b"
                        $newText = [Regex]::Replace($newText, $pattern, $abbreviation.ToUpper(), [Text.RegularExpressions.RegexOptions]::IgnoreCase)
                    }

                    $cell.Range.Text = $newText
                    $cell.Range.Font.Name = "Arial"
                    $cell.Range.Font.Size = 8
                    $cell.Range.HighlightColorIndex = [Microsoft.Office.Interop.Word.WdColorIndex]::wdNoHighlight
                }
            }
        }
    }

    [void] SaveAndClose([string]$newFilePath) {
        $this.TemplateDoc.SaveAs([ref] $newFilePath)
        $this.TemplateDoc.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.TemplateDoc) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($this.Word) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}



function Find-PlaceholderText {
    param (
        [Parameter(Mandatory=$true)]
        [object]$Document,

        [Parameter(Mandatory=$true)]
        [string]$FindText,

        [Parameter(Mandatory=$true)]
        [string]$ReplaceText
    )

    $find = $Document.Content.Find
    $find.ClearFormatting()

    while ($find.Execute($FindText)) {
        $textRange = $find.Parent
        $textRange.Text = $ReplaceText
        $find.Wrap = 1 # wdFindContinue
    }
}


Find-PlaceholderText -Document $templateDoc -FindText "<<customer name>>" -ReplaceText $customerName
Find-PlaceholderText -Document $templateDoc -FindText "<<customer contact>>" -ReplaceText $customerContact
Find-PlaceholderText -Document $templateDoc -FindText "<<customer email>>" -ReplaceText $customerEmail
Find-PlaceholderText -Document $templateDoc -FindText "<<customer phone>>" -ReplaceText $customerPhone
Find-PlaceholderText -Document $templateDoc -FindText "<<customer street>>" -ReplaceText $customerStreet
Find-PlaceholderText -Document $templateDoc -FindText "<<customer city>>" -ReplaceText $customerCity
Find-PlaceholderText -Document $templateDoc -FindText "<<customer state>>" -ReplaceText $customerState
Find-PlaceholderText -Document $templateDoc -FindText "<<customer zip>>" -ReplaceText $customerZip

Find-PlaceholderText -Document $templateDoc -FindText "<<sales name>>" -ReplaceText $salesName
Find-PlaceholderText -Document $templateDoc -FindText "<<sales street>>" -ReplaceText $salesStreetAddress
Find-PlaceholderText -Document $templateDoc -FindText "<<sales city>>" -ReplaceText $salesCity
Find-PlaceholderText -Document $templateDoc -FindText "<<sales state>>" -ReplaceText $salesState
Find-PlaceholderText -Document $templateDoc -FindText "<<sales zip>>" -ReplaceText $salesZip
Find-PlaceholderText -Document $templateDoc -FindText "<<sales manager>>" -ReplaceText $salesManagerName




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


    
###Border removal for first and 2nd rows###

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

###End of border removal for first and second rows###




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

for ($j = 0; $j -lt $indexArray1.Count - 1; $j++) {
    $index = $indexArray1[$j]
    $rowIndex = $index + 3 # Adjusting each index as specified
    
    # Ensure the row index is within the table's bounds
    if ($rowIndex -le $wordTable1.Rows.Count) {
        # Loop through all but the last cell in the specified row to remove inner vertical borders
        for ($i = 2; $i -le $wordTable1.Columns.Count; $i++) {
            $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        }
    }
}


#Removing bottom left outer borders


# Calculate the starting row for the last 5 rows
$startRow = $wordTable1.Rows.Count - 5

# Iterate over the last 5 rows
for ($rowIndex = $startRow; $rowIndex -le $wordTable1.Rows.Count; $rowIndex++) {
    # Access the first cell of the current row to remove the left border
    $cell = $wordTable1.Cell($rowIndex, 1)

    # Remove the left border
    $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
}


#$startRow = $wordTable1.Rows.Count - 4

# Iterate over the last 5 rows
#for ($rowIndex = $startRow; $rowIndex -le $wordTable1.Rows.Count; $rowIndex++) {
    # Access the first cell of the current row to remove the left border
#    $cell = $wordTable1.Cell($rowIndex, 3)

    # Remove the left border
#    $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
#}



$startRow = $wordTable1.Rows.Count - 5

# Iterate over the last 5 rows
for ($rowIndex = $startRow; $rowIndex -le $wordTable1.Rows.Count; $rowIndex++) {
    # Iterate over the first 3 columns
    for ($colIndex = 1; $colIndex -le 3; $colIndex++) {
        # Access the cell in the current row and column to remove the specified borders
        $cell = $wordTable1.Cell($rowIndex, $colIndex)

        # Remove the specified borders for this cell
        $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    }
}


$startRow = $wordTable1.Rows.Count - 5

# Iterate over the last 5 rows
for ($rowIndex = $startRow; $rowIndex -le $wordTable1.Rows.Count; $rowIndex++) {
    # Iterate over the first 3 columns
    for ($colIndex = 1; $colIndex -le 3; $colIndex++) {
        # Access the cell in the current row and column to remove the specified borders
        $cell = $wordTable1.Cell($rowIndex, $colIndex)

        # Remove the specified borders for this cell
        $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    }
}



#foreach ($index in $indexArray1) {
#    $rowIndex = $index + 4 # Adjusting each index as specified (+1)
    
    # Ensure the row index is within the table's bounds
#    if ($rowIndex -le $wordTable1.Rows.Count) {
        # Loop through all but the last cell in the specified row to remove inner vertical borders
#        for ($i = 2; $i -lt $wordTable1.Columns.Count; $i++) {
#            $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
#        }
#    }
#}

for ($j = 0; $j -lt $indexArray1.Count - 1; $j++) {
    $index = $indexArray1[$j]
    $rowIndex = $index + 4 # Adjusting each index as specified
    
    # Ensure the row index is within the table's bounds
    if ($rowIndex -le $wordTable1.Rows.Count) {
        # Loop through all but the last cell in the specified row to remove inner vertical borders
        for ($i = 2; $i -lt $wordTable1.Columns.Count; $i++) {
            $wordTable1.Cell($rowIndex, $i).Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderRight).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        }
    }
}


# Calculate the indices for the 2nd to last and 3rd to last rows
$secondToLastRow = $wordTable1.Rows.Count - 1
$thirdToLastRow = $wordTable1.Rows.Count - 2

# Define the columns to modify
$columnsToModify = 4, 5

# Create an array of the specific rows to modify
$rowsToModify = @($thirdToLastRow, $secondToLastRow)

foreach ($rowIndex in $rowsToModify) {
    foreach ($colIndex in $columnsToModify) {
        # Access the cell in the current row and column to remove the specified borders
        $cell = $wordTable1.Cell($rowIndex, $colIndex)

        # Remove the left, vertical, and bottom borders for this cell
        $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderLeft).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderVertical).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
        #$cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
    }
}


# Calculate the index for the 3rd to last row
$thirdToLastRow = $wordTable1.Rows.Count - 2

# Define the columns from which to remove the bottom border
$columnsToRemoveBottomBorder = 4, 5

# Iterate over the specified columns
foreach ($colIndex in $columnsToRemoveBottomBorder) {
    # Access the cell at the 3rd to last row and current column
    $cell = $wordTable1.Cell($thirdToLastRow, $colIndex)
    
    # Remove the bottom border for this cell
    $cell.Borders.Item([Microsoft.Office.Interop.Word.WdBorderType]::wdBorderBottom).LineStyle = [Microsoft.Office.Interop.Word.WdLineStyle]::wdLineStyleNone
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
$stateAbbreviations = @("AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA", "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ", "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT", "VA", "WA", "WV", "WI", "WY", "USA", "US", "NW", "NE", "SW", "SE", "LLC")

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
