Add-Type -AssemblyName System.Windows.Forms


class PdfFileSelectionForm {
    [string] $InitialDirectory
    [string] $Filter

    PdfFileSelectionForm() {
        $this.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        $this.Filter = 'PDF files (*.pdf)|*.pdf'
        $this.ShowInitialMessage()
    }

    [string] SelectFile() {
        $fileDialog = New-Object System.Windows.Forms.OpenFileDialog
        $fileDialog.InitialDirectory = $this.InitialDirectory
        $fileDialog.Filter = $this.Filter
        $fileDialog.ShowDialog() | Out-Null
        return $fileDialog.FileName
    }

    [void] ShowInitialMessage() {
        [System.Windows.Forms.MessageBox]::Show('SELECT THE PDF QUOTE', '')
    }
}




class PdfProcessor {
    [string] $PdfToTextPath

    PdfProcessor() {
        $this.PdfToTextPath = "C:\Program Files\xpdf-tools-win-4.04\xpdf-tools-win-4.04\bin64\pdftotext.exe"
    }

    [string] ConvertToText([string] $pdfFilePath) {
        $outputTxtPath = [System.IO.Path]::ChangeExtension($pdfFilePath, '.txt')
        & $this.PdfToTextPath -table $pdfFilePath $outputTxtPath
        if (Test-Path $outputTxtPath) {
            return Get-Content $outputTxtPath -Raw
        } else {
            Write-Host "Failed to convert PDF to text."
            return $null
        }
    }
}



class RegexOperations {
    static [string] ExtractQuotation([string] $textContent) {
        if ($textContent -match "Quotation[\s\S]+?Quoted") {
            return $matches[0]
        } else {
            return ""
        }
    }

    static [string] ExtractItemDescription([string] $textContent) {
        if ($textContent -match "Item Description[\s\S]+?Final Quote") {
            return $matches[0]
        } else {
            return ""
        }
    }

    static [string] RemovePricingDetails([string] $textContent) {
        return $textContent -replace "\d+\s*\d{1,3},\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\[[^\]]*\]|\$\s*\d+\s*\d*\.\d{2}|\d+\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}", ""
    }

    static [string] ExtractPaymentTenure([string] $textContent) {
        $term = "Payment Tenure\s*:\s*(\d+)\s*Months"
        if ($textContent -match $term) {
            return "Tenure: $($matches[1]) months"
        } else {
            return "Pattern not found."
        }
    }

    static [string] ExtractShippingCost([string] $textContent) {
        $shipping = "Shipping\s*Cost\s*for\s*(\d{1,3}) Qty\s*\$\s*([\d\.]+)"
        if ($textContent -match $shipping) {
            $quantity = $matches[1]
            $price = [double]$matches[2] / 0.85
            return "Quantity: $quantity, Price: $price"
        } else {
            return "Pattern not found."
        }
    }
}



#Sample usage for regex variables
# Assuming $textContent is defined and contains the text extracted from a PDF
#$regex0 = [RegexOperations]::ExtractQuotation($textContent)
#$regex1 = [RegexOperations]::ExtractItemDescription($textContent)
#$regex2 = [RegexOperations]::RemovePricingDetails($regex1)
#$tenure = [RegexOperations]::ExtractPaymentTenure($textContent)
#$shippingInfo = [RegexOperations]::ExtractShippingCost($textContent)



class MarginSelectionForm {
    [double] $MarginSelection
    [double] $MarginSelectionShow

    MarginSelectionForm() {
        $this.MarginSelection = 0
        $this.MarginSelectionShow = 0
    }

    [void] ShowDialog() {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = "Please Select Margin"
        $form.Size = New-Object System.Drawing.Size(400, 300)
        $form.StartPosition = "CenterScreen"

        $button26 = New-Object System.Windows.Forms.Button
        $button26.Text = "26%"
        $button26.Location = New-Object System.Drawing.Point(30, 50)  # Set position
        $button26.Size = New-Object System.Drawing.Size(100, 23)
        $button26.Add_Click({
            $this.MarginSelection = 0.74
            $this.MarginSelectionShow = 0.26
            $form.Close()
        })

        $form.Controls.Add($button26)
        

        $button35 = New-Object System.Windows.Forms.Button
        $button35.Location = New-Object System.Drawing.Point(150, 50)  # Set position
        $button35.Size = New-Object System.Drawing.Size(100, 23)
        $button35.Text = "35%"
        $button35.Add_Click({
            $this.MarginSelection = 0.65
            $this.MarginSelectionShow = 0.35
            $form.Close()
        })

        
        $form.Controls.Add($button35)

        $form.Add_Shown{{$form.Activate()}}
        
        $form.ShowDialog() | Out-Null
    }
}


class InputDialogWithSkip {
    [string] $Message
    [string] $WindowTitle

    InputDialogWithSkip([string] $Message, [string] $WindowTitle = "Input") {
        $this.Message = $Message
        $this.WindowTitle = $WindowTitle
    }

    [string] ShowDialog() {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $this.WindowTitle
        $form.Size = New-Object System.Drawing.Size(300,200)
        $form.StartPosition = "CenterScreen"
    
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $this.Message
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
    
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10,40)
        $textBox.Size = New-Object System.Drawing.Size(260,20)
    
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
        $skipButton = New-Object System.Windows.Forms.Button
        $skipButton.Text = "Skip"
        $skipButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    
        $form.Controls.Add($label)
        $form.Controls.Add($textBox)
        $form.Controls.Add($okButton)
        $form.Controls.Add($skipButton)
    
        $form.AcceptButton = $okButton
        $form.CancelButton = $skipButton
    
        $result = $form.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $textBox.Text
        } else {
            return $null
        }
    }
}





# Add necessary assembly for Outlook and Windows Forms
Add-Type -AssemblyName "Microsoft.Office.Interop.Outlook"
Add-Type -AssemblyName System.Windows.Forms

class OutlookGALUserDetails {
    [Microsoft.Office.Interop.Outlook.Application] $outlook
    [Microsoft.Office.Interop.Outlook.NameSpace] $namespace

    OutlookGALUserDetails() {
        $this.outlook = New-Object -ComObject Outlook.Application
        $this.namespace = $this.outlook.GetNamespace("MAPI")
    }

    [PSObject] GetGALUserDetails([string] $emailAddress) {
        $Recipient = $this.namespace.CreateRecipient($emailAddress)
        $Recipient.Resolve()

        if ($Recipient.Resolved -and $Recipient.AddressEntry.GetExchangeUser()) {
            $ExchangeUser = $Recipient.AddressEntry.GetExchangeUser()
            $Manager = $ExchangeUser.GetExchangeUserManager()
            $ManagerName = $null

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

    [string] ShowInputDialog([string] $message, [string] $title) {
        $form = New-Object System.Windows.Forms.Form
        $form.Text = $title
        $form.Size = New-Object System.Drawing.Size(300,200)
        $form.StartPosition = "CenterScreen"
    
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $message
        $label.Location = New-Object System.Drawing.Point(10,20)
        $label.Size = New-Object System.Drawing.Size(280,20)
    
        $textBox = New-Object System.Windows.Forms.TextBox
        $textBox.Location = New-Object System.Drawing.Point(10,40)
        $textBox.Size = New-Object System.Drawing.Size(260,20)
    
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(10,70)
        $okButton.Size = New-Object System.Drawing.Size(75,23)
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    
        $cancelButton = New-Object System.Windows.Forms.Button
        $cancelButton.Text = "Cancel"
        $cancelButton.Location = New-Object System.Drawing.Point(195,70)
        $cancelButton.Size = New-Object System.Drawing.Size(75,23)
        $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    
        $form.Controls.Add($label)
        $form.Controls.Add($textBox)
        $form.Controls.Add($okButton)
        $form.Controls.Add($cancelButton)
    
        $form.AcceptButton = $okButton
        $form.CancelButton = $cancelButton
    
        $result = $form.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            return $textBox.Text
        } else {
            return $null
        }
    }

}
