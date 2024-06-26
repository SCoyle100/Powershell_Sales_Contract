Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing


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
    #static [string] $PdfText

    PdfProcessor() {
        $this.PdfToTextPath = "C:\Program Files\xpdf-tools-win-4.04\xpdf-tools-win-4.04\bin64\pdftotext.exe"
    }

    [string] ConvertToText([string] $pdfFilePath) {
        Write-Host "Debug: ConvertToText called"
        if ([string]::IsNullOrWhiteSpace($pdfFilePath)) {
            Write-Host "PDF file path is empty."
            return $null
        }
        $outputTxtPath = [System.IO.Path]::ChangeExtension($pdfFilePath, '.txt')
        & $this.PdfToTextPath -table $pdfFilePath $outputTxtPath
        if (Test-Path $outputTxtPath) {
            
            $pdfText = Get-Content $outputTxtPath -Raw
            Write-Host "Debug: Extracted text length is $($pdfText.Length)"
            
            return $pdfText
        } else {
            Write-Host "Failed to convert PDF to text."
            return $null
        }
    }

    static [string] ExtractQuotation([string] $pdfText) {
        Write-Host "Debug: Extracting Quotation from text"
        if ($pdfText -match "Quotation[\s\S]+?BUSINESS") {
            return $matches[0]
        } else {
            Write-Host "Debug: Pattern not found in text"
            return "Pattern not found"
        }
    }

}







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
        
        $button35 = New-Object System.Windows.Forms.Button
        $button35.Location = New-Object System.Drawing.Point(150, 50)  # Set position
        $button35.Size = New-Object System.Drawing.Size(100, 23)
        $button35.Text = "35%"

        # Use the current instance for event handlers
        $currentInstance = $this

        $button26.Add_Click([System.EventHandler]{
            $currentInstance.OnButton26Click()
            $form.Close()
        })
        
        $button35.Add_Click([System.EventHandler]{
            $currentInstance.OnButton35Click()
            $form.Close()
        })

        $form.Controls.Add($button26)
        $form.Controls.Add($button35)

        $form.Add_Shown({$form.Activate()})
        
        $form.ShowDialog() | Out-Null
    }

    [void] OnButton26Click() {
        $this.MarginSelection = 0.74
        $this.MarginSelectionShow = 0.26
    }

    [void] OnButton35Click() {
        $this.MarginSelection = 0.65
        $this.MarginSelectionShow = 0.35
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

        # OK button creation similar to MarginSelectionForm
        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Text = "OK"
        $okButton.Location = New-Object System.Drawing.Point(30, 70)
        $okButton.Size = New-Object System.Drawing.Size(100, 23)
        $okButton.Add_Click({
            $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $form.Close()
        })
    
        # Skip button creation similar to MarginSelectionForm
        $skipButton = New-Object System.Windows.Forms.Button
        $skipButton.Text = "Skip"
        $skipButton.Location = New-Object System.Drawing.Point(160, 70)
        $skipButton.Size = New-Object System.Drawing.Size(100, 23)
        $skipButton.Add_Click({
            $form.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
            $form.Close()
        })
    
        $form.Controls.Add($label)
        $form.Controls.Add($textBox)
        $form.Controls.Add($okButton)
        $form.Controls.Add($skipButton)
        
        $form.Add_Shown({$form.Activate()})
    
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

    [string] ConvertNameFormat([string] $name) {
        if ($name -like '*,*') {
            $parts = $name -split ','
            $formattedName = $($parts[1].Trim()) + " " + $($parts[0].Trim())
            return $formattedName
        } else {
            return $name
        }
    }
    



    [void] ProcessSalesDetails([PSObject] $UserDetails) {
        $salesName = $UserDetails.Name
        $salesJobTitle = $UserDetails.JobTitle
        $salesStreetAddress = $UserDetails.BusinessAddress
        $salesCity = $UserDetails.BusinessCity
        $salesState = $UserDetails.BusinessState
        $salesZip = $UserDetails.BusinessZip
        $salesPhone = $UserDetails.BusinessPhone
        $salesManagerName = $UserDetails.ManagerName

        # Convert names to preferred format
        $salesName = $this.ConvertNameFormat($salesName)
        $salesManagerName = $this.ConvertNameFormat($salesManagerName)

        # Additional logic to use these details can be added here
    }
}


    


