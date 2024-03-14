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
    $price = $matches[2]/.85
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
    $script:marginSelection_Show = .26
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
    $script:marginSelection_Show = .35
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
