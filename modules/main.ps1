# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook


# This constructs the path based on the directory of the currently running script
. "$PSScriptRoot\forms.ps1"




# Function to test PdfFileSelectionForm
function Test-PdfFileSelectionForm {
    $pdfForm = [PdfFileSelectionForm]::new()
    $selectedFile = $pdfForm.SelectFile()
    Write-Host "Selected PDF File: $selectedFile"
}

# Function to test PdfProcessor
function Test-PdfProcessor {
    $pdfProcessor = [PdfProcessor]::new()
    $pdfText = $pdfProcessor.ConvertToText($pdfFilePath)  # Specify a test PDF file path here
    Write-Host "Extracted Text: $pdfText"
}

# Function to test RegexOperations
function Test-RegexOperations {
    
    $quotation = [RegexOperations]::ExtractQuotation($textContent)
    $itemDescription = [RegexOperations]::ExtractItemDescription($textContent)
    $cleanedText = [RegexOperations]::RemovePricingDetails($itemDescription)
    $paymentTenure = [RegexOperations]::ExtractPaymentTenure($textContent)
    $shippingCost = [RegexOperations]::ExtractShippingCost($textContent)
    Write-Host "Quotation: $quotation"
    Write-Host "Item Description: $itemDescription"
    Write-Host "Cleaned Text: $cleanedText"
    Write-Host "Payment Tenure: $paymentTenure"
    Write-Host "Shipping Cost: $shippingCost"
}

# Function to test MarginSelectionForm
function Test-MarginSelectionForm {
    $marginForm = [MarginSelectionForm]::new()
    $marginForm.ShowDialog()
    Write-Host "Selected Margin: $($marginForm.MarginSelectionShow)"
}

# Function to test InputDialogWithSkip
function Test-InputDialogWithSkip {
    $inputDialog = [InputDialogWithSkip]::new('Enter your name:', 'Name Input')
    $inputResult = $inputDialog.ShowDialog()
    Write-Host "Input Result: $inputResult"
}

# Function to test OutlookGALUserDetails
function Test-OutlookGALUserDetails {
    $outlookDetails = [OutlookGALUserDetails]::new()
    $outlookDetails.GetUserDetailsLoop()
}

# Calling test functions
#Test-PdfFileSelectionForm
#Test-PdfProcessor
#Test-RegexOperations
Test-MarginSelectionForm
#Test-InputDialogWithSkip
#Test-OutlookGALUserDetails
