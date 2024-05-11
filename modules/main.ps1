# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook


# This constructs the path based on the directory of the currently running script
. "$PSScriptRoot\forms.ps1"



function Test-PdfProcessor {
    $pdfForm = [PdfFileSelectionForm]::new()
    $selectedFile = $pdfForm.SelectFile()
    Write-Host "Selected PDF File: $selectedFile"  # Debug output

    if (-not [string]::IsNullOrWhiteSpace($selectedFile)) {
        $pdfProcessor = [PdfProcessor]::new()
        $pdfText = $pdfProcessor.ConvertToText($selectedFile)
        Write-Host "Extracted Text: $pdfText"
    } else {
        Write-Host "No PDF file selected or file path is empty."
    }
}

# Function to test RegexOperations
function Test-RegexOperations {

    #param([string] $pdfText)
    
    $quotation = [RegexOperations]::ExtractQuotation($pdfText)
    $itemDescription = [RegexOperations]::ExtractItemDescription($pdfText)
    $cleanedText = [RegexOperations]::RemovePricingDetails($itemDescription)
    $paymentTenure = [RegexOperations]::ExtractPaymentTenure($pdfText)
    $shippingCost = [RegexOperations]::ExtractShippingCost($pdfText)
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
    $inputDialog = [InputDialogWithSkip]::new('Enter email address or UID:', 'Name Input')
    $inputResult = $inputDialog.ShowDialog()
    Write-Host "Input Result: $inputResult"
}

# Function to test OutlookGALUserDetails
function Test-OutlookGALUserDetails {
    $outlookDetails = [OutlookGALUserDetails]::new()
    $outlookDetails.GetUserDetailsLoop()
}

# Calling test functions
Test-PdfProcessor
Test-RegexOperations $pdfText
#Test-MarginSelectionForm
#Test-InputDialogWithSkip
#Test-OutlookGALUserDetails


$regex0 = [RegexOperations]::ExtractQuotation($pdfText)
$regex1 = [RegexOperations]::ExtractItemDescription($pdfText)
$regex2 = [RegexOperations]::RemovePricingDetails($regex1)
$tenure = [RegexOperations]::ExtractPaymentTenure($pdfText)
$shippingInfo = [RegexOperations]::ExtractShippingCost($pdfText)



Write-Host $regex0
Write-Host $regex1
Write-Host $regex2
Write-Host $tenure
Write-Host $shippingInfo


