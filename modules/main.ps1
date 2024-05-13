# Load necessary assemblies
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook


# This constructs the path based on the directory of the currently running script
. "$PSScriptRoot\forms.ps1"
. "$PSScriptRoot\regex_operations.ps1"
. "$PSScriptRoot\datatables_variables.ps1"



function Test-PdfProcessor {
    $pdfForm = [PdfFileSelectionForm]::new()
    $selectedFile = $pdfForm.SelectFile()
    Write-Host "Selected PDF File: $selectedFile"  # Debug output

    if (-not [string]::IsNullOrWhiteSpace($selectedFile)) {
        $pdfProcessor = [PdfProcessor]::new()
        if ($null -eq $pdfProcessor ){
            Write-Host "Failed to create PdfProcessor"
        } else {

        
        $pdfText = $pdfProcessor.ConvertToText($selectedFile)

        }

        return $pdfText # This needs to be returned in order to be accessed from other functions in this main script
        #Write-Host "Extracted Text: $pdfText"

        #$quotation = [RegexOperations]::ExtractQuotation($pdfText)

        $quotation2 = [RegexOperations1]::ExtractItemDescription($pdfText)

        $quotation3 = [RegexOperations1]::RemovePricingDetails($pdfText)


        #Write-Host "REGEX CLASS WORKS!: $quotation"

        Write-Host "REGEX CLASS WORKS!: $quotation2"

        Write-Host "REGEX CLASS WORKS!: $quotation3"


    } else {
        Write-Host "No PDF file selected or file path is empty."
    }


}

# Function to test RegexOperations
function Test-RegexOperations {

    param([string] $pdfText) #this passes the variable pdfText into this function, otherwise, it would be bound to the local scope in the function where it was declared
    
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



# Function to test RegexOperations
function Test-RegexOperations_2 {
    param([string] $pdfText) #this passes the variable pdfText into this function, otherwise, it would be bound to the local scope in the function where it was declared
    
    Write-Host "Debug: Testing RegexOperations with text length: $($pdfText.Length)"  # Debug output to confirm function call and text length

    $quotation = [RegexOperations]::ExtractQuotation($pdfText)
    $itemDescription = [RegexOperations]::ExtractItemDescription($pdfText)
    

    Write-Host $quotation
    Write-Host $itemDescription
}


function Test-DataTableOperations{

    $regex1 = [RegexOperations]::ExtractItemDescription($pdfText)
    $regex2 = [RegexOperations]::RemovePricingDetails($regex1)

    $sitesStatesFiltered = [DataTableOperations1]::Build_SitesStatesFiltered_DT($regex2)
    $sitesStatesFinal = [DataTableOperations1]::Build_SitesStatesFinal_DT($sitesStatesFiltered)

    # Example operation: printing row counts
    Write-Host $sitesStatesFiltered.Rows.Count
    Write-Host $sitesStatesFinal.Rows.Count

}






$pdfText = Test-PdfProcessor






Test-RegexOperations $pdfText

Test-DataTableOperations 



#Test-RegexOperations_2 $pdfText





#Test-RegexOperations $pdfText
#Test-MarginSelectionForm
#Test-InputDialogWithSkip
#Test-OutlookGALUserDetails





