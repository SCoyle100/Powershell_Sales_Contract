
. "$PSScriptRoot\forms.ps1"



class RegexOperations1 {

    static [string] ExtractQuotation([string] $pdfText) {

        #I believe this is to build the $customerInfoDT table

        Write-Host "Debug: Extracting Quotation from text"  # Debug output

        if ($pdfText -match "Quotation[\s\S]+?Quoted") {
            return $matches[0]
        } else {
            Write-Host "Debug: Pattern not found in text"  # Debug output
            return "Pattern not found"
        }
    }

    static [string] ExtractItemDescription([string] $pdfText) {
        if ($pdfText -match "Item Description[\s\S]+?Final Quote") {
            return $matches[0]
        } else {
            return "Pattern not found"
        }
    }

    static [string] RemovePricingDetails([string] $pdfText) {
        return $pdfText -replace "\d+\s*\d{1,3},\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\d+\s*\d{3}\.\d{2}\s*\s*\d{1,3},\d{3}\.\d{2}\s*\d{1,3}\.\d{2}|\[[^\]]*\]|\$\s*\d+\s*\d*\.\d{2}|\d+\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}\s+\d{1,3}(,\d{3})*\.\d{2}", ""
    }

    static [string] ExtractPaymentTenure([string] $pdfText) {
        $term = "Payment Tenure\s*:\s*(\d+)\s*Months"
        if ($pdfText -match $term) {
            return "Tenure: $($matches[0]) months"
        } else {
            return "Pattern not found."
        }
    }

    static [string] ExtractShippingCost([string] $pdfText) {
        $shipping = "Shipping\s*Cost\s*for\s*(\d{1,3}) Qty\s*\$\s*([\d\.]+)"
        if ($pdfText -match $shipping) {
            $quantity = $matches[1] #this was 1 before? 
            $price = [double]$matches[2] / 0.85 #try using 1 or 2 here? this was 2 before
            return "Quantity: $quantity, Price: $price"
        } else {
            return "Pattern not found."
        }
    }
}
