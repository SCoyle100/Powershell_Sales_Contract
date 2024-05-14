# Sales Contract Automation Script

This PowerShell script is designed to be an attended automation that aids in generating contractual sales agreements.
It is a PowerShell adaptation of a C# UiPath project I created [UiPath_Sales_Contract](https://github.com/SCoyle100/UiPath_Sales_Contract). 
This script leverages Windows Forms to provide user-friendly pop-up dialog boxes, guiding users through the process of selecting necessary documents such as a PDF quote and a vendor's statement of work.

## **Features**

- **PDF Quote Selection:** Utilizes a Windows Form dialog box to allow users to easily select the PDF file containing the quote for the sales contract.
- **Data Parsing and Rearrangement:** Parses the datatable contained within the PDF quote, rearranges the data to fit the required format for the sales contract.
- **Margin Calculations:** Performs margin calculations to ensure profitability and accuracy of the sales contract.
- **Excel Creation** Information extracted from the pdf quote is also written to 2 different Excel files
  

## **Latest Updates - 5/13/2024**

I am in the process of refactoring this into an object oriented, facade pattern.  


