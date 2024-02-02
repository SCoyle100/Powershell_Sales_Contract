# Sales Contract Automation Script

This PowerShell script is designed to be an attended automation that aids in generating contractual sales agreements.
It is a PowerShell adaptation of a C# UiPath project I created [UiPath_Sales_Contract](https://github.com/SCoyle100/UiPath_Sales_Contract)
The script leverages Windows Forms to provide user-friendly pop-up dialog boxes, guiding users through the process of selecting necessary documents such as a PDF quote and a vendor's statement of work.

## **Features**

- **PDF Quote Selection:** Utilizes a Windows Form dialog box to allow users to easily select the PDF file containing the quote for the sales contract.
- **Data Parsing and Rearrangement:** Parses the datatable contained within the PDF quote, rearranges the data to fit the required format for the sales contract.
- **Margin Calculations:** Performs margin calculations to ensure profitability and accuracy of the sales contract.
- **Excel Creation** Information extracted from the pdf quote is also written to 2 different Excel files
  

## To Do:
- There is currently a word document selection that selects a table, creates a datatable from it for later use, as well as copying the table to be pasted, but that will be removed for this project, as the requirements have changed.
- Adding regular expressions to parse an HTML page that the user will select which contains key data (The format doesn't change, and I am aiming to reduce the usage of additional libraries such as HtmlAgilityPack)
- Reducing technical debt (There are a lot of testing portions that print out various datatables that won't be used in the final product)
- Modularizing the script into different sections for easier maintenance  


