Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data.DataSetExtensions
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName System.Data

# Set the absolute path to your config.ps1 script
$configPath = "D:\Programming\PowerShell\Sales Contract\config.ps1"
# Dot source the config script
. $configPath


#user input pop ups
. .\forms.ps1

#logic for building datatables and capturing names, addresses, etc via regex
. .\datatables_variables.ps1


#Inserting tables and variables into the contract template
. .\word_insertions.ps1


#Inserting information into the excel template and building the FM loader as .xls
. .\excel_insertions.ps1










