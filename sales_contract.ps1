Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Data.DataSetExtensions
Add-Type -AssemblyName Microsoft.Office.Interop.Word
Add-Type -AssemblyName System.Data


# Get the path of the current script (main.ps1)
$scriptPath = Split-Path -Path $MyInvocation.MyCommand.Definition -Parent

# Construct the path to the Modules folder
$modulesPath = Join-Path -Path $scriptPath -ChildPath "Modules"

# Dot source each module script
. (Join-Path -Path $modulesPath -ChildPath "config.ps1")
. (Join-Path -Path $modulesPath -ChildPath "forms.ps1")
. (Join-Path -Path $modulesPath -ChildPath "datatables_variables.ps1")
. (Join-Path -Path $modulesPath -ChildPath "word_insertions.ps1")
. (Join-Path -Path $modulesPath -ChildPath "excel_insertions.ps1")
