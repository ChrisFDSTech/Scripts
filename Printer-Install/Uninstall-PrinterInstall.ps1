# Define the directory path
$directoryPath = "C:\ProgramData\FDS"
$printersFilePath = Join-Path $directoryPath "Printers.txt"

# Check if the Printers.txt file exists
if (Test-Path $printersFilePath) {
    # Read the printer names from the file
    $printerNames = Get-Content $printersFilePath

    # Loop through each printer name and uninstall the printer
    foreach ($printerName in $printerNames) {
        try {
            # Get the printer object
            $printer = Get-Printer -Name $printerName -ErrorAction Stop

            # Uninstall the printer
            $null = $printer.RemovePrinter()
            Write-Host "Printer '$printerName' has been uninstalled."
        }
        catch {
            Write-Warning "Failed to uninstall printer '$printerName': $_"
        }
    }
}
else {
    Write-Warning "The Printers.txt file does not exist in $directoryPath"
}
