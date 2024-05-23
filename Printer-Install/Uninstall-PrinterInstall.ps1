# Define the path to the text file
$textFilePath = "C:\ProgramData\FDS\PrinterInstall\PrinterUninstall.txt"

# Check if the file exists
if (Test-Path $textFilePath -PathType Leaf) {
    # Read the content of the text file
    $printerNames = Get-Content -Path $textFilePath

    if ($printerNames) {
        $printersUninstalled = @()
        $printersNotFound = @()

        foreach ($printerName in $printerNames) {
            # Trim any leading or trailing whitespace
            $printerName = $printerName.Trim()

            if ($printerName) {
                # Check if the printer is installed
                $installedPrinter = Get-Printer | Where-Object { $_.Name -eq $printerName }

                if ($installedPrinter) {
                    try {
                        # Remove the printer
                        Remove-Printer -Name $printerName
                        $printersUninstalled += $printerName
                        Write-Output "Printer '$printerName' uninstalled successfully."
                    } catch {
                        Write-Output "Failed to uninstall printer '$printerName'. Error: $_"
                    }
                } else {
                    $printersNotFound += $printerName
                    Write-Output "Printer '$printerName' not found."
                }
            }
        }

        if ($printersUninstalled) {
            Write-Output "The following printers were uninstalled: $($printersUninstalled -join ', ')"
        }

        if ($printersNotFound) {
            Write-Output "The following printers were not found: $($printersNotFound -join ', ')"
        }

        # Return a success exit code (0) if at least one printer was uninstalled, otherwise a failure exit code (1)
        if ($printersUninstalled.Count -gt 0) {
            exit 0
        } else {
            exit 1
        }
    } else {
        Write-Output "The file '$textFilePath' is empty or could not read the printer names."
        # Return a failure exit code (1)
        exit 1
    }
} else {
    Write-Output "File '$textFilePath' not found."
    # Return a failure exit code (1)
    exit 1
}
