# Define the list of printer names to check for
$printerNames = @(
    'Azalea Manor',
    'Bennettsville Green',
    'Benson Green',
    'Blanton Green',
    'Bordeaux',
    'Bunce Green',
    'Bunce Manor',
    'Cleveland Green III',
    'Clinton Green',
    'Club Pond',
    'Coldwater Ridge',
    'Cross Creek Pointe',
    'Crosswinds Green',
    'Cypress Manor',
    'Dogwood Manor',
    'Eastside Green',
    'Golfview',
    'Graham Manor',
    'Haymount Manor',
    'Hickory Ridge',
    'Hoke Loop',
    'Legion Crossing',
    'Legion Manor',
    'Longview',
    'McArthur Park',
    'Millstone Landing',
    'Newberry Green',
    'Oak Run',
    'Oak Run II',
    'Palmer Green',
    'Raeford Green',
    'Reidsville Ridge',
    'Riverview Green',
    'Rosehill West',
    'Shallotte Villas',
    'Southview Green',
    'Southview Townhouses',
    'Southview Villas',
    'Spring Lake Green',
    'Sycamore Park',
    'Tokay Green',
    'Wallstreet Green',
    'Watauga Green',
    'West Cumberland',
    'West Fayetteville',
    'Woodgreen',
    'Zebulon Green',
    'Intune Printer'
)

# Get a list of all installed printers
$installedPrinters = Get-Printer -Full

# Loop through each printer name in the list
foreach ($printerName in $printerNames) {
    # Check if the printer is installed
    $printer = $installedPrinters | Where-Object { $_.Name -eq $printerName }

    if ($printer) {
        try {
            # Uninstall the printer
            $null = $printer.RemovePrinter()
            Write-Host "Printer '$printerName' has been uninstalled."
        }
        catch {
            Write-Warning "Failed to uninstall printer '$printerName': $_"
        }
    }
    else {
        Write-Host "Printer '$printerName' is not installed."
    }
}
