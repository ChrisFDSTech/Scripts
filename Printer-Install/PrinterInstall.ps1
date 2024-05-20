<#
.DESCRIPTION
    This script checks the endpoint for the current IP address and installs the associated printer with its properties.

.INPUTS
    Installation from Company Portal.

.OUTPUTS
    Intune scripts are executed to install printers.

.NOTES
    Version:        1.0
    Author:         Chris Braeuer
    Creation Date:  5/17/2024

.RELEASE NOTES
    Version 1.0 (5/17/2024):
    - Initial version of the script written.
#>


# Define an array of hashtables with the printer configurations
$printerConfigs = @(
    @{ Name = 'Azalea Manor'; IPAddress = '192.168.14.69' },
    @{ Name = 'Bennettsville Green'; IPAddress = '192.168.5.69' },
    @{ Name = 'Benson Green'; IPAddress = '192.168.192.69' },
    @{ Name = 'Blanton Green'; IPAddress = '192.168.41.69' },
    @{ Name = 'Bordeaux'; IPAddress = '192.168.40.69' },
    @{ Name = 'Bunce Green'; IPAddress = '192.168.43.69' },
    @{ Name = 'Bunce Manor'; IPAddress = '192.168.8.69' },
    @{ Name = 'Cleveland Green III'; IPAddress = '192.168.71.69' },
    @{ Name = 'Clinton Green'; IPAddress = '192.168.20.69' },
    @{ Name = 'Club Pond'; IPAddress = '192.168.104.69' },
    @{ Name = 'Coldwater Ridge'; IPAddress = '192.168.200.69' },
    @{ Name = 'Cross Creek Pointe'; IPAddress = '192.168.33.69' },
    @{ Name = 'Crosswinds Green'; IPAddress = '192.168.66.69' },
    @{ Name = 'Cypress Manor'; IPAddress = '192.168.68.69' },
    @{ Name = 'Dogwood Manor'; IPAddress = '192.168.10.69' },
    @{ Name = 'Eastside Green'; IPAddress = '192.168.46.69' },
    @{ Name = 'Golfview'; IPAddress = '192.168.47.69' },
    @{ Name = 'Graham Manor'; IPAddress = '192.168.194.69' },
    @{ Name = 'Haymount Manor'; IPAddress = '192.168.48.69' },
    @{ Name = 'Hickory Ridge'; IPAddress = '192.168.11.69' },
    @{ Name = 'Hoke Loop'; IPAddress = '192.168.73.69' },
    @{ Name = 'Legion Crossing'; IPAddress = '192.168.54.69' },
    @{ Name = 'Legion Manor'; IPAddress = '192.168.50.69' },
    @{ Name = 'Longview'; IPAddress = '192.168.51.69' },
    @{ Name = 'McArthur Park'; IPAddress = '192.168.89.69' },
    @{ Name = 'Millstone Landing'; IPAddress = '192.168.111.69' },
    @{ Name = 'Newberry Green'; IPAddress = '192.168.52.69' },
    @{ Name = 'Oak Run'; IPAddress = '192.168.13.69' },
    @{ Name = 'Oak Run II'; IPAddress = '192.168.15.69' },
    @{ Name = 'Palmer Green'; IPAddress = '192.168.53.69' },
    @{ Name = 'Raeford Green'; IPAddress = '192.168.42.69' },
    @{ Name = 'Reidsville Ridge'; IPAddress = '192.168.72.69' },
    @{ Name = 'Riverview Green'; IPAddress = '192.168.44.69' },
    @{ Name = 'Rosehill West'; IPAddress = '192.168.88.69' },
    @{ Name = 'Shallotte Villas'; IPAddress = '192.168.58.69' },
    @{ Name = 'Southview Green'; IPAddress = '192.168.59.69' },
    @{ Name = 'Southview Townhouses'; IPAddress = '192.168.67.69' },
    @{ Name = 'Southview Villas'; IPAddress = '192.168.45.69' },
    @{ Name = 'Spring Lake Green'; IPAddress = '192.168.61.69' },
    @{ Name = 'Sycamore Park'; IPAddress = '192.168.12.69' },
    @{ Name = 'Tokay Green'; IPAddress = '192.168.49.69' },
    @{ Name = 'Wallstreet Green'; IPAddress = '192.168.120.69' },
    @{ Name = 'Watauga Green'; IPAddress = '192.168.69.69' },
    @{ Name = 'West Cumberland'; IPAddress = '192.168.92.69' },
    @{ Name = 'West Fayetteville'; IPAddress = '192.168.204.69' },
    @{ Name = 'Woodgreen'; IPAddress = '192.168.64.69' },
    @{ Name = 'Zebulon Green'; IPAddress = '192.168.103.69' }
)

# Define the GitHub URLs for the files
$msiUrl = "https://github.com/ChrisFDSTech/Scripts/blob/main/Printer-Install/6900.msi"
$modelDatUrl = "https://github.com/ChrisFDSTech/Scripts/blob/main/Printer-Install/model023.dat"

# Define the directory and temp paths
$directoryPath = "C:\Program Files\FDS"
$tempDirectoryPath = "$directoryPath\temp"
$tempMsiPath = [System.IO.Path]::Combine($tempDirectoryPath, "6900.msi")
$tempModelDatPath = [System.IO.Path]::Combine($tempDirectoryPath, "model023.dat")

# Ensure the temp directory exists
if (-Not (Test-Path -Path $tempDirectoryPath)) {
    New-Item -ItemType Directory -Path $tempDirectoryPath -Force
}

# Download the MSI file from GitHub to the temp directory
Invoke-WebRequest -Uri $msiUrl -OutFile $tempMsiPath

# Download the model023.dat file from GitHub to the temp directory
Invoke-WebRequest -Uri $modelDatUrl -OutFile $tempModelDatPath

# Define a template for the command
$commandTemplate = 'msiexec /i "{0}" /q DRIVERNAME="Brother MFC-L6900DW series" PRINTERNAME="{1}" ISDEFAULTPRINTER="0" IPADDRESS="{2}"'

# Get the current IP address of the machine
$CurrentIPAddress = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.InterfaceAlias -notlike '*Loopback*' }).IPAddress

# Iterate through the array to find a matching IP address and execute the command
$matched = $false
foreach ($config in $printerConfigs) {
    if ($CurrentIPAddress -eq $config.IPAddress) {
        $command = [string]::Format($commandTemplate, $tempMsiPath, $config.Name, $config.IPAddress)
        Invoke-Expression $command
        $matched = $true
        break
    }
}

# Clean up: Delete the temp directory
if (Test-Path -Path $tempDirectoryPath) {
    Remove-Item -Path $tempDirectoryPath -Recurse -Force
}


# Define the directory and GitHub URL for the image
$directoryPath = "C:\Program Files\FDS"
$tempDirectoryPath = "$directoryPath\temp"
$githubImageUrl = "https://github.com/ChrisFDSTech/Scripts/blob/main/Printer-Install/FDSLogo.ico"
$tempImagePath = [System.IO.Path]::Combine($tempDirectoryPath, "FDSLogo.ico")

# Ensure the temp directory exists
if (-Not (Test-Path -Path $tempDirectoryPath)) {
    New-Item -ItemType Directory -Path $tempDirectoryPath -Force
}

# Download the image from GitHub
Invoke-WebRequest -Uri $githubImageUrl -OutFile $tempImagePath

# If no matching IP address was found, show a popup message
if (-not $matched) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Printer Setup Issue"
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"

    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(100, 100)
    $pictureBox.Location = New-Object System.Drawing.Point(150, 20)
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $pictureBox.ImageLocation = $tempImagePath

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "There was an issue with installing the printer, please call support at 910-483-5395"
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(50, 150)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    $form.Controls.Add($pictureBox)
    $form.Controls.Add($label)

    $form.ShowDialog()
}

# Clean up: Delete the temp directory
if (Test-Path -Path $tempDirectoryPath) {
    Remove-Item -Path $tempDirectoryPath -Recurse -Force
}
