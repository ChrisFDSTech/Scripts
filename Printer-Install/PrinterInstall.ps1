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

# Enhanced logging function
function Write-Log {
    [CmdletBinding()]
    param(
         [Parameter(Mandatory=$True)]
         [ValidateNotNullOrEmpty()]
         [string]$Source,
 
         [Parameter()]
         [ValidateNotNullOrEmpty()]
         [int]$EventId = 0,
                         
         [Parameter(Mandatory=$True)]
         [ValidateNotNullOrEmpty()]
         [string]$Message,
  
         [Parameter(Mandatory=$True)]
         [ValidateNotNullOrEmpty()]
         [ValidateSet('Information','Warning','Error')]
         [string]$EntryType = 'Information'
     )
 
    $LogFolder = "$($env:ProgramData)\IntuneLogging"
    $LogFile = "$($LogFolder )\$($Source).log"
    if(-not (Test-Path $LogFolder)) {
        New-Item -Path $LogFolder -ItemType directory
    }
    if(-not (Test-Path $LogFile)) {
        Add-Content -Path $LogFile -Value "Date;Level;Message" -ErrorAction SilentlyContinue
    }
    Add-Content -Path $LogFile -Value "$(Get-Date -Format "yyyy-MM-dd HH:mm:ss");$($EntryType);$($Message)" -ErrorAction SilentlyContinue
}

Write-Log -Source "Script" -Message "Script execution started" -EntryType 'Information'

# Define the GitHub URLs for the files
$msiUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/6900.msi"
$modelDatUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/model023.dat"
$imageUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/FDSLogo.ico" 

# Define the GitHub URLs for the files
$msiUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/6900.msi"
$modelDatUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/model023.dat"

# Define the directory and temp paths
$directoryPath = "C:\Program Files\FDS"
$tempDirectoryPath = "$directoryPath\temp"
$tempMsiPath = [System.IO.Path]::Combine($tempDirectoryPath, "6900.msi")
$tempModelDatPath = [System.IO.Path]::Combine($tempDirectoryPath, "model023.dat")

# Ensure the temp directory exists
if (-Not (Test-Path -Path $tempDirectoryPath)) {
    New-Item -ItemType "Directory" -Path "$tempDirectoryPath" -Force | Out-Null
}

# Download the MSI file from GitHub to the temp directory with error handling
try {
    Invoke-WebRequest -Uri $msiUrl -OutFile $tempMsiPath -UseBasicParsing -ErrorAction Stop
}
catch {
    Write-Error "Failed to download MSI file: $_"
    exit 1
}

# Download the MSI file from GitHub to the temp directory with error handling
try {
    Write-Log -Source "Script" -Message "Downloading MSI file from $msiUrl to $tempMsiPath" -EntryType 'Information'
    Invoke-WebRequest -Uri $msiUrl -OutFile $tempMsiPath -ErrorAction Stop
}
catch {
    Write-Log -Source "Script" -Message "Failed to download MSI file: $_" -EntryType 'Error'
    exit 1
}

# Download the model023.dat file from GitHub to the temp directory with error handling
try {
    Write-Log -Source "Script" -Message "Downloading model023.dat file from $modelDatUrl to $tempModelDatPath" -EntryType 'Information'
    Invoke-WebRequest -Uri $modelDatUrl -OutFile $tempModelDatPath -ErrorAction Stop
}
catch {
    Write-Log -Source "Script" -Message "Failed to download model023.dat file: $_" -EntryType 'Error'
    exit 1
}

# Download the image file from GitHub to the temp directory with error handling
try {
    Write-Log -Source "Script" -Message "Downloading image file from $imageUrl to $tempImagePath" -EntryType 'Information'
    Invoke-WebRequest -Uri $imageUrl -OutFile $tempImagePath -ErrorAction Stop
}
catch {
    Write-Log -Source "Script" -Message "Failed to download image file: $_" -EntryType 'Error'
    exit 1
}

# Define a template for the command
$commandTemplate = 'msiexec /i "{0}" /q DRIVERNAME="Brother MFC-L6900DW series" PRINTERNAME="{1}" ISDEFAULTPRINTER="0" IPADDRESS="{2}"'

# Get the current IP address of the machine
$CurrentIPAddress = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.InterfaceAlias -notlike '*Loopback*' }).IPAddress


# Define an array of hashtables with the printer configurations
$printerConfigs = @(
    @{ Name = 'Azalea Manor'; IPAddress = '192.168.14.200' },
    @{ Name = 'Bennettsville Green'; IPAddress = '192.168.5.200' },
    @{ Name = 'Benson Green'; IPAddress = '192.168.192.200' },
    @{ Name = 'Blanton Green'; IPAddress = '192.168.41.200' },
    @{ Name = 'Bordeaux'; IPAddress = '192.168.40.200' },
    @{ Name = 'Bunce Green'; IPAddress = '192.168.43.200' },
    @{ Name = 'Bunce Manor'; IPAddress = '192.168.8.200' },
    @{ Name = 'Cleveland Green III'; IPAddress = '192.168.71.200' },
    @{ Name = 'Clinton Green'; IPAddress = '192.168.20.200' },
    @{ Name = 'Club Pond'; IPAddress = '192.168.104.200' },
    @{ Name = 'Coldwater Ridge'; IPAddress = '192.168.200.200' },
    @{ Name = 'Cross Creek Pointe'; IPAddress = '192.168.33.200' },
    @{ Name = 'Crosswinds Green'; IPAddress = '192.168.66.200' },
    @{ Name = 'Cypress Manor'; IPAddress = '192.168.68.200' },
    @{ Name = 'Dogwood Manor'; IPAddress = '192.168.10.200' },
    @{ Name = 'Eastside Green'; IPAddress = '192.168.46.200' },
    @{ Name = 'Golfview'; IPAddress = '192.168.47.200' },
    @{ Name = 'Graham Manor'; IPAddress = '192.168.194.200' },
    @{ Name = 'Haymount Manor'; IPAddress = '192.168.48.200' },
    @{ Name = 'Hickory Ridge'; IPAddress = '192.168.11.200' },
    @{ Name = 'Hoke Loop'; IPAddress = '192.168.73.200' },
    @{ Name = 'Legion Crossing'; IPAddress = '192.168.54.200' },
    @{ Name = 'Legion Manor'; IPAddress = '192.168.50.200' },
    @{ Name = 'Longview'; IPAddress = '192.168.51.200' },
    @{ Name = 'McArthur Park'; IPAddress = '192.168.89.200' },
    @{ Name = 'Millstone Landing'; IPAddress = '192.168.111.200' },
    @{ Name = 'Newberry Green'; IPAddress = '192.168.52.200' },
    @{ Name = 'Oak Run'; IPAddress = '192.168.13.200' },
    @{ Name = 'Oak Run II'; IPAddress = '192.168.15.200' },
    @{ Name = 'Palmer Green'; IPAddress = '192.168.53.200' },
    @{ Name = 'Raeford Green'; IPAddress = '192.168.42.200' },
    @{ Name = 'Reidsville Ridge'; IPAddress = '192.168.72.200' },
    @{ Name = 'Riverview Green'; IPAddress = '192.168.44.200' },
    @{ Name = 'Rosehill West'; IPAddress = '192.168.88.200' },
    @{ Name = 'Shallotte Villas'; IPAddress = '192.168.58.200' },
    @{ Name = 'Southview Green'; IPAddress = '192.168.59.200' },
    @{ Name = 'Southview Townhouses'; IPAddress = '192.168.67.200' },
    @{ Name = 'Southview Villas'; IPAddress = '192.168.45.200' },
    @{ Name = 'Spring Lake Green'; IPAddress = '192.168.61.200' },
    @{ Name = 'Sycamore Park'; IPAddress = '192.168.12.200' },
    @{ Name = 'Tokay Green'; IPAddress = '192.168.49.200' },
    @{ Name = 'Wallstreet Green'; IPAddress = '192.168.120.200' },
    @{ Name = 'Watauga Green'; IPAddress = '192.168.69.200' },
    @{ Name = 'West Cumberland'; IPAddress = '192.168.92.200' },
    @{ Name = 'West Fayetteville'; IPAddress = '192.168.204.200' },
    @{ Name = 'Woodgreen'; IPAddress = '192.168.64.200' },
    @{ Name = 'Zebulon Green'; IPAddress = '192.168.103.200' },
    @{ Name = 'TEST PRINT'; IPAddress = '172.30.125.200' }
)

# Extract the first three octets of the current IP address
$CurrentIPOctets = $CurrentIPAddress -replace '^(\d+\.\d+\.\d+)\.\d+$', '$1'
Write-Log -Source "Script" -Message "Current IP Address: $CurrentIPAddress" -EntryType 'Information'
Write-Log -Source "Script" -Message "Extracted Octets: $CurrentIPOctets" -EntryType 'Information'

# Iterate through the array to find a matching IP address (based on first three octets) and execute the command
$matched = $false
foreach ($config in $printerConfigs) {
    $ConfigIPOctets = $config.IPAddress -replace '^(\d+\.\d+\.\d+)\.\d+$', '$1'
    Write-Log -Source "Script" -Message "Checking printer config: $($config.Name) with IP Octets: $ConfigIPOctets" -EntryType 'Information'
    if ($CurrentIPOctets -eq $ConfigIPOctets) {
        Write-Log -Source "Script" -Message "Found matching IP address: $CurrentIPAddress (based on first three octets)" -EntryType 'Information'
        $command = [string]::Format($commandTemplate, $tempMsiPath, $config.Name, $config.IPAddress)
        try {
            Write-Log -Source "Script" -Message "Executing command: $command" -EntryType 'Information'
            Invoke-Expression $command
            $matched = $true
            Write-Log -Source "Script" -Message "Printer $($config.Name) installed successfully" -EntryType 'Information'
        } catch {
            Write-Log -Source "Script" -Message "Failed to install printer: $_" -EntryType 'Error'
        }
        break
    }
}

# Define the function to display a pop-up message with an image
function Show-PopupMessageWithImage {
    param(
        [string]$Message,
        [string]$Title,
        [string]$ImagePath
    )
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.Size = New-Object System.Drawing.Size(400, 300)
    $form.StartPosition = "CenterScreen"

    $pictureBox = New-Object System.Windows.Forms.PictureBox
    $pictureBox.Size = New-Object System.Drawing.Size(100, 100)
    $pictureBox.Location = New-Object System.Drawing.Point(150, 20)
    $pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::StretchImage
    $pictureBox.ImageLocation = $ImagePath

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Message
    $label.AutoSize = $true
    $label.Location = New-Object System.Drawing.Point(50, 150)
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter

    $form.Controls.Add($pictureBox)
    $form.Controls.Add($label)

    $form.ShowDialog()
}

# If a matching IP address was found and printer installed, show a popup message with the printer name
if ($matched) {
    $printerName = $config.Name
    $message = "The $printerName printer was installed."
    Write-Log -Source "Script" -Message "Displaying popup: $message" -EntryType 'Information'
    Show-PopupMessageWithImage $message "Printer Installed" $tempImagePath
}
else {
    # If no matching IP address was found, show a popup message with error
    Write-Error "No printer configuration found for the current IP address: $CurrentIPAddress"
}

# Clean up: Delete the temp directory
if (Test-Path -Path $tempDirectoryPath) {
    Write-Log -Source "Script" -Message "Cleaning up temp directory at $tempDirectoryPath" -EntryType 'Information'
    Remove-Item -Path $tempDirectoryPath -Recurse -Force
}
