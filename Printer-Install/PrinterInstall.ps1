<#
.DESCRIPTION
    This script checks the endpoint for the current IP address and installs the associated printer with its properties.

.INPUTS
    Installation from Company Portal..

.OUTPUTS
    Intune scripts are executed to install printers.

.NOTES
    Version:        10.0
    Author:         Chris Braeuer
    Creation Date:  5/17/2024

.RELEASE NOTES
    Version 1.0 (5/17/2024):
    - Initial version of the script written.
    Version 2.0 (5/17/2024
#>

$wgetPath = Get-Command wget -ErrorAction SilentlyContinue

if (-not $wgetPath) {
    Write-Host "wget is not installed. Installing wget using Chocolatey..."
    try {
        Invoke-Expression ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))
        choco install wget -y
    }
    catch {
        Write-Error "Failed to install wget: $_"
        exit 1
    }
}

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
    @{ Name = 'Intune Printer'; IPAddress = '172.30.125.200' }
)

# Define the GitHub URLs for the files
$msiUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/blob/main/Printer-Install/6900.msi"
$modelDatUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/blob/main/Printer-Install/model023.dat"
$ImageURL = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/blob/main/Printer-Install/FDSLogo.png"

# Define the directory and temp paths
$directoryPath = "C:\ProgramData\FDS"
$tempDirectoryPath = Join-Path $directoryPath "temp"
$tempMsiPath = Join-Path $tempDirectoryPath "6900.msi"
$tempModelDatPath = Join-Path $tempDirectoryPath "model023.dat"
$tempImagePath = Join-Path $tempDirectoryPath "FDSLogo.png"

# Ensure the temp directory exists
if (-not (Test-Path $tempDirectoryPath)) {
    try {
        New-Item -ItemType Directory -Path $tempDirectoryPath -Force | Out-Null
    }
    catch {
        Write-Warning "Failed to create directory: $tempDirectoryPath"
        Write-Warning $_.Exception.Message
    }
}

# Download the MSI file from GitHub to the temp directory
$wgetArgs = @("-O", $tempMsiPath, $msiUrl)
& wget $wgetArgs

# Download the model023.dat file from GitHub to the temp directory
$wgetArgs = @("-O", $tempModelDatPath, $modelDatUrl)
& wget $wgetArgs

# Download the FDSLogo.png file from GitHub to the temp directory
$wgetArgs = @("-O", $tempImagePath, $ImageURL)
& wget $wgetArgs

# Define the Show-PopupMessageWithImage function
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

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(150, 220)
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $form.Controls.AddRange(@($pictureBox, $label, $okButton))
    $form.AcceptButton = $okButton

    $form.ShowDialog() | Out-Null
}

# Define a template for the command
$commandTemplate = '/i "{0}" /quiet DRIVERNAME="Brother MFC-L6900DW series" PRINTERNAME="{1}" ISDEFAULTPRINTER="0" IPADDRESS="{2}" /qn /NORESTART'

# Get the current IP address of the machine
$CurrentIPAddress = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.InterfaceAlias -notlike '*Loopback*' }).IPAddress

# Truncate the current IP address to the first three octets
$TruncatedCurrentIPAddress = $CurrentIPAddress -replace '\.\d+$'

# Iterate through the array to find a matching IP address and execute the command
$matched = $false
foreach ($config in $printerConfigs) {
    # Truncate the IP address from the configurations to the first three octets
    $TruncatedConfigIPAddress = $config.IPAddress -replace '\.\d+$'
    if ($TruncatedCurrentIPAddress -eq $TruncatedConfigIPAddress) {
        $logFilePath = Join-Path $tempDirectoryPath "printer-install.log"

        $arguments = $commandTemplate -f $tempMsiPath, $config.Name, $config.IPAddress
        Write-Host "Executing command: msiexec $arguments"
        $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow -PassThru
        $process.WaitForExit()

        if ($process.ExitCode -ne 0) {
            $errorMessage = "Failed to install printer $($config.Name). Exit code: $($process.ExitCode)"
            Write-Warning $errorMessage
            Add-Content -Path $logFilePath -Value $errorMessage
            Add-Content -Path $logFilePath -Value $process.StandardOutput
            Add-Content -Path $logFilePath -Value $process.StandardError
        } else {
            $message = "The $($config.Name) printer was installed."
            Write-Host $message
            Show-PopupMessageWithImage $message "Printer Installed" $tempImagePath
        }
        $matched = $true
        break
    }
}

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
