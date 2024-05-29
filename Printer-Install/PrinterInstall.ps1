# Part 1: UI Setup and Helper Functions

# Load the WPF assemblies
Add-Type -AssemblyName PresentationFramework, PresentationCore

# Define the XAML markup for the window
[xml]$xaml = @"
<Window
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="Printer Installation" Height="400" Width="600" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Grid>
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="*"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>

        <Label Grid.Row="0" Content="Printer Installation" FontSize="24" FontWeight="Bold" HorizontalAlignment="Center" Margin="10"/>

        <TextBox Grid.Row="1" Name="logTextBox" TextWrapping="Wrap" VerticalScrollBarVisibility="Auto" Margin="10" IsReadOnly="True"/>

        <ProgressBar Grid.Row="2" Name="progressBar" Height="20" Margin="10" IsIndeterminate="False"/>
    </Grid>
</Window>
"@

# Create the window from the XAML markup
try {
    $window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
}
catch {
    Write-Error "Failed to create WPF window: $_"
    return
}

# Function to update the log text box
function Update-LogTextBox($message) {
    if ($window -ne $null) {
        try {
            $window.Dispatcher.Invoke([Action]{
                $window.logTextBox.AppendText("$message`n")
                $window.logTextBox.ScrollToEnd()
            }, "Normal")
        }
        catch {
            Write-Error "Failed to update log text box: $_"
        }
    }
    else {
        Write-Output $message
    }
}

# Function to update the progress bar
function Update-ProgressBar($value) {
    if ($window -ne $null) {
        try {
            $window.Dispatcher.Invoke([Action]{
                $window.progressBar.Value = $value
            }, "Normal")
        }
        catch {
            Write-Error "Failed to update progress bar: $_"
        }
    }
}

# Function to show the installation window
function Show-InstallWindow($printerName) {
    if ($window -ne $null) {
        $window.Title = "Installing $printerName"
        $null = $window.Dispatcher.InvokeAsync({
            $window.ShowDialog()
        }).AsTask().Result
    }
    else {
        Write-Error "Window object is null. Unable to show installation window."
    }
}

# Function to show a success notification
function Show-SuccessNotification($printerName) {
    Update-LogTextBox "The $printerName printer was installed successfully."
}

# Function to show a failure notification
function Show-FailureNotification($printerName) {
    Update-LogTextBox "Failed to install printer $printerName. Please call support at 910-483-5395."
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
    @{ Name = 'TEST Green'; IPAddress = '172.30.125.202' }
)

# Function to update the PrintersInstalled.txt and PrinterUninstall.txt files
function UpdatePrinterLogFiles($printerName) {
    if (Test-Path $printerInstalledLogPath) {
        $existingContent = Get-Content $printerInstalledLogPath
        Set-Content -Path $printerInstalledLogPath -Value ($printerName, $existingContent)
    } else {
        Set-Content -Path $printerInstalledLogPath -Value $printerName
    }

    if (Test-Path $printerUninstallLogPath) {
        $existingContent = Get-Content $printerUninstallLogPath
        Set-Content -Path $printerUninstallLogPath -Value ($printerName, $existingContent)
    } else {
        Set-Content -Path $printerUninstallLogPath -Value $printerName
    }
}

# Define the GitHub URLs for the files
$msiUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/6900.msi"
$modelDatUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/model023.dat"
$ImageURL = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/FDSLogo.png"
$schedTaskUrl = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/Printer-Install/ScheduleTask.ps1"

# Define the directory and temp paths
$directoryPath = "C:\ProgramData\FDS"
$tempDirectoryPath = Join-Path $directoryPath "PrinterInstall"
$tempMsiPath = Join-Path $tempDirectoryPath "6900.msi"
$tempModelDatPath = Join-Path $tempDirectoryPath "model023.dat"
$tempImagePath = Join-Path $tempDirectoryPath "FDSLogo.png"
$schedTaskPath = Join-Path $tempDirectoryPath "ScheduleTask.ps1"
$printerInstalledLogPath = Join-Path $tempDirectoryPath "PrintersInstalled.txt"
$printerUninstallLogPath = Join-Path $tempDirectoryPath "PrinterUninstall.txt"

# Ensure the temp directory exists
if (-not (Test-Path $directoryPath)) {
    try {
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }
    catch {
        Write-Warning "Failed to create directory: $directoryPath"
        Write-Warning $_.Exception.Message
    }
}

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

# Download the MSI file, .dat File, and the Logo file from GitHub to the temp directory
Update-LogTextBox "Downloading required files..."
Invoke-WebRequest -Uri $msiUrl -OutFile $tempMsiPath
Invoke-WebRequest -Uri $modelDatUrl -OutFile $tempModelDatPath
Invoke-WebRequest -Uri $imageURL -OutFile $tempImagePath
Invoke-WebRequest -Uri $schedTaskUrl -OutFile $schedTaskPath

# Define the printer driver name
$driverName = "Brother MFC-L6900DW series"

# Check if the printer driver is installed
try {
    $printerDriver = Get-PrinterDriver -Name $driverName -ErrorAction Stop
    if ($printerDriver) {
        Update-LogTextBox "The printer driver '$driverName' is already installed."
    }
}
catch {
    Update-LogTextBox "The printer driver '$driverName' is not installed. Installing the driver and printer..."

    # Define a template for the command (same as the old script)
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
            Update-LogTextBox "Executing command: msiexec $arguments"
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            $process.WaitForExit()

            if ($process.ExitCode -ne 0) {
                $errorMessage = "Failed to install printer $($config.Name). Exit code: $($process.ExitCode)"
                Update-LogTextBox $errorMessage
                Add-Content -Path $logFilePath -Value $errorMessage
                Add-Content -Path $logFilePath -Value $process.StandardOutput
                Add-Content -Path $logFilePath -Value $process.StandardError

                # Show the failure notification
                Show-FailureNotification -printerName $config.Name

            } else {
                $message = "The $($config.Name) printer was installed."
                Update-LogTextBox $message

                # Update the PrintersInstalled.txt and PrinterUninstall.txt files
                UpdatePrinterLogFiles $config.Name

                # Show the success notification
                Show-SuccessNotification -printerName $config.Name
            }

            $matched = $true
            break
        }
    }

    # If no matching IP address was found, log the error
    if (-not $matched) {
        $errorMessage = "No matching IP address found for printer installation."
        Update-LogTextBox $errorMessage

        # Show the failure notification
        Show-FailureNotification -printerName $config.Name

        Add-Content -Path $logFilePath -Value $errorMessage
    }
}

# If the driver is already installed, proceed with adding the printer
if ($printerDriver) {
    # Get the current IP address of the machine
    $CurrentIPAddress = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.InterfaceAlias -notlike '*Loopback*' }).IPAddress

    # Truncate the current IP address to the first three octets
    $TruncatedCurrentIPAddress = $CurrentIPAddress -replace '\.\d+$'

    # Iterate through the array to find a matching IP address and add the printer
    $matched = $false
    foreach ($config in $printerConfigs) {
        # Truncate the IP address from the configurations to the first three octets
        $TruncatedConfigIPAddress = $config.IPAddress -replace '\.\d+$'
        if ($TruncatedCurrentIPAddress -eq $TruncatedConfigIPAddress) {
            # Check if the printer port already exists
            $portName = "IP_$($config.IPAddress)"
            $existingPort = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue

            if (-not $existingPort) {
                # Add the printer port
                Update-LogTextBox "Adding printer port: $portName"
                Add-PrinterPort -Name $portName -PrinterHostAddress "$($config.IPAddress)"
            }

            # Add the printer
            Update-LogTextBox "Adding printer: $($config.Name)"
            Add-Printer -Name $config.Name -DriverName "Brother MFC-L6900DW series" -PortName $portName

            $message = "The $($config.Name) printer was installed."
            Update-LogTextBox $message

            # Update the PrintersInstalled.txt and PrinterUninstall.txt files
            UpdatePrinterLogFiles $config.Name

            # Show the success notification
            Show-SuccessNotification -printerName $config.Name

            $matched = $true
            break
        }
    }

    # If no matching IP address was found, log the error
    if (-not $matched) {
        $errorMessage = "No matching IP address found for printer installation."
        Update-LogTextBox $errorMessage

        # Show the failure notification
        Show-FailureNotification -printerName $config.Name

        Add-Content -Path $logFilePath -Value $errorMessage
    }
}

try {
    Update-LogTextBox "Executing scheduled task script..."
    $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processStartInfo.FileName = "powershell.exe"
    $processStartInfo.Arguments = "-ExecutionPolicy Bypass -File `"$schedTaskPath`""
    $processStartInfo.Verb = "RunAs"

    $process = [System.Diagnostics.Process]::Start($processStartInfo)
    $process.WaitForExit()

    if ($process.ExitCode -ne 0) {
        Update-LogTextBox "Failed to execute the scheduled task script with exit code: $($process.ExitCode)"
    }
}
catch {
    Update-LogTextBox "Failed to execute the scheduled task script: $_"
}

# Start the script in an interactive mode
$null = $window.Dispatcher.InvokeAsync({
    # Initiate the installation process
    $installJob = Start-Installation -printerName "Brother MFC-L6900DW"
}).AsTask().Result

# Keep the script running until the installation is complete
$installJob | Wait-Job

# Close the window
$window.Close()
