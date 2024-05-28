<#
.DESCRIPTION
    This script checks the endpoint for the current IP address and installs the associated printer with its properties.

.INPUTS
    Installation from Company Portal..

.OUTPUTS
    Intune scripts are executed to install printers.

.NOTES
    Version:        5.0
    Author:         Chris Braeuer
    Creation Date:  5/17/2024

.RELEASE NOTES
    Version 1.0 (5/17/2024):
    - Initial version of the script written.
    
    Version 2.0 (5/20/2024
    - Fixed Issue with logo (Changed from .ico to .png)
    - Changed the way the downloads were retrieved with WindowsInstallModule
    - Added a popup windows for the printer install

    Version 3.0 (5/21/2024)
    - Changed to Invoke-WebRequest cmdlet
    - Created a more appealing popup window
    - Fixed issue with exit code: 1619
    - Fixed popup window with ShowDiag
    - removed WGET module

    Version 4.0 (5/22/2024)
    - Added check for printer driver installation
    - If driver is installed, use the existing driver instead of attempting to install it again
    - Removed code to download MSI and DAT files
    - Added logic to download and install the printer driver if it's not already installed
    - Updated script to use Add-PrinterDriver and Add-Printer cmdlets instead of rundll32.exe
    - Fixed issues with downloading the DAT file and adding the printer port
    - Added logic to download and install the printer driver files (MSI and DAT) if the driver is not already installed
    - Updated script to use pnputil.exe to install the printer driver
    - Added a check to skip adding the printer port if it already exists
    - Capture and display the error output from pnputil.exe when installing the printer driver
    - Capture and display the error output from pnputil.exe when installing the printer driver
    
    Version 5.0 (5/25/2024)
    - Improved the logic for installing the printer driver and printer
    - If the printer driver is not installed, the script now uses the same logic as the previous version to install the driver and printer using the `msiexec` command
    - If the printer driver is already installed, the script proceeds with the new logic to check for the printer port, add the printer port if it doesn't exist, and then add the printer
    - Added better error handling and logging for failed printer installations
    - Retained the existing functionality to display a popup message with the FDS logo upon successful printer installation or if there's an issue with the installation

#>

If ($PSVersionTable.PSVersion -ge [version]"5.0" -and (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full\').Release -ge 379893) {

    If ([Net.ServicePointManager]::SecurityProtocol -ne [Net.SecurityProtocolType]::SystemDefault) {
         Try { [Net.ServicePointManager]::SecurityProtocol = @([Net.SecurityProtocolType]::Tls,[Net.SecurityProtocolType]::Tls11,[Net.SecurityProtocolType]::Tls12)}
         Catch { Exit }
    }

    If ((Get-PackageProvider).Name -notcontains "NuGet") {
        Try { Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction Stop }
        Catch { Exit }
    }
    $ArrPSRepos = Get-PSRepository
    If ($ArrPSRepos.Name -notcontains "PSGallery") {
        Try { Register-PSRepository -Default -InstallationPolicy Trusted -ErrorAction Stop }
        Catch { Exit }
    } ElseIf ($ArrPSRepos | ?{$_.Name -eq "PSGallery" -and $_.InstallationPolicy -ne "Trusted"}) {
        Try { Set-PSRepository PSGallery -InstallationPolicy Trusted -ErrorAction Stop }
        Catch { Exit }
    }
    If ((Get-Module -ListAvailable).Name -notcontains "PSReadLine") {
        Try { Install-Module PSReadLine -Force -ErrorAction Stop }
        Catch { Exit }
    }

}

# Check if the BurntToast module is installed, and install the latest version if not
$module = Get-Module -ListAvailable -Name BurntToast
if (-not $module) {
    Write-Host "Installing the latest version of the BurntToast module..."
    Install-Module -Name BurntToast -Force -Scope CurrentUser
}
else {
    # Update the BurntToast module to the latest version
    Write-Host "Updating the BurntToast module to the latest version..."
    Install-Module -Name BurntToast -Force -Scope CurrentUser
}

# Import the BurntToast module
Import-Module BurntToast



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
    @{ Name = 'InTune'; IPAddress = '172.30.125.202' }
)

function Show-ToastNotification($printerName, $isSuccess) {
    $logoPath = Join-Path $tempDirectoryPath "FDSLogo.png"
    $toastParams = @{
        AppLogo = $logoPath
    }

    if ($PSBoundParameters.ContainsKey('isSuccess')) {
        if ($isSuccess) {
            $toastParams.Text = "$printerName has been installed."
            Write-Host "Displaying success notification for $printerName"
        }
        else {
            $toastParams.Text = "Printer for '$printerName' failed to install. Please call support at 910-483-5395."
            Write-Host "Displaying failure notification for $printerName"
        }
    }

    $null = New-BurntToastNotification @toastParams
}

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

# Define the directory and temp paths
$directoryPath = "C:\ProgramData\FDS"
$tempDirectoryPath = Join-Path $directoryPath "PrinterInstall"
$tempMsiPath = Join-Path $tempDirectoryPath "6900.msi"
$tempModelDatPath = Join-Path $tempDirectoryPath "model023.dat"
$tempImagePath = Join-Path $tempDirectoryPath "FDSLogo.png"
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
Invoke-WebRequest -Uri $msiUrl -OutFile $tempMsiPath
Invoke-WebRequest -Uri $modelDatUrl -OutFile $tempModelDatPath
Invoke-WebRequest -Uri $imageURL -OutFile $tempImagePath

# Define the printer driver name
$driverName = "Brother MFC-L6900DW series"

# Check if the printer driver is installed
try {
    $printerDriver = Get-PrinterDriver -Name $driverName -ErrorAction Stop
    if ($printerDriver) {
        Write-Host "The printer driver '$driverName' is already installed."
    }
}
catch {
    Write-Warning "The printer driver '$driverName' is not installed. Installing the driver and printer..."

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
            Write-Host "Executing command: msiexec $arguments"
            $process = Start-Process -FilePath "msiexec.exe" -ArgumentList $arguments -Wait -NoNewWindow -PassThru
            $process.WaitForExit()

            if ($process.ExitCode -ne 0) {
                $errorMessage = "Failed to install printer $($config.Name). Exit code: $($process.ExitCode)"
                Write-Warning $errorMessage
                Add-Content -Path $logFilePath -Value $errorMessage
                Add-Content -Path $logFilePath -Value $process.StandardOutput
                Add-Content -Path $logFilePath -Value $process.StandardError

   		# Show the failure notification
   		Show-ToastNotification -printerName $config.Name -isSuccess $false

            } else {
                $message = "The $($config.Name) printer was installed."
                Write-Host $message

                # Update the PrintersInstalled.txt and PrinterUninstall.txt files
                UpdatePrinterLogFiles $config.Name
		
		# Show the success notification
    		Show-ToastNotification -printerName $config.Name -isSuccess $true
            }

            $matched = $true
            break
        }
    }

    # If no matching IP address was found, log the error
    if (-not $matched) {
        $errorMessage = "No matching IP address found for printer installation."
        Write-Warning $errorMessage

	# Show the failure notification
    	Show-ToastNotification -printerName $config.Name -isSuccess $false

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
                Add-PrinterPort -Name $portName -PrinterHostAddress "$($config.IPAddress)"
            }

            # Add the printer
            Add-Printer -Name $config.Name -DriverName "Brother MFC-L6900DW series" -PortName $portName

            $message = "The $($config.Name) printer was installed."
            Write-Host $message

            # Update the PrintersInstalled.txt and PrinterUninstall.txt files
            UpdatePrinterLogFiles $config.Name

	    # Show the success notification
	    Show-ToastNotification -printerName $config.Name -isSuccess $true


            $matched = $true
            break
        }
    }

    # Remove any existing scheduled task with the same name
    $existingTask = Get-ScheduledTask -TaskName "DeleteTempFiles" -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "Removing existing scheduled task 'DeleteTempFiles'..."
        Unregister-ScheduledTask -TaskName "DeleteTempFiles" -Confirm:$false
}


    # If no matching IP address was found, log the error
    if (-not $matched) {
        $errorMessage = "No matching IP address found for printer installation."
        Write-Warning $errorMessage

   	# Show the failure notification
        Show-ToastNotification -printerName $config.Name -isSuccess $false


        Add-Content -Path $logFilePath -Value $errorMessage
    }
}

# Create a scheduled task to delete the specified files after 5 minutes
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-Command `"Remove-Item -Path '$tempMsiPath', '$tempModelDatPath', '$printerInstalledLogPath', '$tempImagePath'  -Force`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5)
$principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$task = Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "DeleteTempFiles" -Description "Delete temporary files after 5 minutes"

Write-Host "Scheduled task 'DeleteTempFiles' created. It will run in 5 minutes."
