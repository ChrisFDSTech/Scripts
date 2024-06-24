$PackageName = "DesktopIcon_YardiOne"
$Version = "1"

# Define the GitHub URLs for the files
$Icon = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/YardiOne/YardiOne.ico"

# Define the directory
$directoryName = "FDS\Icons"
$directoryPath = Join-Path -Path $env:ProgramData -ChildPath $directoryName

# Define the Shortcut Info
$WebAddress = "https://www.yardiasp13.com/80032dth"
$ShortcutName = "YardiOne.lnk"
$ShortcutPath = Join-Path -Path ([Environment]::GetFolderPath("CommonDesktopDirectory")) -ChildPath $ShortcutName

# Start logging
$logPath = "C:\PS\YardiOneShortcut_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
Start-Transcript -Path $logPath -Append

# Ensure the temp FDS directory exists
if (-not (Test-Path $directoryPath)) {
    try {
        Write-Host "Creating directory: $directoryPath"
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }
    catch {
        Write-Warning "Failed to create directory: $directoryPath"
        Write-Warning $_.Exception.Message
    }
}

# Download the ico file from GitHub to the FDS Directory
Write-Host "Downloading icon file from: $Icon"
try {
    Invoke-WebRequest -Uri $Icon -OutFile (Join-Path -Path $directoryPath -ChildPath "YardiOne.ico")
}
catch {
    Write-Warning "Failed to download icon file from: $Icon"
    Write-Warning $_.Exception.Message
}

# Create the desktop shortcut
$iconPath = Join-Path -Path $directoryPath -ChildPath "YardiOne.ico"
try {
    $shortcut = New-Shortcut -Path $shortcutPath -TargetPath $WebAddress -IconLocation $iconPath -Description "YardiOne"
    Write-Host "Desktop shortcut created: $($shortcut.FullName)"
}
catch {
    Write-Warning "Failed to create desktop shortcut: $ShortcutPath"
    Write-Warning $_.Exception.Message
}

# Stop logging
Stop-Transcript



<#
$PackageName = "DesktopIcon_YardiOne"
$Version = "1"

# Define the GitHub URLs for the files
$Icon = "https://raw.githubusercontent.com/ChrisFDSTech/Scripts/main/YardiOne/YardiOne.ico"

# Define the directory
$directoryName = "FDS\Icons"
$directoryPath = Join-Path -Path $env:ProgramData -ChildPath $directoryName

# Define the Shortcut Info
$WebAddress = "https://www.yardiasp13.com/80032dth"
$ShortcutName = "YardiOne.lnk"
$ShortcutPath = Join-Path -Path ([Environment]::GetFolderPath("Desktop")) -ChildPath $ShortcutName

# Ensure the temp FDS directory exists
if (-not (Test-Path $directoryPath)) {
    try {
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }
    catch {
        Write-Warning "Failed to create directory: $directoryPath"
        Write-Warning $_.Exception.Message
    }
}

# Download the ico file from GitHub to the FDS Directory
Invoke-WebRequest -Uri $Icon -OutFile (Join-Path -Path $directoryPath -ChildPath "YardiOne.ico")

# Create the desktop shortcut
$iconPath = Join-Path -Path $directoryPath -ChildPath "YardiOne.ico"
$shortcut = New-Shortcut -Path $shortcutPath -TargetPath $WebAddress -IconLocation $iconPath -Description "YardiOne"
#>
<#
$Shell = New-Object -ComObject ("WScript.Shell")
$ShortCut = $Shell.CreateShortcut($ShortcutPath)
$ShortCut.TargetPath = $WebAddress
$ShortCut.IconLocation = (Join-Path -Path $directoryPath -ChildPath "YardiOne.ico")
$ShortCut.Description = "YardiOne"
$ShortCut.Save()
#>




