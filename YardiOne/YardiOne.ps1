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

# Start logging if any part of the process fails
$logPath = "C:\FDSLogs"
if (-not (Test-Path $logPath)) {
    try {
        New-Item -ItemType Directory -Path $logPath -Force | Out-Null
    }
    catch {
        Write-Host "Failed to create directory: $logPath"
        Start-Transcript -Path (Join-Path -Path $env:ProgramData -ChildPath "$($PackageName)_$(Get-Date -Format 'MM-dd-yyyy_HH:mm:ss').txt") -Append
    }
}
else {
    Start-Transcript -Path (Join-Path -Path $logPath -ChildPath "$($PackageName)_$(Get-Date -Format 'MM-dd-yyyy_HH:mm:ss').txt") -Append
}

# Create the directory
if (-not (Test-Path $directoryPath)) {
    try {
        Write-Host "Creating directory: $directoryPath"
        New-Item -ItemType Directory -Path $directoryPath -Force | Out-Null
    }
    catch {
        Write-Warning "Failed to create directory: $directoryPath"
    }
}

# Download the ico file from GitHub to the FDS Directory
Write-Host "Downloading icon file from: $Icon"
try {
    Invoke-WebRequest -Uri $Icon -OutFile (Join-Path -Path $directoryPath -ChildPath "YardiOne.ico")
}
catch {
    Write-Warning "Failed to download icon file from: $Icon"
}

# Create the desktop shortcut using WScript.Shell
$iconPath = Join-Path -Path $directoryPath -ChildPath "YardiOne.ico"
try {
    $Shell = New-Object -ComObject ("WScript.Shell")
    $ShortCut = $Shell.CreateShortcut($ShortcutPath)
    $ShortCut.TargetPath = $WebAddress
    $ShortCut.IconLocation = $iconPath
    $ShortCut.Description = "YardiOne"
    $ShortCut.Save()
}
catch {
    Write-Warning "Failed to create desktop shortcut: $ShortcutPath"
}

# Stop logging
Stop-Transcript
