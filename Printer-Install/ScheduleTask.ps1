# Define the directory and temp paths
$directoryPath = "C:\ProgramData\FDS"
$tempDirectoryPath = Join-Path $directoryPath "PrinterInstall"
$tempMsiPath = Join-Path $tempDirectoryPath "6900.msi"
$tempModelDatPath = Join-Path $tempDirectoryPath "model023.dat"
$tempImagePath = Join-Path $tempDirectoryPath "FDSLogo.png"
$schedTaskPath = Join-Path $tempDirectoryPath "ScheduleTask.ps1"
$printerInstalledLogPath = Join-Path $tempDirectoryPath "PrintersInstalled.txt"
$printerUninstallLogPath = Join-Path $tempDirectoryPath "PrinterUninstall.txt"


# Remove any existing scheduled task with the same name
    $existingTask = Get-ScheduledTask -TaskName "DeleteFiles" -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "Removing existing scheduled task 'DeleteFiles'..."
        Unregister-ScheduledTask -TaskName "DeleteFiles" -Confirm:$false
}

# Create a scheduled task to delete the specified files after 5 minutes
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-Command `"Remove-Item -Path '$tempMsiPath', '$tempModelDatPath', '$printerInstalledLogPath', '$tempImagePath', '$schedTaskPath'  -Force`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5)
$principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$null = Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "DeleteFiles" -Description "Delete temporary files after 5 minutes"

Write-Host "Scheduled task 'DeleteFiles' created. It will run in 5 minutes."
