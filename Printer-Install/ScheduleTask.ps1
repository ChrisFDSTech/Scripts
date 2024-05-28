# Remove any existing scheduled task with the same name
    $existingTask = Get-ScheduledTask -TaskName "DeleteTempFiles" -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "Removing existing scheduled task 'DeleteTempFiles'..."
        Unregister-ScheduledTask -TaskName "DeleteFiles" -Confirm:$false
}

# Create a scheduled task to delete the specified files after 5 minutes
$action = New-ScheduledTaskAction -Execute 'PowerShell.exe' -Argument "-Command `"Remove-Item -Path '$tempMsiPath', '$tempModelDatPath', '$printerInstalledLogPath', '$tempImagePath', '$schedTaskPath'  -Force`""
$trigger = New-ScheduledTaskTrigger -Once -At (Get-Date).AddMinutes(5)
$principal = New-ScheduledTaskPrincipal -UserID "NT AUTHORITY\SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$null = Register-ScheduledTask -Action $action -Trigger $trigger -Principal $principal -TaskName "DeleteFiles" -Description "Delete temporary files after 5 minutes"

Write-Host "Scheduled task 'DeleteTempFiles' created. It will run in 5 minutes."
