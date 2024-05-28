 # Remove any existing scheduled task with the same name
    $existingTask = Get-ScheduledTask -TaskName "DeleteTempFiles" -ErrorAction SilentlyContinue
    if ($existingTask) {
        Write-Host "Removing existing scheduled task 'DeleteTempFiles'..."
        Unregister-ScheduledTask -TaskName "DeleteTempFiles" -Confirm:$false
}

