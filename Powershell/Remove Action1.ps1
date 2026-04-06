# Get the Action1 Agent software entry
$action1 = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "*Action1*" }

if ($action1) {
    Write-Host "Found Action1 Agent: $($action1.Name). Proceeding with uninstall..."
    
    # Execute uninstallation via MSIExec
    $process = Start-Process "msiexec.exe" -ArgumentList "/x $($action1.IdentifyingNumber) /qn /norestart" -Wait -PassThru
    
    if ($process.ExitCode -eq 0) {
        Write-Host "Action1 successfully removed."
        exit 0
    } else {
        Write-Error "Uninstallation failed with exit code $($process.ExitCode)."
        exit 1
    }
} else {
    Write-Host "Action1 Agent was not found on this system."
    exit 0
}
