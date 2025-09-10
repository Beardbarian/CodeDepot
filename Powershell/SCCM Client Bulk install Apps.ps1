# Define the list of applications to check and install.
# Ensure these names match the 'Application Name' in your MECM console.
# Example:
# $ApplicationNames = @(
#     "Citrix Workspace",
#     "Microsoft 365 Apps for enterprise",
#     "7-Zip"
# )

$ApplicationNames = @(
      "Citrix Workspace",
      "Microsoft 365 Apps for enterprise",
      "7-Zip"
)

# Start the script with a clear status message.
Write-Host "Starting software installation check for $($ApplicationNames.Count) applications." -ForegroundColor Cyan
Write-Host "This script will trigger installations via Software Center if required." -ForegroundColor Yellow
Write-Host "--------------------------------------------------------------------------`n"

# Loop through each application name in the defined list.
foreach ($appName in $ApplicationNames) {
    Write-Host "--> Checking for: '$appName'..." -ForegroundColor White

    try {
        # Use Get-CimInstance to retrieve all CCM_Application objects and then filter them
        # with Where-Object. This is a more robust method to prevent "Provider not capable" errors.
        # We also use 'Select-Object -First 1' to handle cases where multiple apps have similar names.
        $appInstance = Get-CimInstance -Namespace "ROOT\ccm\ClientSDK" -ClassName "CCM_Application" -ErrorAction Stop | Where-Object { $_.Name -like "$appName" } | Select-Object -First 1
        
        # Check if the application object was found AND it's not already installed.
        if ($null -ne $appInstance -and $appInstance.InstallState -ne "Installed") {
            Write-Host "    '$appName' is available but not installed. Triggering installation..." -ForegroundColor Yellow

            # Call the 'Install' method on the application's WMI instance.
            # This is the equivalent of a user clicking 'Install' in Software Center.
            # We are now using only the essential parameters to ensure compatibility.
            Invoke-CimMethod -Namespace "ROOT\ccm\ClientSDK" -ClassName "CCM_Application" -MethodName "Install" -Arguments @{
                id              = $appInstance.id
                Revision        = $appInstance.Revision
                IsMachineTarget = $appInstance.IsMachineTarget
            }

            Write-Host "    Installation for '$appName' has been successfully triggered." -ForegroundColor Green
            Write-Host ""
        } elseif ($null -eq $appInstance) {
            Write-Host "    '$appName' was not found in Software Center. Please verify the application name." -ForegroundColor Red
            Write-Host ""
        } else {
            # Application is found and its InstallState is "Installed".
            Write-Host "    '$appName' is already installed. No action required." -ForegroundColor Green
            Write-Host ""
        }
    } catch {
        # Catch any errors from Get-CimInstance or Invoke-CimMethod.
        Write-Host "    An error occurred while processing '$appName'. Error: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host "--------------------------------------------------------------------------"
Write-Host "Script execution complete." -ForegroundColor Cyan
