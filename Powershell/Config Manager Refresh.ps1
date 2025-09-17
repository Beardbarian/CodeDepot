# Script to run GPUpdate and install SCCM applications
# Requires administrative privileges

# Check if running as administrator and self-elevate if needed
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator

if (-not $principal.IsInRole($adminRole)) {
    try {
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
        exit
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "This tool requires administrative privileges to function properly. Please run as administrator.",
            "Administrator Rights Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        exit
    }
}

function Write-Status {
    param(
        [string]$Message,
        [string]$Status = "Info"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    switch ($Status) {
        "Info"    { Write-Host "[$timestamp] $Message" -ForegroundColor White }
        "Success" { Write-Host "[$timestamp] $Message" -ForegroundColor Green }
        "Warning" { Write-Host "[$timestamp] $Message" -ForegroundColor Yellow }
        "Error"   { Write-Host "[$timestamp] $Message" -ForegroundColor Red }
    }
}

# Step 1: Run GPUpdate /force with no restart
Write-Status "Starting Group Policy Update..." "Info"
try {
    & gpupdate.exe /force 2>&1 | Out-Null
    $exitCode = $LASTEXITCODE
    
    if ($exitCode -eq 0) {
        Write-Status "Group Policy Update completed successfully." "Success"
    } else {
        Write-Status "Group Policy Update completed with exit code: $exitCode" "Warning"
    }
} catch {
    Write-Status "Error running GPUpdate: $_" "Error"
}

# Step 2: Function to trigger Configuration Manager actions
function Invoke-CMActions {
    Write-Status "Starting Configuration Manager action cycles..." "Info"
    
    $actions = @{
        "{00000000-0000-0000-0000-000000000121}" = "Application Deployment Evaluation Cycle"
        "{00000000-0000-0000-0000-000000000003}" = "Discovery Data Collection Cycle"
        "{00000000-0000-0000-0000-000000000001}" = "Hardware Inventory Cycle"
        "{00000000-0000-0000-0000-000000000021}" = "Machine Policy Retrieval & Evaluation Cycle"
        "{00000000-0000-0000-0000-000000000031}" = "Software Metering Usage Report Cycle"
        "{00000000-0000-0000-0000-000000000108}" = "Software Updates Deployment Evaluation Cycle"
        "{00000000-0000-0000-0000-000000000113}" = "Software Updates Scan Cycle"
        "{00000000-0000-0000-0000-000000000022}" = "User Policy Retrieval & Evaluation Cycle"
        "{00000000-0000-0000-0000-000000000032}" = "Windows Installer Source List Update Cycle"
    }
    
    try {
        $CCMNamespace = "root\ccm"
        foreach ($action in $actions.GetEnumerator()) {
            Write-Status "Triggering $($action.Value)..." "Info"
            try {
                $schedule = [wmiclass]"$CCMNamespace`:SMS_Client"
                $schedule.TriggerSchedule($action.Key) | Out-Null
                Write-Status "$($action.Value) triggered successfully" "Success"
                # Small delay between actions to prevent overwhelming the system
                Start-Sleep -Seconds 2
            }
            catch {
                Write-Status "Failed to trigger $($action.Value): $_" "Error"
            }
        }
        
        # Give some time for actions to process
        Write-Status "Waiting 30 seconds for actions to process..." "Info"
        Start-Sleep -Seconds 30
        Write-Status "Configuration Manager actions completed" "Success"
    }
    catch {
        Write-Status "Error accessing Configuration Manager: $_" "Error"
    }
}


# Step 3: Run Configuration Manager Actions
Invoke-CMActions

Write-Status "`nScript completed. Group Policy and Configuration Manager actions have been run." "Info"

# Keep window open if running in PowerShell ISE or directly from PowerShell
if ($Host.Name -eq "ConsoleHost" -or $Host.Name -eq "Windows PowerShell ISE Host") {
    Write-Host "`nPress Enter to exit..."
    Read-Host
}
