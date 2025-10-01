#Requires -RunAsAdministrator
<#
    Windows 11 Upgrade Compatibility Check and Upgrade Tool (RMM Version)
    This script performs hardware compatibility checks for Windows 11
    and initiates the upgrade process if requirements are met.
    Designed to run at SYSTEM level via RMM.
    Creates detailed logs and uses a structured folder system.
    RMM Usage: This script is designed to run silently at the SYSTEM level.
    Exit Codes:
        0 = Success or already Windows 11
        1 = Hardware not compatible
        2 = Error during process
        3 = Already in upgrade process
#>

# Check if running as SYSTEM
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$isSystem = $currentUser.User.Value -eq "S-1-5-18"
if (-not $isSystem) {
    Write-Warning "This script is designed to run at the SYSTEM level via RMM."
}

# Initialize logging and folder structure
$baseFolder = "C:\Windows11Upgrade"
$logFolder = Join-Path $baseFolder "Logs"
$downloadFolder = Join-Path $baseFolder "Downloads"
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$logFile = Join-Path $logFolder "Windows11Upgrade_$timestamp.log"

# Create required folders
$folders = @($baseFolder, $logFolder, $downloadFolder)
foreach ($folder in $folders) {
    if (-not (Test-Path $folder)) {
        New-Item -ItemType Directory -Path $folder -Force | Out-Null
    }
}

# Logging function
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('Info','Warning','Error')]
        [string]$Level = 'Info'
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $logFile -Value $logMessage
    
    # Write to console with color coding
    switch ($Level) {
        'Info'    { Write-Host $logMessage }
        'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
        'Error'   { Write-Host $logMessage -ForegroundColor Red }
    }
}

# Hardware check function
function Test-Windows11Compatibility {
    Write-Log "Starting Windows 11 compatibility check..."
    $compatible = $true
    $results = @()

    # Check current Windows version
    $osInfo = Get-WmiObject Win32_OperatingSystem
    $osVersion = [System.Environment]::OSVersion.Version
    $results += @{
        Check = "Windows Version"
        Required = "Windows 10 22H2 or newer"
        Current = "$($osInfo.Caption) ($($osVersion.Build))"
        Pass = $osVersion.Build -ge 19045  # Windows 10 22H2 build number
    }
    
    # Check CPU
    $cpu = Get-WmiObject Win32_Processor
    $results += @{
        Check = "CPU Cores"
        Required = "2 or more cores"
        Current = $cpu.NumberOfCores
        Pass = $cpu.NumberOfCores -ge 2
    }
    
    $results += @{
        Check = "CPU Speed"
        Required = "1 GHz or faster"
        Current = "$($cpu.MaxClockSpeed) MHz"
        Pass = $cpu.MaxClockSpeed -ge 1000
    }

    # Check RAM
    $ram = Get-WmiObject Win32_ComputerSystem
    $ramGB = [math]::Round($ram.TotalPhysicalMemory / 1GB, 2)
    $results += @{
        Check = "RAM"
        Required = "4 GB or more"
        Current = "$ramGB GB"
        Pass = $ramGB -ge 4
    }

    # Check Storage
    $disk = Get-WmiObject Win32_LogicalDisk -Filter "DeviceID='C:'"
    $freeSpaceGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $results += @{
        Check = "Free Disk Space"
        Required = "64 GB or more"
        Current = "$freeSpaceGB GB"
        Pass = $freeSpaceGB -ge 64
    }

    # Check TPM
    try {
        $tpm = Get-WmiObject -Namespace "root\CIMV2\Security\MicrosoftTpm" -Class Win32_Tpm
        $tpmVersion = $tpm.SpecVersion
        $results += @{
            Check = "TPM Version"
            Required = "2.0"
            Current = $tpmVersion
            Pass = $tpmVersion -like "*2.0*"
        }
    }
    catch {
        $results += @{
            Check = "TPM Version"
            Required = "2.0"
            Current = "Not Found"
            Pass = $false
        }
    }

    # Check Secure Boot
    try {
        $secureBootStatus = Confirm-SecureBootUEFI
        $results += @{
            Check = "Secure Boot"
            Required = "Enabled"
            Current = if ($secureBootStatus) { "Enabled" } else { "Disabled" }
            Pass = $secureBootStatus
        }
    }
    catch {
        $results += @{
            Check = "Secure Boot"
            Required = "Enabled"
            Current = "Not Available"
            Pass = $false
        }
    }

    # Log results
    Write-Log "`nWindows 11 Compatibility Check Results:"
    foreach ($result in $results) {
        $status = if ($result.Pass) { "PASS" } else { "FAIL" }
        $message = "{0}: {1} (Required: {2}, Current: {3})" -f $result.Check, $status, $result.Required, $result.Current
        $level = if ($result.Pass) { "Info" } else { "Warning" }
        Write-Log $message $level
        
        if (-not $result.Pass) {
            $compatible = $false
        }
    }

    return $compatible
}

# Download Windows 11 Update Assistant
function Get-Windows11Upgrade {
    $downloadUrl = "https://go.microsoft.com/fwlink/?linkid=2171764"
    $upgradeAssistant = Join-Path $downloadFolder "Windows11UpgradeAssistant.exe"
    
    Write-Log "Downloading Windows 11 Upgrade Assistant..."
    try {
        Invoke-WebRequest -Uri $downloadUrl -OutFile $upgradeAssistant
        Write-Log "Download completed successfully"
        return $upgradeAssistant
    }
    catch {
        Write-Log "Failed to download Windows 11 Upgrade Assistant: $_" "Error"
        return $null
    }
}

# Main execution
Write-Log "Starting Windows 11 Upgrade Assessment"
Write-Log "Log file: $logFile"

# Check for existing upgrade process
$setupPath = "C:\$WINDOWS.~BT\Sources\SetupHost.exe"
if (Test-Path $setupPath) {
    Write-Log "Windows upgrade process already in progress. Exiting." "Warning"
    exit 3
}

# Check if already running Windows 11
$osInfo = Get-WmiObject Win32_OperatingSystem
if ($osInfo.Caption -like "*Windows 11*") {
    Write-Log "System is already running Windows 11. No upgrade needed." "Info"
    exit 0
}

# Check hardware compatibility
$isCompatible = Test-Windows11Compatibility

if ($isCompatible) {
    Write-Log "System meets Windows 11 requirements. Proceeding with upgrade preparation..." "Info"
    
    # Download Windows 11 Upgrade Assistant
    $upgradeAssistant = Get-Windows11Upgrade
    
    if ($upgradeAssistant -and (Test-Path $upgradeAssistant)) {
        Write-Log "Starting Windows 11 Upgrade Assistant in silent mode..."
        
        try {
            # Start the upgrade assistant silently
            $startInfo = New-Object System.Diagnostics.ProcessStartInfo
            $startInfo.FileName = $upgradeAssistant
            $startInfo.Arguments = "/quietinstall /skipeula /auto upgrade /copylogs $logFolder"
            $startInfo.UseShellExecute = $false
            $startInfo.RedirectStandardOutput = $true
            $startInfo.RedirectStandardError = $true
            $startInfo.WindowStyle = [System.Diagnostics.ProcessWindowStyle]::Hidden

            $process = New-Object System.Diagnostics.Process
            $process.StartInfo = $startInfo
            $process.Start() | Out-Null
            
            Write-Log "Windows 11 upgrade process initiated successfully in silent mode."
            exit 0
        }
        catch {
            Write-Log "Failed to start upgrade process: $_" "Error"
            exit 2
        }
    }
    else {
        Write-Log "Failed to prepare for upgrade. Please check the logs for details." "Error"
        exit 2
    }
}
else {
    Write-Log "System does not meet Windows 11 requirements. Check logs for details." "Error"
    exit 1
}
