#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Windows 11 Update Script for Domain Computers
.DESCRIPTION
    Checks for and installs Windows updates on Windows 11 systems
    Designed to run as domain admin on remote or local computers
.NOTES
    Author: PowerShell Automation
    Requires: Administrator privileges
#>

# Configuration
$LogPath = "C:\Windows\Temp\WindowsUpdate_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$RebootRequired = $false

# Function to write logs
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogPath -Value $LogMessage
    Write-Host $LogMessage
}

# Function to check if reboot is pending
function Test-PendingReboot {
    $rebootPending = $false
    
    # Check Component Based Servicing
    $cbs = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction SilentlyContinue
    if ($cbs) { $rebootPending = $true }
    
    # Check Windows Update
    $wu = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
    if ($wu) { $rebootPending = $true }
    
    # Check PendingFileRenameOperations
    $pfro = Get-ItemProperty "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
    if ($pfro) { $rebootPending = $true }
    
    return $rebootPending
}

# Start logging
Write-Log "========================================" "INFO"
Write-Log "Windows Update Script Started" "INFO"
Write-Log "Computer: $env:COMPUTERNAME" "INFO"
Write-Log "User: $env:USERNAME" "INFO"
Write-Log "========================================" "INFO"

# Check Windows version
$OSInfo = Get-CimInstance -ClassName Win32_OperatingSystem
Write-Log "OS: $($OSInfo.Caption) - Build: $($OSInfo.BuildNumber)" "INFO"

# Method 1: Try PSWindowsUpdate module (if available)
$PSWindowsUpdateAvailable = Get-Module -ListAvailable -Name PSWindowsUpdate

if ($PSWindowsUpdateAvailable) {
    Write-Log "PSWindowsUpdate module found, using it for updates" "INFO"
    
    try {
        Import-Module PSWindowsUpdate -ErrorAction Stop
        Write-Log "Checking for available updates..." "INFO"
        
        # Get available updates
        $Updates = Get-WindowsUpdate -MicrosoftUpdate -Verbose
        
        if ($Updates) {
            Write-Log "Found $($Updates.Count) update(s)" "INFO"
            
            # Install updates
            Write-Log "Installing updates..." "INFO"
            Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot:$false -Verbose
            
            Write-Log "Updates installed successfully" "INFO"
        } else {
            Write-Log "No updates available" "INFO"
        }
    } catch {
        Write-Log "Error using PSWindowsUpdate: $($_.Exception.Message)" "ERROR"
        Write-Log "Falling back to COM object method" "WARN"
        $PSWindowsUpdateAvailable = $false
    }
}

# Method 2: Use Windows Update COM objects (fallback)
if (-not $PSWindowsUpdateAvailable) {
    Write-Log "Using Windows Update COM objects" "INFO"
    
    try {
        # Create Windows Update session
        $UpdateSession = New-Object -ComObject Microsoft.Update.Session
        $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
        
        Write-Log "Searching for updates..." "INFO"
        $SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
        
        if ($SearchResult.Updates.Count -eq 0) {
            Write-Log "No updates available" "INFO"
        } else {
            Write-Log "Found $($SearchResult.Updates.Count) update(s)" "INFO"
            
            # Create update collection
            $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
            
            foreach ($Update in $SearchResult.Updates) {
                Write-Log "  - $($Update.Title)" "INFO"
                $UpdatesToInstall.Add($Update) | Out-Null
            }
            
            # Download updates
            Write-Log "Downloading updates..." "INFO"
            $Downloader = $UpdateSession.CreateUpdateDownloader()
            $Downloader.Updates = $UpdatesToInstall
            $DownloadResult = $Downloader.Download()
            
            if ($DownloadResult.ResultCode -eq 2) {
                Write-Log "Updates downloaded successfully" "INFO"
                
                # Install updates
                Write-Log "Installing updates..." "INFO"
                $Installer = $UpdateSession.CreateUpdateInstaller()
                $Installer.Updates = $UpdatesToInstall
                $InstallResult = $Installer.Install()
                
                if ($InstallResult.ResultCode -eq 2) {
                    Write-Log "Updates installed successfully" "INFO"
                } else {
                    Write-Log "Installation completed with result code: $($InstallResult.ResultCode)" "WARN"
                }
                
                if ($InstallResult.RebootRequired) {
                    $RebootRequired = $true
                    Write-Log "Reboot is required to complete installation" "WARN"
                }
            } else {
                Write-Log "Download failed with result code: $($DownloadResult.ResultCode)" "ERROR"
            }
        }
    } catch {
        Write-Log "Error during update process: $($_.Exception.Message)" "ERROR"
    }
}

# Check for pending reboot
if (Test-PendingReboot) {
    Write-Log "A reboot is pending on this system" "WARN"
    $RebootRequired = $true
}

# Final summary
Write-Log "========================================" "INFO"
Write-Log "Windows Update Script Completed" "INFO"
Write-Log "Reboot Required: $RebootRequired" "INFO"
Write-Log "Log file: $LogPath" "INFO"
Write-Log "========================================" "INFO"

Optional: Uncomment to auto-reboot after updates
if ($RebootRequired) {
    Write-Log "Initiating reboot in 60 seconds..." "WARN"
    shutdown /r /t 60 /c "Windows updates installed. System will reboot in 60 seconds."
}

# Return exit code
if ($RebootRequired) {
    exit 3010  # Reboot required
} else {
    exit 0  # Success
}
