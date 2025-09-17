# Script to remove all network printers and clean up unused printer ports
# Requires administrative privileges

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [switch]$Force,
    
    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

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

# Check if running with administrator privileges and self-elevate if needed
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator

if (-not $principal.IsInRole($adminRole)) {
    Write-Status "Requesting administrative privileges..." "Info"
    try {
        $argList = @()
        if ($Force) { $argList += "-Force" }
        if ($WhatIf) { $argList += "-WhatIf" }
        
        $scriptPath = $MyInvocation.MyCommand.Path
        $argString = $argList -join ' '
        
        Write-Status "Relaunching script with elevated privileges..." "Info"
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" $argString" -Verb RunAs -Wait
        exit
    }
    catch {
        Write-Status "Failed to relaunch with administrative privileges: $_" "Error"
        Write-Status "Please right-click the script and select 'Run as Administrator'" "Error"
        exit 1
    }
}

try {
    Write-Status "Starting printer cleanup process..." "Info"
    
    # Get all printers
    Write-Status "Getting list of all printers..." "Info"
    $allPrinters = Get-CimInstance -ClassName Win32_Printer
    $networkPrinters = $allPrinters | Where-Object { $_.Network }
    $localPrinters = $allPrinters | Where-Object { -not $_.Network }
    
    Write-Status "Found $($networkPrinters.Count) network printer(s)" "Info"
    
    if (-not $Force -and $networkPrinters.Count -gt 0) {
        Write-Status "Network printers to be removed:" "Warning"
        $networkPrinters | ForEach-Object {
            Write-Status "  - $($_.Name) ($($_.PortName))" "Warning"
        }
        
        $title = "Remove Network Printers"
        $message = "Do you want to remove these network printers?"
        $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Remove all network printers"
        $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel the operation"
        $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
        
        $result = $host.UI.PromptForChoice($title, $message, $options, 1)
        
        if ($result -eq 1) {
            Write-Status "Operation cancelled by user" "Info"
            exit 0
        }
    }
    
    # Remove network printers
    foreach ($printer in $networkPrinters) {
        if ($WhatIf) {
            Write-Status "What if: Would remove network printer: $($printer.Name)" "Warning"
            continue
        }
        
        try {
            Write-Status "Removing printer: $($printer.Name)..." "Info"
            Remove-CimInstance -InputObject $printer -ErrorAction Stop
            Write-Status "Successfully removed printer: $($printer.Name)" "Success"
        }
        catch {
            Write-Status "Failed to remove printer $($printer.Name): $_" "Error"
        }
    }
    
    # Get all printer ports
    Write-Status "`nGetting list of printer ports..." "Info"
    $printerPorts = Get-CimInstance -ClassName Win32_TCPIPPrinterPort
    
    # Find unused ports (not connected to any local printer)
    $usedPorts = $localPrinters | Select-Object -ExpandProperty PortName
    $unusedPorts = $printerPorts | Where-Object { $usedPorts -notcontains $_.Name }
    
    if ($unusedPorts.Count -gt 0) {
        Write-Status "Found $($unusedPorts.Count) unused printer port(s)" "Info"
        
        if (-not $Force) {
            Write-Status "Unused ports to be removed:" "Warning"
            $unusedPorts | ForEach-Object {
                Write-Status "  - $($_.Name) ($($_.Description))" "Warning"
            }
            
            $title = "Remove Unused Ports"
            $message = "Do you want to remove these unused printer ports?"
            $yes = New-Object System.Management.Automation.Host.ChoiceDescription "&Yes", "Remove all unused ports"
            $no = New-Object System.Management.Automation.Host.ChoiceDescription "&No", "Cancel port removal"
            $options = [System.Management.Automation.Host.ChoiceDescription[]]($yes, $no)
            
            $result = $host.UI.PromptForChoice($title, $message, $options, 1)
            
            if ($result -eq 1) {
                Write-Status "Port removal cancelled by user" "Info"
                exit 0
            }
        }
        
        # Remove unused ports
        foreach ($port in $unusedPorts) {
            if ($WhatIf) {
                Write-Status "What if: Would remove unused port: $($port.Name)" "Warning"
                continue
            }
            
            try {
                Write-Status "Removing port: $($port.Name)..." "Info"
                Remove-CimInstance -InputObject $port -ErrorAction Stop
                Write-Status "Successfully removed port: $($port.Name)" "Success"
            }
            catch {
                Write-Status "Failed to remove port $($port.Name): $_" "Error"
            }
        }
    }
    else {
        Write-Status "No unused printer ports found" "Success"
    }
}
catch {
    Write-Status "Error during cleanup process: $_" "Error"
    exit 1
}

Write-Status "`nPrinter cleanup completed" "Success"

# Keep window open if running in PowerShell ISE or directly from PowerShell
if ($Host.Name -eq "ConsoleHost" -or $Host.Name -eq "Windows PowerShell ISE Host") {
    Write-Host "`nPress Enter to exit..."
    Read-Host
}
