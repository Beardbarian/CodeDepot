# PowerShell System Management Tool v1.3

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

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ==========================================
# Form Setup and Styling
# ==========================================

# Color scheme
$darkBackground = [System.Drawing.ColorTranslator]::FromHtml("#1a1a1a")    # Dark background
$panelBackground = [System.Drawing.ColorTranslator]::FromHtml("#2d2d2d")   # Slightly lighter panels
$controlBackground = [System.Drawing.ColorTranslator]::FromHtml("#363636") # Input/control background
$accentColor = [System.Drawing.ColorTranslator]::FromHtml("#007acc")       # Blue accent
$textColor = [System.Drawing.ColorTranslator]::FromHtml("#ffffff")         # White text
$subTextColor = [System.Drawing.ColorTranslator]::FromHtml("#cccccc")      # Light gray text

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "PowerShell System Management Tool v1.3"
$form.Size = New-Object System.Drawing.Size(840, 740)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = $darkBackground
$form.ForeColor = $textColor


# ==========================================
# Control Creation
# ==========================================

# Target Computer Input Group
$targetGroup = New-Object System.Windows.Forms.GroupBox
$targetGroup.Text = "Target Computer"
$targetGroup.Location = New-Object System.Drawing.Point($margin, 15)
$targetGroup.Size = New-Object System.Drawing.Size(822, 55)
$targetGroup.BackColor = $panelBackground
$targetGroup.ForeColor = $textColor

$computerLabel = New-Object System.Windows.Forms.Label
$computerLabel.Text = "Computer Name:"
$computerLabel.Location = New-Object System.Drawing.Point(10, 25)
$computerLabel.AutoSize = $true

$computerInput = New-Object System.Windows.Forms.TextBox
$computerInput.Location = New-Object System.Drawing.Point(110, 22)
$computerInput.Size = New-Object System.Drawing.Size(250, 20)
$computerInput.Text = $env:COMPUTERNAME
$computerInput.BackColor = $controlBackground
$computerInput.ForeColor = $textColor

# Add TextChanged event handler
$computerInput.Add_TextChanged({
    if ([string]::IsNullOrWhiteSpace($computerInput.Text)) {
        $outputBox.Text = "Please enter computer name and click a button."
    }
})

# Output Area
$outputGroup = New-Object System.Windows.Forms.GroupBox
$outputGroup.Text = "Output"
$outputGroup.Location = New-Object System.Drawing.Point($margin, 230)
$outputGroup.Size = New-Object System.Drawing.Size(822, 440)
$outputGroup.BackColor = $panelBackground
$outputGroup.ForeColor = $textColor

# Action Buttons Group
$buttonGroup = New-Object System.Windows.Forms.GroupBox
$buttonGroup.Text = "Actions"
$buttonGroup.Location = New-Object System.Drawing.Point($margin, 80)
$buttonGroup.Size = New-Object System.Drawing.Size(822, 140)  # Increased height for third row
$buttonGroup.BackColor = $panelBackground
$buttonGroup.ForeColor = $textColor

# Button row positions
$row1Y = 20
$row2Y = $row1Y + $buttonHeight + 12  # 12px spacing between rows
$row3Y = $row2Y + $buttonHeight + 12

# Button styling function
function Style-Button {
    param($button)
    $button.BackColor = $accentColor
    $button.ForeColor = $textColor
    $button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $button.FlatAppearance.BorderSize = 0
    $button.Cursor = [System.Windows.Forms.Cursors]::Hand
    
    # Hover effects
    $button.Add_MouseEnter({
        $this.BackColor = [System.Drawing.ColorTranslator]::FromHtml("#1a8fff")
    })
    $button.Add_MouseLeave({
        $this.BackColor = $accentColor
    })
}

# Calculate button positions
$margin = 20
$buttonWidth = 120
$buttonHeight = 32
$buttonSpacing = 12
$buttonsPerRow = 6

# Fixed starting position for left alignment
$buttonStartX = 20

# Row 1 - System Information and Diagnostics
$btnPing = New-Object System.Windows.Forms.Button
$btnPing.Text = "Ping"
$btnPing.Location = New-Object System.Drawing.Point($buttonStartX, 20)
$btnPing.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnPing

$btnSysInfo = New-Object System.Windows.Forms.Button
$btnSysInfo.Text = "System Info"
$btnSysInfo.Location = New-Object System.Drawing.Point(($buttonStartX + $buttonWidth + $buttonSpacing), 20)
$btnSysInfo.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnSysInfo

$btnUptime = New-Object System.Windows.Forms.Button
$btnUptime.Text = "Uptime"
$btnUptime.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 2), 20)
$btnUptime.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnUptime

$btnUsers = New-Object System.Windows.Forms.Button
$btnUsers.Text = "User Sessions"
$btnUsers.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 3), 20)
$btnUsers.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnUsers

$btnPrinters = New-Object System.Windows.Forms.Button
$btnPrinters.Text = "Printers"
$btnPrinters.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 4), 20)
$btnPrinters.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnPrinters

$btnApps = New-Object System.Windows.Forms.Button
$btnApps.Text = "Applications"
$btnApps.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 5), 20)
$btnApps.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnApps

# Row 2 - Management and Control
$btnServices = New-Object System.Windows.Forms.Button
$btnServices.Text = "Services"
$btnServices.Location = New-Object System.Drawing.Point($buttonStartX, 60)
$btnServices.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnServices

$btnDiskSpace = New-Object System.Windows.Forms.Button
$btnDiskSpace.Text = "Disk Space"
$btnDiskSpace.Location = New-Object System.Drawing.Point(($buttonStartX + $buttonWidth + $buttonSpacing), 60)
$btnDiskSpace.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnDiskSpace

$btnCompMgmt = New-Object System.Windows.Forms.Button
$btnCompMgmt.Text = "Computer Mgmt"
$btnCompMgmt.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 2), 60)
$btnCompMgmt.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnCompMgmt

$btnOpenShare = New-Object System.Windows.Forms.Button
$btnOpenShare.Text = "C$ Share"
$btnOpenShare.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 3), 60)
$btnOpenShare.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnOpenShare

$btnLogOff = New-Object System.Windows.Forms.Button
$btnLogOff.Text = "Log Off Users"
$btnLogOff.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 4), 60)
$btnLogOff.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnLogOff

# Row 3 - Maintenance and Cleanup
$btnPrinterCleanup = New-Object System.Windows.Forms.Button
$btnPrinterCleanup.Text = "Clean Printers"
$btnPrinterCleanup.Location = New-Object System.Drawing.Point($buttonStartX, 100)  # Third row Y position
$btnPrinterCleanup.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$btnPrinterCleanup.Add_Click({ Start-PrinterCleanup })
Style-Button $btnPrinterCleanup

$btnRenamePC = New-Object System.Windows.Forms.Button
$btnRenamePC.Text = "Rename PC"
$btnRenamePC.Location = New-Object System.Drawing.Point(($buttonStartX + $buttonWidth + $buttonSpacing), 100)  # Next to Clean Printers
$btnRenamePC.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
$btnRenamePC.Add_Click({ Start-ComputerRename })
Style-Button $btnRenamePC

$btnRestart = New-Object System.Windows.Forms.Button
$btnRestart.Text = "Restart PC"
$btnRestart.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 5), 60)
$btnRestart.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnRestart

# Calculate the same width as button rows
$totalButtonsWidth = ($buttonWidth * 6) + ($buttonSpacing * 5)

$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point($buttonStartX, 20)
$outputBox.Size = New-Object System.Drawing.Size($totalButtonsWidth, 410)
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.BackColor = [System.Drawing.Color]::Black
$outputBox.ForeColor = [System.Drawing.Color]::White
$outputBox.ReadOnly = $true
$outputBox.MultiLine = $true
$outputBox.ScrollBars = "Vertical"

# Status Bar with dark theme
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $panelBackground
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusLabel.ForeColor = $subTextColor
$statusStrip.Items.Add($statusLabel)

# ==========================================
# Helper Functions
# ==========================================

# Function to update status
function Update-Status {
    param([string]$status)
    $statusLabel.Text = $status
    [System.Windows.Forms.Application]::DoEvents()
}

# Function to test remote connectivity
function Test-RemoteConnection {
    param([string]$computerName)
    
    if ($computerName -eq $env:COMPUTERNAME -or $computerName -eq ".") { return $true }
    
    Update-Status "Testing connection to $computerName..."
    $result = Test-Connection -ComputerName $computerName -Count 1 -Quiet
    
    if (-not $result) {
        [System.Windows.Forms.MessageBox]::Show(
            "Cannot connect to $computerName. Please verify the computer name and network connectivity.",
            "Connection Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    
    return $result
}

# Function to run commands (local or remote)
function Invoke-RemoteCommand {
    param(
        [string]$scriptBlock,
        [string]$computerName
    )
    
    try {
        if ($computerName -eq $env:COMPUTERNAME -or $computerName -eq ".") {
            return Invoke-Expression $scriptBlock
        } else {
            $scriptBlockObj = [ScriptBlock]::Create($scriptBlock)
            return Invoke-Command -ComputerName $computerName -ScriptBlock $scriptBlockObj
        }
    }
    catch {
        throw $_.Exception.Message
    }
}

# ==========================================
# Command Functions
# ==========================================

# Get System Information
function Get-SystemInformation {
    Update-Status "Getting system information..."
    
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    try {
        $script = {
            $cs = Get-CimInstance Win32_ComputerSystem
            $os = Get-CimInstance Win32_OperatingSystem
            $proc = Get-CimInstance Win32_Processor
            
            "SYSTEM INFORMATION`n==================`n"
            "Computer Name : $env:COMPUTERNAME"
            "OS Version    : $($os.Caption)"
            "OS Build      : $($os.BuildNumber)"
            "Manufacturer  : $($cs.Manufacturer)"
            "Model        : $($cs.Model)"
            "Processor    : $($proc.Name)"
            "Memory (GB)  : $([math]::Round($cs.TotalPhysicalMemory/1GB, 2))"
            "Free Mem(GB) : $([math]::Round($os.FreePhysicalMemory/1MB, 2))"
            "Last Boot    : $($os.LastBootUpTime)"
        }
        
        $result = Invoke-Command -scriptBlock $script -computerName $computerName
        $outputBox.Text = $result | Out-String
    }
    catch {
        $outputBox.Text = "Error: $_"
    }
    
    Update-Status "Ready"
}

# Get Running Services
function Get-RunningServices {
    Update-Status "Getting running services..."
    
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    try {
        $script = {
            "RUNNING SERVICES`n================`n"
            Get-Service | 
                Where-Object {$_.Status -eq 'Running'} |
                Sort-Object DisplayName |
                Format-Table -AutoSize @{N='Service';E={$_.DisplayName}}, Status, StartType |
                Out-String -Width 120
        }
        
        $result = Invoke-Command -scriptBlock $script -computerName $computerName
        $outputBox.Text = $result
    }
    catch {
        $outputBox.Text = "Error: $_"
    }
    
    Update-Status "Ready"
}

# Get Disk Space
function Get-DiskSpaceInfo {
    Update-Status "Getting disk space information..."
    
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    try {
        $script = {
            "DISK SPACE INFORMATION`n=====================`n"
            Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' |
                ForEach-Object {
                    "Drive: $($_.DeviceID)"
                    "Label: $($_.VolumeName)"
                    "Size (GB): $([math]::Round($_.Size/1GB, 2))"
                    "Free (GB): $([math]::Round($_.FreeSpace/1GB, 2))"
                    "Free (%): $([math]::Round(($_.FreeSpace/$_.Size)*100, 1))"
                    "------------------------`n"
                }
        }
        
        $result = Invoke-Command -scriptBlock $script -computerName $computerName
        $outputBox.Text = $result
    }
    catch {
        $outputBox.Text = "Error: $_"
    }
    
    Update-Status "Ready"
}

# Function to ping computer
function Test-ComputerPing {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    Update-Status "Pinging $computerName..."
    try {
        $result = Test-Connection -ComputerName $computerName -Count 4 -ErrorAction Stop
        $outputBox.Text = $result | Format-Table -AutoSize | Out-String
    }
    catch {
        $outputBox.Text = "Error pinging $computerName`: $($_.Exception.Message)"
    }
    Update-Status "Ready"
}

# Function to get uptime
function Get-ComputerUptime {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    Update-Status "Getting uptime for $computerName..."
    try {
        $script = {
            $os = Get-CimInstance Win32_OperatingSystem
            $lastBoot = $os.LastBootUpTime
            $uptime = (Get-Date) - $lastBoot
            
            "SYSTEM UPTIME`n=============`n"
            "Last Boot Time: $lastBoot"
            "Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"
        }
        
        $result = Invoke-Command -scriptBlock $script -computerName $computerName
        $outputBox.Text = $result | Out-String
    }
    catch {
        $outputBox.Text = "Error getting uptime: $_"
    }
    Update-Status "Ready"
}

# Function to get current and last users
function Get-UserSessions {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    Update-Status "Getting user sessions for $computerName..."
    try {
        $script = {
            "CURRENT SESSIONS`n===============`n"
            quser 2>&1 | ForEach-Object { $_ | Out-String }
            "`nLAST LOGON INFORMATION`n=====================`n"
            Get-CimInstance Win32_LogonSession -Filter "LogonType=2 OR LogonType=10" | 
                Select-Object -Last 5 |
                ForEach-Object {
                    $user = Get-CimAssociatedInstance -InputObject $_ -Association Win32_LoggedOnUser
                    [PSCustomObject]@{
                        Username = $user.Caption
                        StartTime = $_.StartTime
                        LogonType = $_.LogonType
                    }
                } | Format-Table -AutoSize
        }
        
        $result = Invoke-Command -scriptBlock $script -computerName $computerName
        $outputBox.Text = $result | Out-String
    }
    catch {
        $outputBox.Text = "Error getting user sessions: $_"
    }
    Update-Status "Ready"
}

# Function to get printers
function Get-PrinterInfo {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    Update-Status "Getting printer information for $computerName..."
    try {
        $script = {
            "PRINTER INFORMATION`n==================`n"
            Get-Printer | Select-Object Name, DriverName, PortName, Shared, Published, Status |
            Format-Table -AutoSize
        }
        
        $result = Invoke-Command -scriptBlock $script -computerName $computerName
        $outputBox.Text = $result | Out-String
    }
    catch {
        $outputBox.Text = "Error getting printer information: $_"
    }
    Update-Status "Ready"
}

# Function to log off users
function Invoke-LogOffUsers {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    # Get current sessions first
    try {
        $sessions = $null
        if ($computerName -eq "." -or $computerName -eq $env:COMPUTERNAME) {
            $sessions = query session 2>&1
        } else {
            $sessions = Invoke-Command -ComputerName $computerName -ScriptBlock { query session 2>&1 }
        }
        
        # Filter out the header and console session
        $activeSessions = $sessions | Select-Object -Skip 1 | Where-Object { 
            $_ -match '\s+(rdp-tcp|console)\s+' -and $_ -notmatch '^\s*SESSIONNAME'
        }
        
        if (-not $activeSessions) {
            $outputBox.Text = "No active user sessions found on $computerName"
            Update-Status "Ready"
            return
        }
        
        # Show sessions that will be logged off
        $sessionList = $activeSessions | ForEach-Object { $_.Trim() } | Out-String
        $result = [System.Windows.Forms.MessageBox]::Show(
            "The following sessions will be logged off on $computerName`:`n`n$sessionList`n`nDo you want to continue?",
            "Confirm Log Off",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
            Update-Status "Logging off users on $computerName..."
            
            if ($computerName -eq "." -or $computerName -eq $env:COMPUTERNAME) {
                # Local logoff
                foreach ($session in $activeSessions) {
                    if ($session -match '\s+(\d+)\s+') {
                        $sessionId = $Matches[1]
                        Start-Process "logoff.exe" -ArgumentList $sessionId -Wait
                    }
                }
            } else {
                # Remote logoff
                Invoke-Command -ComputerName $computerName -ScriptBlock {
                    param($sessions)
                    foreach ($session in $sessions) {
                        if ($session -match '\s+(\d+)\s+') {
                            $sessionId = $Matches[1]
                            Start-Process "logoff.exe" -ArgumentList $sessionId -Wait
                        }
                    }
                } -ArgumentList $activeSessions
            }
            
            Start-Sleep -Seconds 1
            
            # Verify logoff
            $remainingSessions = $null
            if ($computerName -eq "." -or $computerName -eq $env:COMPUTERNAME) {
                $remainingSessions = query session 2>&1
            } else {
                $remainingSessions = Invoke-Command -ComputerName $computerName -ScriptBlock { query session 2>&1 }
            }
            
            $activeRemaining = $remainingSessions | Select-Object -Skip 1 | Where-Object { 
                $_ -match '\s+(rdp-tcp|console)\s+' -and $_ -notmatch '^\s*SESSIONNAME'
            }
            
            if ($activeRemaining) {
                $outputBox.Text = "Warning: Some sessions may still be active on $computerName`:`n`n$($activeRemaining | Out-String)"
            } else {
                $outputBox.Text = "Successfully logged off all users on $computerName"
            }
        }
    }
    catch {
        $outputBox.Text = "Error managing user sessions: $($_.Exception.Message)"
    }
    Update-Status "Ready"
}

# Function to open computer management
function Open-ComputerManagement {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    Update-Status "Opening Computer Management for $computerName..."
    try {
        if ($computerName -eq "." -or $computerName -eq $env:COMPUTERNAME) {
            Start-Process "compmgmt.msc"
        } else {
            Start-Process "compmgmt.msc" -ArgumentList "/computer=$computerName"
        }
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error opening Computer Management: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
    }
    Update-Status "Ready"
}

# Function to restart computer
function Restart-TargetComputer {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    $result = [System.Windows.Forms.MessageBox]::Show(
        "Are you sure you want to restart $computerName?",
        "Confirm Restart",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )
    
    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Update-Status "Initiating restart of $computerName..."
        try {
            if ($computerName -eq "." -or $computerName -eq $env:COMPUTERNAME) {
                Restart-Computer -Force
            } else {
                Invoke-Command -ComputerName $computerName -ScriptBlock { Restart-Computer -Force }
            }
            $outputBox.Text = "Restart initiated on $computerName"
        }
        catch {
            $outputBox.Text = "Error restarting computer: $($_.Exception.Message)"
        }
    }
    Update-Status "Ready"
}

# Function to get installed applications
function Get-InstalledApplications {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    Update-Status "Getting installed applications for $computerName..."
    try {
        $script = {
            "INSTALLED APPLICATIONS`n=====================`n"
            Get-CimInstance -ClassName Win32_Product | 
                Select-Object Name, Version, Vendor, InstallDate |
                Sort-Object Name |
                Format-Table -AutoSize
                
            "`nUNINSTALL REGISTRY ENTRIES (More Complete List)`n=========================================`n"
            $paths = @(
                'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*',
                'HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*'
            )
            Get-ItemProperty $paths | 
                Where-Object DisplayName -ne $null |
                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate |
                Sort-Object DisplayName |
                Format-Table -AutoSize
        }
        
        $result = Invoke-Command -scriptBlock $script -computerName $computerName
        $outputBox.Text = $result | Out-String
    }
    catch {
        $outputBox.Text = "Error getting installed applications: $($_.Exception.Message)"
    }
    Update-Status "Ready"
}

# Function to open C$ share
function Open-AdminShare {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    try {
        $sharePath = "\\$computerName\C$"
        Update-Status "Opening $sharePath..."
        Start-Process "explorer.exe" -ArgumentList $sharePath
        Update-Status "Ready"
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Error opening C$ share: $_",
            "Error",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        Update-Status "Ready"
    }
}

# Function to remove network printers and ports
function Remove-NetworkPrintersAndPorts {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName
    )

    try {
        $script = {
            # Get all network printers
            $networkPrinters = Get-WmiObject Win32_Printer | Where-Object { $_.Network -eq $true }
            
            # Remove each network printer
            foreach ($printer in $networkPrinters) {
                try {
                    $printer.Delete()
                    Write-Host "Removed printer: $($printer.Name)"
                }
                catch {
                    Write-Warning "Failed to remove printer $($printer.Name): $_"
                }
            }

            # Get all printer ports
            $ports = Get-WmiObject Win32_TCPIPPrinterPort

            # Get ports that are actually in use by local printers
            $usedPorts = (Get-WmiObject Win32_Printer | Where-Object { $_.Network -eq $false }).PortName

            # Remove unused printer ports
            foreach ($port in $ports) {
                if ($usedPorts -notcontains $port.Name) {
                    try {
                        $port.Delete()
                        Write-Host "Removed unused port: $($port.Name)"
                    }
                    catch {
                        Write-Warning "Failed to remove port $($port.Name): $_"
                    }
                }
            }

            return $true
        }

        $result = Invoke-Command -ComputerName $ComputerName -ScriptBlock $script
        return $result
    }
    catch {
        Write-Error "Error during printer cleanup: $_"
        return $false
    }
}

# Function to handle printer cleanup from GUI
function Start-PrinterCleanup {
    if ([string]::IsNullOrWhiteSpace($computerInput.Text)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a target computer name first.",
            "No Computer Specified",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    $result = [System.Windows.Forms.MessageBox]::Show(
        "This will remove all network printers and unused printer ports on $($computerInput.Text).`n`nDo you want to continue?",
        "Confirm Printer Cleanup",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::Yes) {
        Update-Status "Starting printer cleanup on $($computerInput.Text)..."
        
        try {
            $success = Remove-NetworkPrintersAndPorts -ComputerName $computerInput.Text
            
            if ($success) {
                $outputBox.Text = "Successfully removed network printers and unused ports on $($computerInput.Text)"
                [System.Windows.Forms.MessageBox]::Show(
                    "Printer cleanup completed successfully.",
                    "Success",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Information
                )
            }
            else {
                throw "Cleanup operation returned failure"
            }
        }
        catch {
            $outputBox.Text = "Error: $_"
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to complete printer cleanup: $_",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            )
        }
        Update-Status "Ready"
    }
}

# Function to rename a remote computer
function Rename-RemoteComputer {
    param(
        [string]$ComputerName,
        [string]$NewName
    )

    try {
        # Input validation
        if ([string]::IsNullOrWhiteSpace($ComputerName) -or [string]::IsNullOrWhiteSpace($NewName)) {
            throw "Both computer name and new name are required."
        }

        # Check if new name is valid
        if ($NewName -notmatch '^[a-zA-Z0-9-]{1,15}$') {
            throw "New computer name can only contain letters, numbers, and hyphens, and must be 1-15 characters long."
        }

        # Test connection first
        if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
            throw "Cannot connect to computer '$ComputerName'"
        }

        # Get credentials if needed
        if ($ComputerName -ne $env:COMPUTERNAME -and $ComputerName -ne ".") {
            if (-not $script:credential) {
                $script:credential = $host.ui.PromptForCredential("Authentication Required", 
                    "Enter your administrative credentials", "", "NetBiosUserName")
                
                if (-not $script:credential) {
                    throw "Credentials are required for remote operations"
                }
            }
        }

        # Create the rename script block
        $scriptBlock = {
            param($newName)
            
            try {
                # Get current computer info
                $currentName = $env:COMPUTERNAME
                
                # Rename the computer
                $result = Rename-Computer -NewName $newName -Force -PassThru -ErrorAction Stop
                
                return @{
                    Success = $true
                    Message = "Computer renamed from $currentName to $newName. A restart is required to apply the change."
                }
            }
            catch {
                return @{
                    Success = $false
                    Message = "Failed to rename computer: $_"
                }
            }
        }

        # Execute the rename operation
        if ($ComputerName -eq $env:COMPUTERNAME -or $ComputerName -eq ".") {
            # For local computer, still need to run elevated
            $result = Invoke-Command -ScriptBlock {
                param($newName)
                Add-Type -TypeDefinition @"
                    using System;
                    using System.Runtime.InteropServices;
                    
                    public class ComputerName {
                        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
                        public static extern bool SetComputerNameEx(int NameType, string lpBuffer);
                    }
"@
                try {
                    # Get current computer info
                    $currentName = $env:COMPUTERNAME
                    
                    # Try using .NET method first
                    $renamed = [ComputerName]::SetComputerNameEx(5, $newName)
                    if (-not $renamed) {
                        # If .NET method fails, try PowerShell method
                        Rename-Computer -NewName $newName -Force -PassThru -ErrorAction Stop
                    }
                    
                    return @{
                        Success = $true
                        Message = "Computer renamed from $currentName to $newName. A restart is required to apply the change."
                    }
                }
                catch {
                    return @{
                        Success = $false
                        Message = "Failed to rename computer: $_"
                    }
                }
            } -ArgumentList $NewName
        }
        else {
            # For remote computer
            # Always prompt for fresh domain admin credentials for rename operation
            $script:credential = Get-Credential -Message "Enter domain admin credentials for renaming $ComputerName" -UserName "$env:USERDOMAIN\$env:USERNAME"
            if (-not $script:credential) {
                throw "Credentials are required for remote operations"
            }

            # Test credential access with explicit domain admin requirements
            try {
                Write-Host "Testing connection with credentials..."
                
                # First test basic connectivity
                if (-not (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
                    throw "Cannot ping $ComputerName"
                }
                
                # Test admin access by attempting a WMI query
                try {
                    $null = Get-WmiObject -Class Win32_ComputerSystem -ComputerName $ComputerName -Credential $script:credential -ErrorAction Stop
                }
                catch {
                    throw "Access denied. Please ensure you are using domain admin credentials."
                }
                
                # Additional verification of admin rights
                $adminTest = Invoke-Command -ComputerName $ComputerName -Credential $script:credential -ScriptBlock {
                    ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
                } -ErrorAction Stop

                if (-not $adminTest) {
                    throw "The provided credentials do not have administrative rights on $ComputerName"
                }
            }
            catch {
                $errorMsg = $_.Exception.Message
                if ($errorMsg -like "*Access is denied*") {
                    throw "Access denied. Please ensure you are using domain admin credentials."
                }
                elseif ($errorMsg -like "*The user name or password is incorrect*") {
                    throw "Invalid credentials. Please verify your username and password."
                }
                elseif ($errorMsg -like "*Cannot ping*") {
                    throw $errorMsg
                }
                else {
                    throw "Unable to access ${ComputerName}: ${errorMsg}"
                }
            }

            Write-Host "Attempting to rename computer..."
            try {
                # Get current computer name first
                $currentName = Invoke-Command -ComputerName $ComputerName -Credential $script:credential -ScriptBlock {
                    $env:COMPUTERNAME
                }

                # Perform the rename operation
                $null = Rename-Computer -ComputerName $ComputerName -NewName $NewName -DomainCredential $script:credential -Force -PassThru -ErrorAction Stop -Restart:$false
                
                $result = @{
                    Success = $true
                    Message = "Computer renamed from $currentName to $NewName. A restart is required to apply the change."
                }
            }
            catch {
                $result = @{
                    Success = $false
                    Message = "Failed to rename computer: $_"
                }
            }
        }

        return $result
    }
    catch {
        return @{
            Success = $false
            Message = "Error: $_"
        }
    }
}

# Function to handle rename operation from GUI
function Start-ComputerRename {
    # Get the current computer name
    $currentComputer = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($currentComputer)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a target computer name first.",
            "No Computer Specified",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Prompt for new name
    $newNameForm = New-Object System.Windows.Forms.Form
    $newNameForm.Text = "Rename Computer"
    $newNameForm.Size = New-Object System.Drawing.Size(400, 150)
    $newNameForm.StartPosition = "CenterParent"
    $newNameForm.FormBorderStyle = "FixedDialog"
    $newNameForm.MaximizeBox = $false
    $newNameForm.MinimizeBox = $false
    $newNameForm.BackColor = $darkBackground
    $newNameForm.ForeColor = $textColor

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter new computer name:"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(360, 20)
    $label.ForeColor = $textColor

    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 45)
    $textBox.Size = New-Object System.Drawing.Size(360, 20)
    $textBox.BackColor = $controlBackground
    $textBox.ForeColor = $textColor

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(200, 75)
    $okButton.BackColor = $accentColor
    $okButton.ForeColor = $textColor

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(290, 75)
    $cancelButton.BackColor = $accentColor
    $cancelButton.ForeColor = $textColor

    $newNameForm.Controls.AddRange(@($label, $textBox, $okButton, $cancelButton))
    $newNameForm.AcceptButton = $okButton
    $newNameForm.CancelButton = $cancelButton

    $result = $newNameForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $newName = $textBox.Text.Trim()

        # Validate the new name
        if ($newName -notmatch '^[a-zA-Z0-9-]{1,15}$') {
            [System.Windows.Forms.MessageBox]::Show(
                "New computer name can only contain letters, numbers, and hyphens, and must be 1-15 characters long.",
                "Invalid Name",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # Confirm the change
        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to rename computer '$currentComputer' to '$newName'?`nThis will require a restart to take effect.",
            "Confirm Rename",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )

        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            Update-Status "Renaming computer $currentComputer to $newName..."
            
            try {
                $result = Rename-RemoteComputer -ComputerName $currentComputer -NewName $newName

                if ($result.Success) {
                    $outputBox.Text = $result.Message
                    [System.Windows.Forms.MessageBox]::Show(
                        $result.Message,
                        "Success",
                        [System.Windows.Forms.MessageBoxButtons]::OK,
                        [System.Windows.Forms.MessageBoxIcon]::Information
                    )

                    # Offer to restart
                    $restartResult = [System.Windows.Forms.MessageBox]::Show(
                        "Would you like to restart the computer now to complete the rename?",
                        "Restart Required",
                        [System.Windows.Forms.MessageBoxButtons]::YesNo,
                        [System.Windows.Forms.MessageBoxIcon]::Question
                    )

                    if ($restartResult -eq [System.Windows.Forms.DialogResult]::Yes) {
                        Restart-TargetComputer
                    }
                }
                else {
                    throw $result.Message
                }
            }
            catch {
                $outputBox.Text = "Error: $_"
                [System.Windows.Forms.MessageBox]::Show(
                    "Failed to rename computer: $_",
                    "Error",
                    [System.Windows.Forms.MessageBoxButtons]::OK,
                    [System.Windows.Forms.MessageBoxIcon]::Error
                )
            }
            
            Update-Status "Ready"
        }
    }
}

# ==========================================
# Event Handlers
# ==========================================

$btnPing.Add_Click({ Test-ComputerPing })
$btnSysInfo.Add_Click({ Get-SystemInformation })
$btnUptime.Add_Click({ Get-ComputerUptime })
$btnUsers.Add_Click({ Get-UserSessions })
$btnPrinters.Add_Click({ Get-PrinterInfo })
$btnServices.Add_Click({ Get-RunningServices })
$btnDiskSpace.Add_Click({ Get-DiskSpaceInfo })
$btnCompMgmt.Add_Click({ Open-ComputerManagement })
$btnOpenShare.Add_Click({ Open-AdminShare })
$btnApps.Add_Click({ Get-InstalledApplications })
$btnLogOff.Add_Click({ Invoke-LogOffUsers })
$btnRestart.Add_Click({ Restart-TargetComputer })

# ==========================================
# Form Assembly and Display
# ==========================================

# Add controls to groups
$targetGroup.Controls.AddRange(@($computerLabel, $computerInput))
$buttonGroup.Controls.AddRange(@(
    $btnPing, $btnSysInfo, $btnUptime, $btnUsers, $btnPrinters, $btnApps,  # Row 1
    $btnServices, $btnDiskSpace, $btnCompMgmt, $btnOpenShare, $btnLogOff, $btnRestart,  # Row 2
    $btnPrinterCleanup, $btnRenamePC  # Row 3
    $btnPing, $btnSysInfo, $btnUptime, $btnUsers, $btnPrinters,
    $btnServices, $btnDiskSpace, $btnCompMgmt, $btnOpenShare, $btnApps, $btnLogOff, $btnRestart
))
$outputGroup.Controls.Add($outputBox)

# Add groups to form
$form.Controls.AddRange(@($targetGroup, $buttonGroup, $outputGroup, $statusStrip))

# Show the form
$form.ShowDialog()
