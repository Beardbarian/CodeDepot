# Powershell System Management Tool v2.0

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Function to test full admin rights by attempting a privileged operation
function Test-FullAdminRights {
    try {
        $key = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System"
        $null = Get-ItemProperty -Path $key -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

# Check if running as administrator and handle elevation/credentials
$currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
$principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
$adminRole = [Security.Principal.WindowsBuiltInRole]::Administrator
$hasFullRights = Test-FullAdminRights

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
elseif (-not $hasFullRights) {
    $credential = $null
    try {
        $credential = Get-Credential -Message "Your current account has limited privileges. Please enter credentials with full administrative rights."
        
        if ($null -eq $credential) {
            [System.Windows.Forms.MessageBox]::Show(
                "Full administrative credentials are required to use this tool.",
                "Full Admin Rights Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            exit
        }

        # Relaunch script with full admin credentials
        $scriptPath = $MyInvocation.MyCommand.Path
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Credential $credential -WorkingDirectory $PSScriptRoot
        exit
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Failed to authenticate with the provided credentials. Please try again with valid administrative credentials.",
            "Authentication Failed",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        exit
    }
}

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

# Layout variables
$margin = 10
$groupWidth = 880

# Create the main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Powershell System Management Tool v2.0"
$form.ClientSize = New-Object System.Drawing.Size(($groupWidth + $margin * 2), 705)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.BackColor = $darkBackground
$form.ForeColor = $textColor

# Add closing confirmation
$form.Add_FormClosing({
    param($sender, $e)
    
    # Only show prompt if it's a user closing action
    if ($e.CloseReason -eq [System.Windows.Forms.CloseReason]::UserClosing) {
        $result = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to exit?",
            "Confirm Exit",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Question
        )
        
        if ($result -eq [System.Windows.Forms.DialogResult]::No) {
            $e.Cancel = $true
        }
    }
})


# ==========================================
# Control Creation
# ==========================================

# Target Computer Input Group
$targetGroup = New-Object System.Windows.Forms.GroupBox
$targetGroup.Text = "Target Computer"
$targetGroup.Location = New-Object System.Drawing.Point($margin, 10)
$targetGroup.Size = New-Object System.Drawing.Size($groupWidth, 55)
$targetGroup.Margin = New-Object System.Windows.Forms.Padding(0)
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
$outputGroup.Location = New-Object System.Drawing.Point($margin, 235)
$outputGroup.Size = New-Object System.Drawing.Size($groupWidth, 440)  # Reduced height to make room for status bar
$outputGroup.BackColor = $panelBackground
$outputGroup.ForeColor = $textColor
$outputGroup.Padding = New-Object System.Windows.Forms.Padding(10, 20, 10, 10)

# Action Buttons Group
$buttonGroup = New-Object System.Windows.Forms.GroupBox
$buttonGroup.Text = "Actions"
$buttonGroup.Location = New-Object System.Drawing.Point($margin, 75)
$buttonGroup.Size = New-Object System.Drawing.Size($groupWidth, 150)  # Increased height for better spacing
$buttonGroup.BackColor = $panelBackground
$buttonGroup.ForeColor = $textColor

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

# Layout settings
$buttonWidth = 130
$buttonHeight = 32
$buttonSpacing = 10
$buttonsPerRow = 6

# Calculate button positions for centering
$totalButtonWidth = ($buttonWidth * $buttonsPerRow) + ($buttonSpacing * ($buttonsPerRow - 1))
$buttonStartX = [Math]::Floor(($groupWidth - $totalButtonWidth) / 2)

# Set row positions
$row1Y = 20
$row2Y = $row1Y + $buttonHeight + 10
$row3Y = $row2Y + $buttonHeight + 10

# Row 1 - System Information and Quick Access
$btnPing = New-Object System.Windows.Forms.Button
$btnPing.Text = "Ping"
$btnPing.Location = New-Object System.Drawing.Point($buttonStartX, $row1Y)
$btnPing.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnPing

$btnUsers = New-Object System.Windows.Forms.Button
$btnUsers.Text = "User Sessions"
$btnUsers.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing)), $row1Y)
$btnUsers.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnUsers

$btnConfigMgr = New-Object System.Windows.Forms.Button
$btnConfigMgr.Text = "ConfigMgr Actions"
$btnConfigMgr.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 2), $row1Y)
$btnConfigMgr.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnConfigMgr

$btnPrinters = New-Object System.Windows.Forms.Button
$btnPrinters.Text = "Printers"
$btnPrinters.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 3), $row1Y)
$btnPrinters.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnPrinters

$btnPrinterMgmt = New-Object System.Windows.Forms.Button
$btnPrinterMgmt.Text = "Add/Remove Printer"
$btnPrinterMgmt.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 4), $row1Y)
$btnPrinterMgmt.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnPrinterMgmt

$btnPrinterCleanup = New-Object System.Windows.Forms.Button
$btnPrinterCleanup.Text = "Clean Printers"
$btnPrinterCleanup.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 5), $row1Y)
$btnPrinterCleanup.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnPrinterCleanup

# Row 2 - System Management
$btnSysInfo = New-Object System.Windows.Forms.Button
$btnSysInfo.Text = "System Info"
$btnSysInfo.Location = New-Object System.Drawing.Point($buttonStartX, $row2Y)
$btnSysInfo.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnSysInfo

$btnOpenShare = New-Object System.Windows.Forms.Button
$btnOpenShare.Text = "C$ Share"
$btnOpenShare.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing)), $row2Y)
$btnOpenShare.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnOpenShare

$btnPowerStates = New-Object System.Windows.Forms.Button
$btnPowerStates.Text = "Power States"
$btnPowerStates.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 2), $row2Y)
$btnPowerStates.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnPowerStates

$btnApps = New-Object System.Windows.Forms.Button
$btnApps.Text = "Applications"
$btnApps.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 3), $row2Y)
$btnApps.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnApps

$btnServices = New-Object System.Windows.Forms.Button
$btnServices.Text = "Services"
$btnServices.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 4), $row2Y)
$btnServices.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnServices

$btnRestartService = New-Object System.Windows.Forms.Button
$btnRestartService.Text = "Restart Service"
$btnRestartService.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 5), $row2Y)
$btnRestartService.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnRestartService

# Row 3 - System Maintenance and Control
$btnCompMgmt = New-Object System.Windows.Forms.Button
$btnCompMgmt.Text = "Computer Mgmt"
$btnCompMgmt.Location = New-Object System.Drawing.Point($buttonStartX, $row3Y)
$btnCompMgmt.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnCompMgmt

$btnRenamePC = New-Object System.Windows.Forms.Button
$btnRenamePC.Text = "Rename PC"
$btnRenamePC.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing)), $row3Y)
$btnRenamePC.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnRenamePC

$btnDismRestore = New-Object System.Windows.Forms.Button
$btnDismRestore.Text = "DISM Restore"
$btnDismRestore.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 2), $row3Y)
$btnDismRestore.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnDismRestore

$btnCleanProfiles = New-Object System.Windows.Forms.Button
$btnCleanProfiles.Text = "Clean User Profiles"
$btnCleanProfiles.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 3), $row3Y)
$btnCleanProfiles.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnCleanProfiles

$btnLogOff = New-Object System.Windows.Forms.Button
$btnLogOff.Text = "Log Off Users"
$btnLogOff.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 4), $row3Y)
$btnLogOff.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnLogOff

$btnRestart = New-Object System.Windows.Forms.Button
$btnRestart.Text = "Restart PC"
$btnRestart.Location = New-Object System.Drawing.Point(($buttonStartX + ($buttonWidth + $buttonSpacing) * 5), $row3Y)
$btnRestart.Size = New-Object System.Drawing.Size($buttonWidth, $buttonHeight)
Style-Button $btnRestart

# Calculate the same width as button rows
$totalButtonsWidth = ($buttonWidth * 6) + ($buttonSpacing * 5)

$outputBox = New-Object System.Windows.Forms.RichTextBox
$outputBox.Location = New-Object System.Drawing.Point(20, 20)
$outputBox.Size = New-Object System.Drawing.Size(840, 410)
$outputBox.Font = New-Object System.Drawing.Font("Consolas", 10)
$outputBox.BackColor = [System.Drawing.Color]::Black
$outputBox.ForeColor = [System.Drawing.Color]::White
$outputBox.ReadOnly = $true
$outputBox.MultiLine = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Margin = New-Object System.Windows.Forms.Padding(0)

# Status Bar with dark theme
$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusStrip.BackColor = $panelBackground
$statusStrip.SizingGrip = $false
$statusStrip.Dock = [System.Windows.Forms.DockStyle]::Bottom

$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Text = "Ready"
$statusLabel.ForeColor = $subTextColor
$statusLabel.Spring = $true  # Makes the label expand to fill space
$statusLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
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
            $bios = Get-CimInstance Win32_BIOS
            
            # Calculate uptime
            $uptime = (Get-Date) - $os.LastBootUpTime
            $uptimeString = "{0} days, {1} hours, {2} minutes" -f $uptime.Days, $uptime.Hours, $uptime.Minutes
            
            "SYSTEM INFORMATION`n==================`n"
            "Computer Name    : $env:COMPUTERNAME"
            "OS Version       : $($os.Caption)"
            "OS Build         : $($os.BuildNumber)"
            "OS Architecture  : $($os.OSArchitecture)"
            "Manufacturer     : $($cs.Manufacturer)"
            "Model            : $($cs.Model)"
            "Serial Number    : $($bios.SerialNumber)"
            "BIOS Version     : $($bios.SMBIOSBIOSVersion)"
            "Processor       : $($proc.Name)"
            "CPU Cores       : $($proc.NumberOfCores)"
            "CPU Threads     : $($proc.NumberOfLogicalProcessors)"
            "Memory (GB)     : $([math]::Round($cs.TotalPhysicalMemory/1GB, 2))"
            "Free Memory(GB) : $([math]::Round($os.FreePhysicalMemory/1MB, 2))"
            "Last Boot Time  : $($os.LastBootUpTime)"
            "System Uptime   : $uptimeString"
            "`nDISK INFORMATION`n==================`n"
            Get-CimInstance Win32_LogicalDisk -Filter 'DriveType=3' | ForEach-Object {
                "Drive: $($_.DeviceID)"
                "Label: $($_.VolumeName)"
                "Size (GB): $([math]::Round($_.Size/1GB, 2))"
                "Free (GB): $([math]::Round($_.FreeSpace/1GB, 2))"
                "Free (%): $([math]::Round(($_.FreeSpace/$_.Size)*100, 1))"
                "------------------------"
            }
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

# Invoke Configuration Manager Actions
function Invoke-ConfigMgrActions {
    Update-Status "Triggering Configuration Manager actions..."
    
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { $computerName = "." }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }
    
    try {
        $script = {
            $actions = @{
                'Machine Policy Retrieval'= '{00000000-0000-0000-0000-000000000021}'
                'Discovery Data Collection'= '{00000000-0000-0000-0000-000000000003}'
                'File Collection'= '{00000000-0000-0000-0000-000000000010}'
                'Software Metering'= '{00000000-0000-0000-0000-000000000031}'
                'Software Updates Deployment'= '{00000000-0000-0000-0000-000000000032}'
                'Windows Installer Source'= '{00000000-0000-0000-0000-000000000021}'
                'Application Deployment'= '{00000000-0000-0000-0000-000000000121}'
                'Hardware Inventory'= '{00000000-0000-0000-0000-000000000001}'
                'Update Deployment'= '{00000000-0000-0000-0000-000000000108}'
                'Software Updates Scan'= '{00000000-0000-0000-0000-000000000113}'
                'Software Inventory'= '{00000000-0000-0000-0000-000000000002}'
            }

            "CONFIGURATION MANAGER ACTIONS`n===========================`n"
            $results = @()
            foreach ($action in $actions.GetEnumerator()) {
                try {
                    $null = Invoke-WmiMethod -Namespace "root\ccm" -Class "SMS_Client" -Name "TriggerSchedule" -ArgumentList $action.Value -ErrorAction Stop
                    $results += "✓ Successfully triggered: $($action.Key)"
                }
                catch {
                    $results += "✗ Failed to trigger: $($action.Key) - $($_.Exception.Message)"
                }
            }
            $results | ForEach-Object { $_ + "`n" }
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
    
    # Show warning message
    $warningResult = [System.Windows.Forms.MessageBox]::Show(
        "This may take a few minutes, and will appear frozen while collecting application info. Click OK to continue and please wait.",
        "Please Wait",
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    if ($warningResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
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
            Write-Host "Starting printer cleanup..."
            
            # Get all network printers except SecurePrint
            $networkPrinters = Get-Printer | Where-Object { 
                ($_.Type -eq "Connection" -or $_.PortName -like "\\*" -or $_.PortName -like "IP_*" -or $_.PortName -match "^PB.*\d+$") -and 
                $_.Name -ne "SecurePrint" -and 
                $_.Name -ne "Microsoft Print to PDF" -and 
                $_.Name -ne "Microsoft XPS Document Writer" -and
                $_.Name -ne "Midmark PDF Converter"
            }
            
            Write-Host "`nFound $($networkPrinters.Count) network printers to remove."
            
            # Remove each network printer
            foreach ($printer in $networkPrinters) {
                try {
                    Write-Host "Removing printer: $($printer.Name)"
                    Remove-Printer -Name $printer.Name -ErrorAction Stop
                    Write-Host "Successfully removed printer: $($printer.Name)"
                }
                catch {
                    Write-Warning "Failed to remove printer $($printer.Name): $_"
                }
            }

            # Get all ports
            $ports = Get-PrinterPort
            
            # Get printers to check which ports are in use
            $activePrinters = Get-Printer
            $usedPorts = $activePrinters | Select-Object -ExpandProperty PortName
            
            Write-Host "`nCleaning up unused printer ports..."
            
            # Remove unused ports (skip standard ports and SecurePrint port)
            $standardPorts = @("FILE:", "LPT1:", "LPT2:", "LPT3:", "COM1:", "COM2:", "COM3:", "COM4:", "PORTPROMPT:", "NUL:")
            foreach ($port in $ports) {
                if ($port.Name -notin $standardPorts -and $port.Name -notin $usedPorts) {
                    try {
                        Write-Host "Removing unused port: $($port.Name)"
                        Remove-PrinterPort -Name $port.Name -ErrorAction Stop
                        Write-Host "Successfully removed port: $($port.Name)"
                    }
                    catch {
                        Write-Warning "Failed to remove port $($port.Name): $_"
                    }
                }
            }

            Write-Host "`nCleanup completed."
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

# Function to run DISM restore health
function Start-DismRestore {
    $targetComputer = $computerInput.Text
    if ([string]::IsNullOrEmpty($targetComputer)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a computer name.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    
    # Show warning message
    $warningResult = [System.Windows.Forms.MessageBox]::Show(
        "A blank window will appear as this runs. This may take a few minutes, and will appear frozen while DISM is running. Click OK to continue and please wait.",
        "Please Wait",
        [System.Windows.Forms.MessageBoxButtons]::OKCancel,
        [System.Windows.Forms.MessageBoxIcon]::Information
    )
    
    if ($warningResult -eq [System.Windows.Forms.DialogResult]::Cancel) {
        Update-Status "Ready"
        return
    }

    Update-Status "Running DISM restore health on $targetComputer..."
    $outputBox.Clear()
    
    try {
        $result = Invoke-Command -ComputerName $targetComputer -ScriptBlock {
            $output = New-Object System.Collections.ArrayList
            
            # Start DISM with redirected output
            $dismProcess = Start-Process -FilePath "DISM.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth" -NoNewWindow -PassThru -Wait -RedirectStandardOutput "$env:TEMP\dismout.txt" -RedirectStandardError "$env:TEMP\dismerr.txt"
            
            # Read the output files
            if (Test-Path "$env:TEMP\dismout.txt") {
                $output.AddRange([System.IO.File]::ReadAllLines("$env:TEMP\dismout.txt"))
                Remove-Item "$env:TEMP\dismout.txt" -Force
            }
            if (Test-Path "$env:TEMP\dismerr.txt") {
                $errContent = [System.IO.File]::ReadAllLines("$env:TEMP\dismerr.txt")
                if ($errContent) {
                    $output.Add("`nErrors:")
                    $output.AddRange($errContent)
                }
                Remove-Item "$env:TEMP\dismerr.txt" -Force
            }
            
            return @{
                Output = $output
                ExitCode = $dismProcess.ExitCode
            }
        } -ErrorAction Stop

        # Display output
        foreach ($line in $result.Output) {
            $outputBox.AppendText("$line`n")
        }

        if ($result.ExitCode -eq 0) {
            $outputBox.AppendText("`nDISM restore health completed successfully.`n")
            Update-Status "DISM restore health completed successfully on $targetComputer"
        } else {
            $outputBox.AppendText("`nDISM restore health completed with exit code: $($result.ExitCode)`n")
            Update-Status "DISM restore health completed with errors on $targetComputer"
        }
    }
    catch {
        $outputBox.AppendText("Error: $_`n")
        Update-Status "Error running DISM restore health"
    }
}

# Function to show power states
function Show-PowerStates {
    $targetComputer = $computerInput.Text
    if ([string]::IsNullOrEmpty($targetComputer)) {
        [System.Windows.Forms.MessageBox]::Show("Please enter a computer name.", "Input Required", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    Update-Status "Retrieving power states from $targetComputer..."
    $outputBox.Clear()
    
    try {
        $powerStates = Invoke-Command -ComputerName $targetComputer -ScriptBlock {
            $powerCfg = & powercfg /list
            $powerPlan = & powercfg /query
            $sleepStates = & powercfg /availablesleepstates
            return @{
                PowerCfg = $powerCfg
                PowerPlan = $powerPlan
                SleepStates = $sleepStates
            }
        } -ErrorAction Stop

        $outputBox.AppendText("=== Current Power Configuration ===`n")
        $outputBox.AppendText(($powerStates.PowerCfg | Out-String))
        $outputBox.AppendText("`n=== Available Sleep States ===`n")
        $outputBox.AppendText(($powerStates.SleepStates | Out-String))
        $outputBox.AppendText("`n=== Detailed Power Plan Settings ===`n")
        $outputBox.AppendText(($powerStates.PowerPlan | Out-String))
        
        Update-Status "Power states retrieved successfully"
    }
    catch {
        $outputBox.AppendText("Error: $_`n")
        Update-Status "Error retrieving power states"
    }
}

# Function to restart a service
function Restart-RemoteService {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) { 
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a computer name.",
            "Input Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return 
    }

    # Prompt for service name
    $serviceForm = New-Object System.Windows.Forms.Form
    $serviceForm.Text = "Restart Service"
    $serviceForm.Size = New-Object System.Drawing.Size(400, 200)
    $serviceForm.StartPosition = "CenterParent"
    $serviceForm.FormBorderStyle = "FixedDialog"
    $serviceForm.MaximizeBox = $false
    $serviceForm.MinimizeBox = $false
    $serviceForm.BackColor = $darkBackground
    $serviceForm.ForeColor = $textColor

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Enter service name (e.g., spooler, wuauserv):"
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(360, 20)
    $label.ForeColor = $textColor

    $serviceBox = New-Object System.Windows.Forms.ComboBox
    $serviceBox.Location = New-Object System.Drawing.Point(10, 45)
    $serviceBox.Size = New-Object System.Drawing.Size(360, 20)
    $serviceBox.BackColor = $controlBackground
    $serviceBox.ForeColor = $textColor
    $serviceBox.AutoCompleteMode = 'SuggestAppend'
    $serviceBox.AutoCompleteSource = 'ListItems'

    # Get list of services for autocomplete
    try {
        $services = Invoke-Command -ComputerName $computerName -ScriptBlock {
            Get-Service | Select-Object DisplayName, Name | Sort-Object DisplayName
        }
        foreach ($service in $services) {
            $serviceBox.Items.Add("$($service.DisplayName) ($($service.Name))")
        }
    }
    catch {
        $outputBox.AppendText("Error getting service list: $_`n")
    }

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $okButton.Location = New-Object System.Drawing.Point(200, 120)
    $okButton.BackColor = $accentColor
    $okButton.ForeColor = $textColor

    $cancelButton = New-Object System.Windows.Forms.Button
    $cancelButton.Text = "Cancel"
    $cancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $cancelButton.Location = New-Object System.Drawing.Point(290, 120)
    $cancelButton.BackColor = $accentColor
    $cancelButton.ForeColor = $textColor

    $serviceForm.Controls.AddRange(@($label, $serviceBox, $okButton, $cancelButton))
    $serviceForm.AcceptButton = $okButton
    $serviceForm.CancelButton = $cancelButton

    $result = $serviceForm.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedService = $serviceBox.Text
        if ([string]::IsNullOrWhiteSpace($selectedService)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter a service name.",
                "Input Required",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            )
            return
        }

        # Extract service name from selection
        if ($selectedService -match '\((.*?)\)$') {
            $serviceName = $matches[1]
        } else {
            $serviceName = $selectedService
        }

        $confirmResult = [System.Windows.Forms.MessageBox]::Show(
            "Are you sure you want to restart the '$serviceName' service on $computerName?",
            "Confirm Service Restart",
            [System.Windows.Forms.MessageBoxButtons]::YesNo,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )

        if ($confirmResult -eq [System.Windows.Forms.DialogResult]::Yes) {
            Update-Status "Restarting $serviceName service on $computerName..."
            $outputBox.Clear()
            
            try {
                $result = Invoke-Command -ComputerName $computerName -ScriptBlock {
                    param($svcName)
                    try {
                        $service = Get-Service -Name $svcName
                        $outputBuilder = New-Object System.Text.StringBuilder
                        
                        $outputBuilder.AppendLine("Current Status: $($service.Status)")
                        $outputBuilder.AppendLine("Stopping service...")
                        $service | Stop-Service -Force -ErrorAction Stop
                        Start-Sleep -Seconds 2
                        
                        $outputBuilder.AppendLine("Starting service...")
                        $service | Start-Service -ErrorAction Stop
                        Start-Sleep -Seconds 2
                        
                        $service = Get-Service -Name $svcName
                        $outputBuilder.AppendLine("New Status: $($service.Status)")
                        
                        return $outputBuilder.ToString()
                    }
                    catch {
                        throw $_
                    }
                } -ArgumentList $serviceName -ErrorAction Stop

                $outputBox.AppendText($result)
                Update-Status "Service restart completed successfully"
            }
            catch {
                $outputBox.AppendText("Error restarting service: $_`n")
                Update-Status "Error restarting service"
            }
        }
    }
}

# Function to clean old user profiles
function Remove-OldUserProfiles {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a computer name.",
            "Input Required",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }

    # Show warning message
    $warningResult = [System.Windows.Forms.MessageBox]::Show(
        "This will remove user profiles older than 7 days on $computerName.`n`n" +
        "IMPORTANT:`n" +
        "- Currently logged in profiles will be skipped`n" +
        "- Special system profiles will be preserved`n" +
        "- This action cannot be undone`n`n" +
        "Do you want to continue?",
        "Warning - Profile Cleanup",
        [System.Windows.Forms.MessageBoxButtons]::YesNo,
        [System.Windows.Forms.MessageBoxIcon]::Warning
    )

    if ($warningResult -eq [System.Windows.Forms.DialogResult]::No) {
        return
    }

    Update-Status "Scanning for old profiles on $computerName..."
    $outputBox.Clear()

    try {
        $result = Invoke-Command -ComputerName $computerName -ScriptBlock {
            $output = New-Object System.Collections.ArrayList
            
            # Get current date and calculate cutoff date (7 days ago)
            $cutoffDate = (Get-Date).AddDays(-7)
            
            # Get all user profiles
            $profiles = Get-CimInstance -ClassName Win32_UserProfile | 
                Where-Object { 
                    -not $_.Special -and                    # Skip special profiles
                    -not $_.Loaded -and                     # Skip logged-in profiles
                    $_.LastUseTime -lt $cutoffDate -and     # Older than 7 days
                    $_.LocalPath -notlike '*systemprofile*' -and
                    $_.LocalPath -notlike '*ServiceAccount*' -and
                    $_.LocalPath -notlike '*NetworkService*' -and
                    $_.LocalPath -notlike '*LocalService*' -and
                    $_.LocalPath -notlike '*system32*'
                }

            if (-not $profiles) {
                $output.Add("No eligible profiles found for removal.")
                return $output
            }

            $output.Add("The following profiles will be removed:")
            $output.Add("=====================================")
            
            foreach ($profile in $profiles) {
                $username = Split-Path $profile.LocalPath -Leaf
                $lastUse = $profile.LastUseTime
                $output.Add("Profile: $username")
                $output.Add("Last Used: $lastUse")
                $output.Add("Path: $($profile.LocalPath)")
                
                try {
                    $profile | Remove-CimInstance -ErrorAction Stop
                    $output.Add("Status: Removed successfully")
                }
                catch {
                    $output.Add("Status: Failed to remove - $($_.Exception.Message)")
                }
                $output.Add("-------------------------------------")
            }

            return $output
        }

        foreach ($line in $result) {
            $outputBox.AppendText("$line`n")
        }
        
        Update-Status "Profile cleanup completed"
    }
    catch {
        $outputBox.AppendText("Error during profile cleanup: $_`n")
        Update-Status "Error during profile cleanup"
    }
}

# Function to manage printers
function Manage-PrinterSetup {
    $computerName = $computerInput.Text
    if ([string]::IsNullOrWhiteSpace($computerName)) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter a computer name first.",
            "No Computer Specified",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        return
    }
    
    if (-not (Test-RemoteConnection $computerName)) {
        Update-Status "Ready"
        return
    }

    # First form - Choose operation
    $choiceForm = New-Object System.Windows.Forms.Form
    $choiceForm.Text = "Printer Management - Select Operation"
    $choiceForm.Size = New-Object System.Drawing.Size(300, 150)
    $choiceForm.StartPosition = "CenterParent"
    $choiceForm.FormBorderStyle = "FixedDialog"
    $choiceForm.MaximizeBox = $false
    $choiceForm.MinimizeBox = $false
    $choiceForm.BackColor = $darkBackground
    $choiceForm.ForeColor = $textColor

    $addButton = New-Object System.Windows.Forms.Button
    $addButton.Location = New-Object System.Drawing.Point(20, 20)
    $addButton.Size = New-Object System.Drawing.Size(240, 30)
    $addButton.Text = "Add Printer"
    $addButton.BackColor = $accentColor
    $addButton.ForeColor = $textColor
    $addButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $addButton.DialogResult = [System.Windows.Forms.DialogResult]::Yes

    $removeButton = New-Object System.Windows.Forms.Button
    $removeButton.Location = New-Object System.Drawing.Point(20, 60)
    $removeButton.Size = New-Object System.Drawing.Size(240, 30)
    $removeButton.Text = "Remove Printer"
    $removeButton.BackColor = $accentColor
    $removeButton.ForeColor = $textColor
    $removeButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $removeButton.DialogResult = [System.Windows.Forms.DialogResult]::No

    $choiceForm.Controls.AddRange(@($addButton, $removeButton))
    $choiceForm.AcceptButton = $addButton
    $choiceForm.CancelButton = $removeButton

    $choice = $choiceForm.ShowDialog()
    $choiceForm.Dispose()

    if ($choice -eq [System.Windows.Forms.DialogResult]::Yes) {
        # Add Printer Form
        $addForm = New-Object System.Windows.Forms.Form
        $addForm.Text = "Add Printer"
        $addForm.Size = New-Object System.Drawing.Size(400, 300)
        $addForm.StartPosition = "CenterParent"
        $addForm.FormBorderStyle = "FixedDialog"
        $addForm.MaximizeBox = $false
        $addForm.MinimizeBox = $false
        $addForm.BackColor = $darkBackground
        $addForm.ForeColor = $textColor

        # Printer Name
        $nameLabel = New-Object System.Windows.Forms.Label
        $nameLabel.Text = "Printer Name:"
        $nameLabel.Location = New-Object System.Drawing.Point(20, 20)
        $nameLabel.AutoSize = $true
        $nameLabel.ForeColor = $textColor

        $nameTextBox = New-Object System.Windows.Forms.TextBox
        $nameTextBox.Location = New-Object System.Drawing.Point(20, 45)
        $nameTextBox.Size = New-Object System.Drawing.Size(340, 25)
        $nameTextBox.BackColor = $controlBackground
        $nameTextBox.ForeColor = $textColor

        # Port/IP
        $portLabel = New-Object System.Windows.Forms.Label
        $portLabel.Text = "Port/IP Address:"
        $portLabel.Location = New-Object System.Drawing.Point(20, 80)
        $portLabel.AutoSize = $true
        $portLabel.ForeColor = $textColor

        $portTextBox = New-Object System.Windows.Forms.TextBox
        $portTextBox.Location = New-Object System.Drawing.Point(20, 105)
        $portTextBox.Size = New-Object System.Drawing.Size(340, 25)
        $portTextBox.BackColor = $controlBackground
        $portTextBox.ForeColor = $textColor

        # Driver
        $driverLabel = New-Object System.Windows.Forms.Label
        $driverLabel.Text = "Select Driver:"
        $driverLabel.Location = New-Object System.Drawing.Point(20, 140)
        $driverLabel.AutoSize = $true
        $driverLabel.ForeColor = $textColor

        $driverCombo = New-Object System.Windows.Forms.ComboBox
        $driverCombo.Location = New-Object System.Drawing.Point(20, 165)
        $driverCombo.Size = New-Object System.Drawing.Size(340, 25)
        $driverCombo.BackColor = $controlBackground
        $driverCombo.ForeColor = $textColor
        $driverCombo.DropDownStyle = "DropDownList"

        # Get installed drivers
        try {
            $drivers = Invoke-Command -ComputerName $computerName -ScriptBlock {
                Get-PrinterDriver | Select-Object -ExpandProperty Name
            }
            $driverCombo.Items.AddRange($drivers)
            if ($driverCombo.Items.Count -gt 0) {
                $driverCombo.SelectedIndex = 0
            }
        }
        catch {
            $outputBox.Text = "Error getting printer drivers: $_"
            return
        }

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(20, 210)
        $okButton.Size = New-Object System.Drawing.Size(340, 30)
        $okButton.Text = "Add Printer"
        $okButton.BackColor = $accentColor
        $okButton.ForeColor = $textColor
        $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $addForm.Controls.AddRange(@(
            $nameLabel, $nameTextBox,
            $portLabel, $portTextBox,
            $driverLabel, $driverCombo,
            $okButton
        ))
        $addForm.AcceptButton = $okButton

        $result = $addForm.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                if ([string]::IsNullOrWhiteSpace($nameTextBox.Text) -or 
                    [string]::IsNullOrWhiteSpace($portTextBox.Text)) {
                    throw "Please fill in all fields"
                }

                $cmdResult = Invoke-Command -ComputerName $computerName -ScriptBlock {
                    param($name, $portName, $driverName)
                    try {
                        if (-not (Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue)) {
                            Add-PrinterPort -Name $portName -PrinterHostAddress $portName -ErrorAction Stop
                        }
                        Add-Printer -Name $name -DriverName $driverName -PortName $portName -ErrorAction Stop
                    }
                    catch {
                        throw $_
                    }
                } -ArgumentList $nameTextBox.Text, $portTextBox.Text, $driverCombo.SelectedItem

                $outputBox.Text = $cmdResult
            }
            catch {
                $outputBox.Text = "Error: $_"
            }
        }
        $addForm.Dispose()
    }
    elseif ($choice -eq [System.Windows.Forms.DialogResult]::No) {
        # Remove Printer Form
        $removeForm = New-Object System.Windows.Forms.Form
        $removeForm.Text = "Remove Printer"
        $removeForm.Size = New-Object System.Drawing.Size(400, 150)
        $removeForm.StartPosition = "CenterParent"
        $removeForm.FormBorderStyle = "FixedDialog"
        $removeForm.MaximizeBox = $false
        $removeForm.MinimizeBox = $false
        $removeForm.BackColor = $darkBackground
        $removeForm.ForeColor = $textColor

        $nameLabel = New-Object System.Windows.Forms.Label
        $nameLabel.Text = "Printer Name:"
        $nameLabel.Location = New-Object System.Drawing.Point(20, 20)
        $nameLabel.AutoSize = $true
        $nameLabel.ForeColor = $textColor

        $nameTextBox = New-Object System.Windows.Forms.TextBox
        $nameTextBox.Location = New-Object System.Drawing.Point(20, 45)
        $nameTextBox.Size = New-Object System.Drawing.Size(340, 25)
        $nameTextBox.BackColor = $controlBackground
        $nameTextBox.ForeColor = $textColor

        $okButton = New-Object System.Windows.Forms.Button
        $okButton.Location = New-Object System.Drawing.Point(20, 80)
        $okButton.Size = New-Object System.Drawing.Size(340, 30)
        $okButton.Text = "Remove Printer"
        $okButton.BackColor = $accentColor
        $okButton.ForeColor = $textColor
        $okButton.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $okButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $removeForm.Controls.AddRange(@($nameLabel, $nameTextBox, $okButton))
        $removeForm.AcceptButton = $okButton

        $result = $removeForm.ShowDialog()
        if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
            try {
                if ([string]::IsNullOrWhiteSpace($nameTextBox.Text)) {
                    throw "Please enter a printer name"
                }

                $cmdResult = Invoke-Command -ComputerName $computerName -ScriptBlock {
                    param($name)
                    try {
                        Remove-Printer -Name $name -ErrorAction Stop
                    }
                    catch {
                        throw $_
                    }
                } -ArgumentList $nameTextBox.Text

                $outputBox.Text = $cmdResult
            }
            catch {
                $outputBox.Text = "Error: $_"
            }
        }
        $removeForm.Dispose()
    }

    Update-Status "Ready"
}

# ==========================================
# Event Handlers
# ==========================================

# Row 1
$btnPing.Add_Click({ Test-ComputerPing })
$btnUsers.Add_Click({ Get-UserSessions })
$btnConfigMgr.Add_Click({ Invoke-ConfigMgrActions })
$btnPrinters.Add_Click({ Get-PrinterInfo })
$btnPrinterMgmt.Add_Click({ Manage-PrinterSetup })
$btnPrinterCleanup.Add_Click({ Start-PrinterCleanup })
# Row 2
$btnSysInfo.Add_Click({ Get-SystemInformation })
$btnOpenShare.Add_Click({ Open-AdminShare })
$btnPowerStates.Add_Click({ Show-PowerStates })
$btnApps.Add_Click({ Get-InstalledApplications })
$btnServices.Add_Click({ Get-RunningServices })
$btnRestartService.Add_Click({ Restart-RemoteService })
# Row 3
$btnCompMgmt.Add_Click({ Open-ComputerManagement })
$btnRenamePC.Add_Click({ Start-ComputerRename })
$btnDismRestore.Add_Click({ Start-DismRestore })
$btnCleanProfiles.Add_Click({ Remove-OldUserProfiles })
$btnLogOff.Add_Click({ Invoke-LogOffUsers })
$btnRestart.Add_Click({ Restart-TargetComputer })

# ==========================================
# Form Assembly and Display
# ==========================================

# Add controls to groups
$targetGroup.Controls.AddRange(@($computerLabel, $computerInput))
$buttonGroup.Controls.AddRange(@(
    $btnPing, $btnUsers, $btnConfigMgr, $btnPrinters, $btnPrinterMgmt, $btnPrinterCleanup,  # Row 1
    $btnSysInfo, $btnOpenShare, $btnPowerStates, $btnApps, $btnServices, $btnRestartService,  # Row 2
    $btnCompMgmt, $btnRenamePC, $btnDismRestore, $btnCleanProfiles, $btnLogOff, $btnRestart  # Row 3
))
$outputGroup.Controls.Add($outputBox)

# Add all controls to form
$form.Controls.AddRange(@($targetGroup, $buttonGroup, $outputGroup, $statusStrip))

# Show the form
[void]$form.ShowDialog()
