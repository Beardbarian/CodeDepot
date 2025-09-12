param(
    [Parameter(Mandatory=$false)]
    [string]$BroadcastAddress = "255.255.255.255",
    
    [Parameter(Mandatory=$false)]
    [int]$Port = 9
)

function Get-MacFromIP {
    param([string]$IPAddress)
    
    Write-Host "NEW VERSION - Resolving MAC for IP: $IPAddress" -ForegroundColor Yellow
    
    # Ping to populate ARP table
    $null = Test-Connection -ComputerName $IPAddress -Count 1 -Quiet -ErrorAction SilentlyContinue
    
    # Search ARP table
    $arpOutput = arp -a
    foreach ($line in $arpOutput) {
        if ($line -match [regex]::Escape($IPAddress)) {
            Write-Host "NEW VERSION - ARP entry: $line" -ForegroundColor Gray
            
            # Extract MAC and clean it immediately
            if ($line -match '([0-9a-fA-F]{2}[:-][0-9a-fA-F]{2}[:-][0-9a-fA-F]{2}[:-][0-9a-fA-F]{2}[:-][0-9a-fA-F]{2}[:-][0-9a-fA-F]{2})') {
                $foundMac = $matches[1]
                $cleanMac = $foundMac -replace '-', '' -replace ':', ''
                $cleanMac = $cleanMac.ToUpper()
                
                Write-Host "NEW VERSION - Found: $foundMac -> Cleaned: $cleanMac" -ForegroundColor Green
                return $cleanMac
            }
        }
    }
    
    return $null
}

function Validate-MacAddress {
    param([string]$MacAddress)
    
    # Remove separators and convert to uppercase
    $cleanMac = $MacAddress -replace '-', '' -replace ':', '' -replace '\s', ''
    $cleanMac = $cleanMac.ToUpper()
    
    # Check if valid MAC (12 hex characters)
    if ($cleanMac -match '^[0-9A-F]{12}$') {
        return $cleanMac
    }
    
    return $null
}

function Send-WakeOnLan {
    param(
        [string]$MacAddress,
        [string]$BroadcastIP,
        [int]$UdpPort
    )
    
    try {
        Write-Host "NEW VERSION - Sending WOL with MAC: $MacAddress" -ForegroundColor Gray
        
        # Convert MAC to byte array
        $macBytes = @()
        for ($i = 0; $i -lt 12; $i += 2) {
            $macBytes += [Convert]::ToByte($MacAddress.Substring($i, 2), 16)
        }
        
        # Create magic packet
        $magicPacket = @()
        
        # 6 bytes of 0xFF
        for ($i = 0; $i -lt 6; $i++) {
            $magicPacket += 0xFF
        }
        
        # MAC address 16 times
        for ($i = 0; $i -lt 16; $i++) {
            $magicPacket += $macBytes
        }
        
        # Send UDP packet
        $udpClient = New-Object System.Net.Sockets.UdpClient
        $udpClient.Connect([System.Net.IPAddress]::Parse($BroadcastIP), $UdpPort)
        $bytesSent = $udpClient.Send($magicPacket, $magicPacket.Length)
        $udpClient.Close()
        
        Write-Host "Magic packet sent successfully!" -ForegroundColor Green
        Write-Host "  Target MAC: $MacAddress" -ForegroundColor Cyan
        Write-Host "  Broadcast: $BroadcastIP" -ForegroundColor Cyan
        Write-Host "  Port: $UdpPort" -ForegroundColor Cyan
        Write-Host "  Bytes sent: $bytesSent" -ForegroundColor Cyan
        
        return $true
    }
    catch {
        Write-Host "Error sending magic packet: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Test-IsIPAddress {
    param([string]$Address)
    
    $ipRegex = '^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$'
    return $Address -match $ipRegex
}

# Main execution
$Target = Read-Host "Please enter IP or MAC address"

if (-not $Target.Trim()) {
    Write-Host "No target specified. Exiting." -ForegroundColor Red
    exit 1
}

Write-Host "NEW VERSION - Wake-on-LAN Script Starting..." -ForegroundColor Green
Write-Host "==============================" -ForegroundColor Green

# Process target
$macAddress = $null

if (Test-IsIPAddress -Address $Target) {
    Write-Host "Target is IP address: $Target" -ForegroundColor Yellow
    $macAddress = Get-MacFromIP -IPAddress $Target
} else {
    Write-Host "Target is MAC address: $Target" -ForegroundColor Yellow
    $macAddress = Validate-MacAddress -MacAddress $Target
}

if (-not $macAddress) {
    Write-Host "Invalid or unresolvable target: $Target" -ForegroundColor Red
    exit 1
}

Write-Host "Final MAC address to use: $macAddress" -ForegroundColor Cyan

# Send Wake-on-LAN
$success = Send-WakeOnLan -MacAddress $macAddress -BroadcastIP $BroadcastAddress -UdpPort $Port

if ($success) {
    Write-Host ""
    Write-Host "Wake-on-LAN packet sent successfully!" -ForegroundColor Green
} else {
    Write-Host ""
    Write-Host "Failed to send Wake-on-LAN packet." -ForegroundColor Red
    exit 1
}
