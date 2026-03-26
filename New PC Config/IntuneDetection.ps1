# 1. Check if the Office Click-to-Run service exists
$OfficePath = "C:\Program Files\Microsoft Office\root\Office16\winword.exe"
$OfficeInstalled = Test-Path $OfficePath

# 2. Check for the Entra ID (Azure AD) Identity Cache
# This ensures the BrokerPlugin has at least initialized once for the user
$IdentityCache = "$env:LOCALAPPDATA\Microsoft\IdentityCache"
$IdentityReady = Test-Path $IdentityCache

# 3. Check if the "New PC Config" has already run (Custom Flag)
# We will have the main script create this key at the end to prevent loops.
$ConfigFlag = Get-ItemProperty -Path "HKLM:\SOFTWARE\MyCustomConfig" -Name "ConfigApplied" -ErrorAction SilentlyContinue

if ($OfficeInstalled -and $IdentityReady -and (-not $ConfigFlag)) {
    # If Office is there and Identity is ready, but we haven't run yet: 
    # Return nothing (or Write-Error) to tell Intune it's NOT detected yet, so it RUNS the app.
    exit 1
} elseif ($ConfigFlag.ConfigApplied -eq 1) {
    # If the flag exists, the "App" is detected as installed/finished.
    Write-Output "Configuration Verified"
    exit 0
} else {
    # Otherwise, keep waiting.
    exit 1
}
