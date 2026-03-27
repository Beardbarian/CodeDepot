# Check if the Office suite has landed
$OfficeInstalled = Test-Path "C:\Program Files\Microsoft Office\root\Office16\winword.exe"

# Check the 64-bit registry directly for success flag
$key = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
$subKey = $key.OpenSubKey("SOFTWARE\MyCustomConfig")
$ConfigApplied = if ($subKey) { $subKey.GetValue("ConfigApplied") } else { $null }

if ($OfficeInstalled -and ($ConfigApplied -eq 1)) {
    # Intune sees this and stops trying to install/run the script
    Write-Output "Detected"
    exit 0
} else {
    # Intune sees this and triggers (or re-triggers) the installation
    exit 1
}
