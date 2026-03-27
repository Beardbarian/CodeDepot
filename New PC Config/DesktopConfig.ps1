# 1. Create Detection Flag IMMEDIATELY (HKLM)
# We use explicit 64-bit registry calls to ensure Intune (32-bit) doesn't redirect this to WOW6432Node
$RegistryPath = "SOFTWARE\MyCustomConfig"
$RegistryKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
$SubKey = $RegistryKey.CreateSubKey($RegistryPath)
$SubKey.SetValue("ConfigApplied", 1, [Microsoft.Win32.RegistryValueKind]::DWord)
$SubKey.Close()
$RegistryKey.Close()

# 2. Provisioned App Removal
$AppsToRemove = @(
    "*Clipchamp.Clipchamp*", 
    "*bing*", 
    "*Microsoft.GetStarted*", 
    "*Microsoft.Messaging*",
    "*Microsoft.Microsoft3DViewer*", 
    "*Microsoft.MicrosoftOfficeHub*", 
    "*Microsoft.MicrosoftSolitaireCollection*", 
    "*Microsoft.MixedReality.Portal*",
    "*Microsoft.News*", 
    "*Microsoft.OneConnect*", 
    "*Microsoft.People*",
    "*Microsoft.PowerAutomateDesktop*", 
    "*Microsoft.SkypeApp*",
    "*microsoft.windowscommunicationsapps*", 
    "*Microsoft.WindowsFeedbackHub*",
    "*Microsoft.WindowsMaps*", 
    "*Microsoft.YourPhone*", 
    "*Microsoft.ZuneMusic*",
    "*Microsoft.ZuneVideo*", 
    "*MicrosoftCorporationII.MicrosoftFamily*",
    "*Microsoft.OutlookForWindows*", 
    "*Microsoft.Todos*", 
    "*MSTeams*", 
    "*Copilot*"
)

foreach ($AppName in $AppsToRemove) {
    Get-AppxPackage -Name $AppName -AllUsers -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $AppName} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
}

# 3. Create the Local User Configuration Script
$UserScriptPath = "C:\Windows\UserConfig.ps1"
$UserScriptContent = @"
Set-ItemProperty -Path "HKCU:\Software\Policies\Microsoft\Windows\CloudContent" -Name "DisableSpotlightCollectionOnDesktop" -Value 1 -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings" -Name "EnabledState" -Value 0 -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers" -Name "BackgroundType" -Value 1 -Force
Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name "WallPaper" -Value ""
Set-ItemProperty -Path "HKCU:\Control Panel\Colors" -Name "Background" -Value "0 0 0"
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarAl" -Value 0 -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search" -Name "SearchboxTaskbarMode" -Value 0 -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "ShowTaskViewButton" -Value 0 -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "TaskbarMn" -Value 0 -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "AppsUseLightTheme" -Value 0 -Force
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize" -Name "SystemUsesLightTheme" -Value 0 -Force
New-Item -Path "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" -Value "" -Force
"@
$UserScriptContent | Out-File -FilePath $UserScriptPath -Encoding utf8 -Force

# 4. Set Active Setup (64-bit hive)
# This points Windows to run the local script that was just created
$ASPath = "SOFTWARE\Microsoft\Active Setup\Installed Components\MyCustomConfig"
$ASKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
$ASSubKey = $ASKey.CreateSubKey($ASPath)
$ASSubKey.SetValue("Version", "5")
$ASSubKey.SetValue("StubPath", "powershell.exe -ExecutionPolicy Bypass -File $UserScriptPath")
$ASSubKey.Close()
$ASKey.Close()

# 5. Forced Reboot
Start-Process "shutdown.exe" -ArgumentList "/r /t 5 /f" -Wait
exit 0
