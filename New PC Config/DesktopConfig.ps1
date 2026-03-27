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

# 3. System-Wide Registry Tweaks (Widgets)
$WidgetPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh'
if (-not (Test-Path $WidgetPath)) { New-Item -Path $WidgetPath -Force | Out-Null }
Set-ItemProperty -Path $WidgetPath -Name AllowNewsAndInterests -Type DWord -Value 0

# 4. Active Setup for User-Specific UI Tweaks (Dark Mode, Black BG, Taskbar)
# Create a local 'UserConfig.ps1' that stays on the machine
$UserScriptContent = @"
    'Set-ItemProperty -Path \"HKCU:\Software\Policies\Microsoft\Windows\CloudContent\" -Name \"DisableSpotlightCollectionOnDesktop\" -Value 1 -Force -ErrorAction SilentlyContinue; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings\" -Name \"EnabledState\" -Value 0 -Force -ErrorAction SilentlyContinue; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers\" -Name \"BackgroundType\" -Value 1 -Force; ' +
    'Set-ItemProperty -Path \"HKCU:\Control Panel\Desktop\" -Name \"WallPaper\" -Value \"\"; ' +
    'Set-ItemProperty -Path \"HKCU:\Control Panel\Colors\" -Name \"Background\" -Value \"0 0 0\"; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" -Name \"TaskbarAl\" -Value 0 -Force; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\Search\" -Name \"SearchboxTaskbarMode\" -Value 0 -Force; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" -Name \"ShowTaskViewButton\" -Value 0 -Force; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\" -Name \"TaskbarMn\" -Value 0 -Force; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\" -Name \"AppsUseLightTheme\" -Value 0 -Force; ' +
    'Set-ItemProperty -Path \"HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\" -Name \"SystemUsesLightTheme\" -Value 0 -Force; ' +
    'New-Item -Path \"HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32\" -Value \"\" -Force' +
"@

$LocalUserScript = "C:\Windows\UserConfig.ps1"
$UserScriptContent | Out-File -FilePath $LocalUserScript -Force

# Register Active Setup in 64-bit hive
$ASPath = "SOFTWARE\Microsoft\Active Setup\Installed Components\MyCustomConfig"
$RegistryKey = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
$SubKey = $RegistryKey.CreateSubKey($ASPath)

# Date-based version so it always increments during testing
$SubKey.SetValue("Version", "2026032601") 
$SubKey.SetValue("StubPath", "powershell.exe -ExecutionPolicy Bypass -File $LocalUserScript")
$SubKey.Close()
$RegistryKey.Close()

# 5. Forced Hard Reboot
# Triggers a restart in 5 seconds to flush the E3 identity broker and trigger Active Setup
Start-Process "shutdown.exe" -ArgumentList "/r /t 5 /f" -Wait
exit 0
