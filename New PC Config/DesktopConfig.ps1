# 1. Create Detection Flag IMMEDIATELY (HKLM)
# We do this first so Intune knows the script at least started.
$registryPath = "HKLM:\SOFTWARE\MyCustomConfig"
if (-not (Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
}
New-ItemProperty -Path $registryPath -Name "ConfigApplied" -Value 1 -PropertyType DWord -Force

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
    Get-AppxPackage -Name $AppName -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
    Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $AppName} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
}

# 3. System-Wide Registry Tweaks
$WidgetPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh'
if (-not (Test-Path $WidgetPath)) { New-Item -Path $WidgetPath -Force | Out-Null }
Set-ItemProperty -Path $WidgetPath -Name AllowNewsAndInterests -Type DWord -Value 0

# 4. Active Setup for User-Specific UI Tweaks
$UserSettingsCmd = 'powershell.exe -ExecutionPolicy Bypass -Command "' +
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
'"'

# Register the Active Setup component in HKLM
$ActiveSetupPath = "HKLM:\SOFTWARE\Microsoft\Active Setup\Installed Components\MyCustomConfig"
if (-not (Test-Path $ActiveSetupPath)) { New-Item -Path $ActiveSetupPath -Force | Out-Null }
Set-ItemProperty -Path $ActiveSetupPath -Name "Version" -Value "1"
Set-ItemProperty -Path $ActiveSetupPath -Name "StubPath" -Value $UserSettingsCmd

# 5. Exit with Soft Reboot Code
exit 3010
