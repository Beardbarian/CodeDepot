# 1. Provisioned App Removal
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
    # Remove from current user session
    Get-AppxPackage -Name $AppName -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
    # Remove from the Windows Image (Provisioned) so they don't come back
    Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $AppName} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
}

# 2. UI Customization (Left Align, Dark Mode, Solid Black)
$RegistryPaths = @(
    "HKCU:\Software\Policies\Microsoft\Windows\CloudContent",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
)

# Ensure paths exist before setting values
foreach ($path in $RegistryPaths) {
    if (-not (Test-Path $path)) { New-Item -Path $path -Force | Out-Null }
}

# Apply Styles
Set-ItemProperty -Path 'HKCU:\Software\Policies\Microsoft\Windows\CloudContent' -Name DisableSpotlightCollectionOnDesktop -Type DWord -Value 1
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings' -Name EnabledState -Type DWord -Value 0
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers' -Name BackgroundType -Type DWord -Value 1
Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallPaper -Value ""
Set-ItemProperty -Path 'HKCU:\Control Panel\Colors' -Name Background -Value "0 0 0"

# Taskbar and Search
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarAl -Type DWord -Value 0 # Left Align
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' -Name SearchboxTaskbarMode -Type DWord -Value 0 # Hide Search
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name ShowTaskViewButton -Type DWord -Value 0
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarMn -Type DWord -Value 0

# Dark Mode
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -Type DWord -Value 0
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name SystemUsesLightTheme -Type DWord -Value 0

# 3. System Tweaks (Widgets & Context Menu)
# Disable Widgets
$WidgetPath = 'HKLM:\SOFTWARE\Policies\Microsoft\Dsh'
if (-not (Test-Path $WidgetPath)) { New-Item -Path $WidgetPath -Force | Out-Null }
Set-ItemProperty -Path $WidgetPath -Name AllowNewsAndInterests -Type DWord -Value 0

# Restore Classic Context Menu (Win10 Style)
$ContextKey = "HKCU:\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32"
if (-not (Test-Path $ContextKey)) { New-Item -Path $ContextKey -Value "" -Force | Out-Null }

# 4. Refresh Explorer
Stop-Process -ProcessName explorer -Force
