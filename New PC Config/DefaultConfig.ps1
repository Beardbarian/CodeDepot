#Remove Apps
Get-AppxPackage -AllUsers *Clipchamp.Clipchamp* | Remove-AppxPackage
Get-AppxPackage -AllUsers *bing* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.GetStarted* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.Messaging* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.Microsoft3DViewer* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.MicrosoftOfficeHub* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.MicrosoftSolitaireCollection* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.MixedReality.Portal* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.News* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.OneConnect* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.People* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.PowerAutomateDesktop* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.SkypeApp* | Remove-AppxPackage
Get-AppxPackage -AllUsers *microsoft.windowscommunicationsapps* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.WindowsFeedbackHub* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.WindowsMaps* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.YourPhone* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.ZuneMusic* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.ZuneVideo* | Remove-AppxPackage
Get-AppxPackage -AllUsers *MicrosoftCorporationII.MicrosoftFamily* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.OutlookForWindows* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Microsoft.Todos* | Remove-AppxPackage
Get-AppxPackage -AllUsers *MSTeams* | Remove-AppxPackage
Get-AppxPackage -AllUsers *Copilot* | Remove-AppxPackage
#Disables Windows Spotlight, sets background to solid color black.
Set-ItemProperty -Path 'HKCU:\Software\Policies\Microsoft\Windows\CloudContent' -Name DisableSpotlightCollectionOnDesktop -type DWord -Value '1'
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\DesktopSpotlight\Settings' -Name EnabledState -type DWord -Value '0'
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Wallpapers' -Name BackgroundType -type DWord -Value '1'
Set-ItemProperty -Path 'HKCU:\Control Panel\Desktop' -Name WallPaper -Value ''
Set-ItemProperty -Path 'HKCU:\Control Panel\Colors' -Name Background -Value '0 0 0'
#Left align taskbar
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarAl -type DWord -Value '0'
#Set dark mode theme
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name AppsUseLightTheme -type DWord -Value '0'
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize' -Name SystemUsesLightTheme -type DWord -Value '0'
#Hide search
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Search' -Name SearchboxTaskbarMode -type DWord -Value '0'
#Hide taskview
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name ShowTaskViewButton -type DWord -Value '0'
#Hide chat
Set-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced' -Name TaskbarMn -type DWord -Value '0'
# Restart windows Explorer
Stop-Process -ProcessName explorer
