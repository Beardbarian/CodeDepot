#!ps
$params = @{
    Name                 = "" # CHANGE ME: Friendly name of the VPN
    ServerAddress        = "" # CHANGE ME: MX IP or hostname
    DnsSuffix            = "" # CHANGE ME: AD domain name
    L2tpPsk              = "" # CHANGE ME: Pre-shared Key
    TunnelType           = "L2tp"     # Required by Meraki
    AuthenticationMethod = "Pap"      # Required by Meraki
    EncryptionLevel      = "Optional" # This is required for PAP auth type.
    AllUserConnection    =  $true     # Remove the '#' before this line if you want it added for ALL user accounts on the computer.
    Force                =  $true     # Ignore errors
}

Remove-VpnConnection $params.Name -Force -AllUserConnection -ErrorAction Ignore # Remove any profile that uses the same name
Remove-VpnConnection $params.Name -Force -ErrorAction Ignore                    # Remove any profile that uses the same name

Add-VpnConnection @params # Add the above VPN profile

##Below this line adds a shortcut to RASPhone for the above profile.
##You do not need to change anything below here. It pulls the "Name" parameter from above
##and uses it to create the shortcut as well as call the right profile in RASPhone.
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("C:\users\public\Desktop\{0}.lnk" -f $params.Name)
$Shortcut.TargetPath = "C:\Windows\System32\rasphone.exe"
$Shortcut.Arguments = '-d "{0}"' -f $params.Name
$Shortcut.IconLocation = "%SystemRoot%\System32\SHELL32.dll, 13"
$Shortcut.Save()
