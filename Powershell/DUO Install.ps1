#DUO installer

write-host "- Install or Update DUO"
write-host "======================================"

#---parameter names passed to the installer------

$DuoIkey="Integration Key"
$DuoSkey="Secret Key"
$DuoAPI="API Hostname"
$AUTOPUSH="#0"
$FAILOPEN="#1"
$RDPONLY="#0"
$SMARTCARD="#0"
$ENABLEOFFLINE="#0"
#$varInstallArgs= "IKEY="@DuoIkey@" SKEY="@DuoSkey@" Host="@DuoHost@" AutoPush="@AutoPush@" FailOpen="@FailOpen@" RDPOnly="@RDPOnly@" SMARTCARD="@SmartCard@" ENABLEOFFLINE="@EnableOffline@" /qn"

#Checking Regkey Characters
$Characters = $DuoIkey.Length
$Characters2 = $DuoSkey.Length

if($Characters -gt 10 -and $Characters2 -gt 10){
#Creating a new directory for DUO installer and deleting old one
rmdir -Path "C:\DUO" -Recurse -Force -ErrorAction SilentlyContinue
rmdir -Path "C:\DUO_DRMM" -Recurse -Force -ErrorAction SilentlyContinue
New-item -path "C:\DUO_DRMM" -ItemType Directory

#---Installation-----------------------------------------------------------------
Function DownloadDuo
{
    Invoke-WebRequest -Uri "https://dl.duosecurity.com/DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip" -OutFile "C:\DUO_DRMM\DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip"
    Start-Sleep -Seconds 10
    $check1=Test-Path -Path "C:\DUO_DRMM\DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip"
    if ($check1 -eq "True")
        {Write-Host "File Downloaded"}
    else
        {start-bitstransfer -source 'https://dl.duosecurity.com/DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip' -Destination 'C:\DUO_DRMM\DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip'
        Start-Sleep -Seconds 120
        Test-Path -Path "C:\DUO_DRMM\DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip"
        Write-Host "File Downloaded"
        }

    
    Expand-Archive -LiteralPath 'C:\DUO_DRMM\DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip' -DestinationPath C:\DUO_DRMM\
    unblock-file -path "C:\DUO_DRMM\DuoWinLogon_MSIs_Policies_and_Documentation-latest.zip"

}
Function InstallDuo
{
    If((Get-CimInStance Win32_OperatingSystem).OSArchitecture.Trim() -eq "64-Bit")
{

        Write-Host "64-Bit OS detected...."
        Write-Host "Installing the 64-bit DUO installer"
        Write-Host "Installing Duo Authentication for Windows"
        DownloadDuo
        sleep 30
        msiexec /i "C:\DUO_DRMM\DuoWindowsLogon64.msi" IKEY="$DuoIKey" SKEY="$DuoSKey" Host="$DuoAPI" AutoPush="$AUTOPUSH" FailOpen="$FAILOPEN" RDPOnly="$RDPONLY" SMARTCARD="$SMARTCARD" ENABLEOFFLINE="$ENABLEOFFLINE" /qn
        sleep 30

}
else
{
        Write-Host "32-Bit OS detected...."
        Write-Host "Installing the 32-bit DUO installer"
        Write-Host "Installing Duo Authentication for Windows"
        DownloadDuo
        sleep 30
        msiexec /i "C:\DUO_DRMM\DuoWindowsLogon32.msi" IKEY="$DuoIKey" SKEY="$DuoSKey" Host="$DuoAPI" AutoPush="$AUTOPUSH" FailOpen="$FAILOPEN" RDPOnly="$RDPONLY" SMARTCARD="$SMARTCARD" ENABLEOFFLINE="$ENABLEOFFLINE" /qn
        sleep 30
}
}

Function Check_Software
{
    $software = "*Duo Authentication for Windows Logon*";
    $installed = (Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\* | Where { $_.DisplayName -like $software }) -ne $null



If(-Not $installed) {
    Write-Host "'Duo Authentication for Windows Logon' Failed to Install." -ForegroundColor Red
} else {
    Write-Host "'Duo Authentication for Windows Logon' is installed successfully." -ForegroundColor Green
}
}

InstallDuo
Check_Software

Get-ChildItem -Path C:\DUO_DRMM\ -File -Recurse| Remove-Item -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath C:\DUO_DRMM\ -Recurse -Force -ErrorAction SilentlyContinue
start-sleep -Seconds 10

write-host "- Exiting..."
}

Else {Write-Host "Reg key is less than 10 characters hence exiting the script without Installing Duo"
}


