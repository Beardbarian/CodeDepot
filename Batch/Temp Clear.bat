$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace(0xA)
$WinTemp = "c:\Windows\Temp\*"

:: Remove Temp Files  
write-Host "Removing Temp" -ForegroundColor Green  
Set-Location “C:\Windows\Temp”  
Remove-Item * -Recurse -Force -ErrorAction SilentlyContinue  
Set-Location “C:\Windows\Prefetch”  
Remove-Item * -Recurse -Force -ErrorAction SilentlyContinue  
Set-Location “C:\Documents and Settings”  
Remove-Item “.\*\Local Settings\temp\*” -Recurse -Force -ErrorAction SilentlyContinue  
Set-Location “C:\Users”  
Remove-Item “.\*\Appdata\Local\Temp\*” -Recurse -Force -ErrorAction SilentlyContinue  

:: Running Disk Clean up Tool  
write-Host "Running the Windows Disk Clean up Tool" -ForegroundColor White  
cleanmgr /sagerun:1 | out-Null  
$([char]7)  
Sleep 3  
write-Host "Cleanup task complete!" -ForegroundColor Yellow  
Sleep 3  
