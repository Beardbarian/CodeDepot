:: Clears WMI configuration so it can be rebuilt. Useful for systems that aren't reporting their configuration.
:: For after script testing: https://www.lisenet.com/2014/get-windows-system-information-via-wmi-command-line-wmic/

write-Host "Disabling and stoping winmgmt service" -ForegroundColor Green  
sc config winmgmt start= disabled
net stop winmgmt /y
write-Host "Clearing repository" -ForegroundColor Green  
del C:\Windows\System32\wbem\repository\* /Q /F /S
write-Host "Enabling and starting the winmgmt service" -ForegroundColor Green  
sc config winmgmt start= auto
net start winmgmt
write-Host "Cleaning up files" -ForegroundColor Green  
cd C:\Windows\System32\wbem\
for /f %s in ('dir /b *.mof') do mofcomp %s
for /f %s in ('dir /b en-us\*.mfl') do mofcomp en-us\%s
