#Clears WMI configuration so it can be rebuilt. Useful for systems that aren't reporting their configuration.
#For after script testing: https://www.lisenet.com/2014/get-windows-system-information-via-wmi-command-line-wmic/

# Disable and stop winmgmt service
sc config winmgmt start= disabled
net stop winmgmt /y
# Clear repository
del C:\Windows\System32\wbem\repository\* /Q /F /S
# Enable and start the winmgmt service
sc config winmgmt start= auto
net start winmgmt
# Cleanup files
cd C:\Windows\System32\wbem\
for /f %s in ('dir /b *.mof') do mofcomp %s
for /f %s in ('dir /b en-us\*.mfl') do mofcomp en-us\%s
