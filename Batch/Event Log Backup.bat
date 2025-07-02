Cd "C:\IT\EventLogs"
psloglist.exe -g "C:\IT\EventLogs\Security.Evt" Security
ren Security.evt "SecurityNEW_%date:~4,2%%date:~7,2%%date:~12,2%_%time:~0,2%%time:~3,2%%time:~6,2%.evt"

Cd "C:\IT\EventLogs"
forfiles /d -14 /m *.evt /c "cmd /c del @fname.evt"

get-eventlog -logname Security | ForEach { Clear-EventLog $_.Log }

for /F “tokens=Security” %1 in ('wevtutil.exe el') DO wevtutil.exe cl “%1”

Clear-Eventlog -Log Security
