forfiles /p "C:\what\ever" /s /m *.* /D -<number of days> /C "cmd /c del @path"

:: Example: older than 3 days
forfiles /p "C:\what\ever" /s /m *.* /D -3 /C "cmd /c del @path"
