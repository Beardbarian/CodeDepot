:: Copy All
robocopy "source" "dest" /E /XO /B /COPYALL /MT:64 /R:0 /W:0 /LOG:C:\IT\Data.txt

:: Primary method with no log
robocopy "source" "dest" /E /XC /XN /XO 

:: Primary method with no log if destination is hidden
robocopy "source path" "destination path" /E /XC /XN /XO
attrib -h  -s  -a "destination path"

:: use /xj to exclude junction points
robocopy "source" "dest" /E /XJ /B /COPYALL /MT:64 /R:3 /W:5 /LOG:C:\IT\Data.txt

:: use /xd to exclude a path
robocopy "source" "dest" /E /XC /XN /XO /w:0 /r:0 /XD "excluded path" /LOG:C:\IT\Data.txt

:: Copy only data and timestamps
robocopy "source" "dest" /E /XO /B /COPY:DT /MT:64 /R:3 /W:5 /LOG:C:\IT\Data.txt
