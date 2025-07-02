robocopy C:\Users\%username% C:\Users\%username%.old /E /Z /ZB /eta /np /v /XC /XN /xjd /xjf /XO /R:3 /W:0 /XD "C:\Users\%username%\Appdata\"

#Alternatively, login as local or domain admin, rename user folder to "user.old" and then delete profile. Use the following script to migrate back data.
#attrib is there for instances where the user folder becomes hidden due to robocopy
takown /f C:\%username%.old /r /d y
robocopy C:\Users\%username%.old C:\Users\%username% /E /Z /ZB /eta /np /v /XC /XN /xjd /xjf /XO /R:3 /W:0 /XD "C:\Users\%username%.old\Appdata\"
attrib -h  -s  -a C:\Users\%username%
pause
