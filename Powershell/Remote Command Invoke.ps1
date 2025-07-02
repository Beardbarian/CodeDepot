Invoke-Command -ComputerName "Name" -scriptblock {cmd command}

#Reboot Example
Invoke-Command -ComputerName My-PC -scriptblock {shutdown /r /t 0}
