$Creds = Get-Credential
Restart-Computer -ComputerName "PC Name" -Credential $Creds -Force
