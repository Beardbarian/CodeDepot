$Creds = Get-Credential
Restart-Computer -ComputerName PHB-CE1 -Credential $Creds -Force
