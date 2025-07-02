$Begin = Get-Date -Date '1/17/2019 08:00:00' 
$End = Get-Date -Date '1/17/2019 17:00:00' 
Get-EventLog -LogName System -EntryType Error -After $Begin -Before $End
