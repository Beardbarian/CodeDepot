:: Remove all network printers
:: Better to use Powershell where possible
:: https://github.com/Beardbarian/Depo/blob/main/Powershell/Remove%20Printers.ps1

strComputer = "."  
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")  
  
Set colInstalledPrinters =  objWMIService.ExecQuery _  
    ("Select * from Win32_Printer Where Network = TRUE")  
  
For Each objPrinter in colInstalledPrinters  
    objPrinter.Delete_  
Next
