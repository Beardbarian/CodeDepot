Get-Printer | Where-Object { $_.Name -notmatch "PDF|txt" } | Remove-Printer -Force
