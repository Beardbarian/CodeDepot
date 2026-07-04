# 1. Clear the read-only safety lock on the migrated pool
Get-StoragePool | Where-Object {$_.IsPrimordial -eq $False} | Set-StoragePool -IsReadOnly $False

# 2. Force the virtual volumes to mount automatically
Get-VirtualDisk | Set-VirtualDisk -IsManualAttach $False
