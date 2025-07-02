#For a 365 Group
Set-UnifiedGroup -Identity "Legal Department" -HiddenFromAddressListsEnabled $true -HiddenFromExchangeClientsEnabled

#For all disabled accounts
$mailboxes = get-user | where {$_.UserAccountControl -like '*AccountDisabled*' -and $_.RecipientType -eq 'UserMailbox' } | get-mailbox | where {$_.HiddenFromAddressListsEnabled -eq $false}
foreach ($mailbox in $mailboxes) { Set-Mailbox -HiddenFromAddressListsEnabled $true -Identity $mailbox }

#For a single mailbox
Set-Mailbox -Identity "UserMailbox" -HiddenFromAddressListsEnabled $true
