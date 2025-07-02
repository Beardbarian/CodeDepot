#remove
Remove-MailboxPermission -Identity user1@domain.com -User user2@domain.com -AccessRights FullAccess

#add
Add-MailboxPermission -Identity user1@domain.com -User user2@domain.com -AccessRights FullAccess -AutoMapping $false
