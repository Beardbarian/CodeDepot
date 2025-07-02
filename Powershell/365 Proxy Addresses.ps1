Set-Mailbox user@domain.com -EmailAddresses @{Add='smtp:user@domain.globaldrilsup.com','smtp:user@domain.internal'}
Get-Mailbox user@domain.com | Format-List EmailAddresses
