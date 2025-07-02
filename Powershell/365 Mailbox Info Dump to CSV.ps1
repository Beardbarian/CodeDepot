#Shared Mailboxes
Get-Mailbox -RecipientTypeDetails SharedMailbox -ResultSize Unlimited | Get-MailboxPermission | Select-Object Identity, User, AccessRights | Where-Object { $_.User -like '*@*' } | Export-Csv -Path C:\IT\sharedmailboxes.csv -NoTypeInformation

#User Mailboxes
Get-Mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Get-MailboxPermission | Select-Object Identity, User, AccessRights | Where-Object { $_.User -like '*@*' } | Export-Csv -Path C:\IT\usermailboxes.csv -NoTypeInformation

#All
Get-Mailbox -ResultSize Unlimited | Get-MailboxPermission | Select-Object Identity, User, AccessRights | Where-Object { $_.User -like '*@*' } | Export-Csv -Path C:\IT\allmailboxes.csv -NoTypeInformation

#Recipient Types:
#DiscoveryMailbox
#EquipmentMailbox
#GroupMailbox (Exchange 2013 or later and cloud)
#LegacyMailbox
#LinkedMailbox
#LinkedRoomMailbox (Exchange 2013 or later and cloud)
#RoomMailbox
#SchedulingMailbox (Exchange 2016 or later and cloud)
#SharedMailbox
#TeamMailbox (Exchange 2013 or later and cloud)
#UserMailbox
