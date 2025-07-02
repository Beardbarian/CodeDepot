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

#365 Groups - Needs Tested
Get-UnifiedGroup -ResultSize Unlimited | Select-Object Identity, User, AccessRights | Where-Object { $_.User -like '*@*' } | Export-Csv -Path C:\IT\365groups.csv -NoTypeInformation

#Distro Group - Needs Tested
Get-DistributionGroup -ResultSize Unlimited | Select-Object Identity, User, AccessRights | Where-Object { $_.User -like '*@*' } | Export-Csv -Path C:\IT\365groups.csv -NoTypeInformation

#Dynamic Distro Group - Needs Tested
Get-DynamicDistributionGroup -ResultSize Unlimited | Select-Object Identity, User, AccessRights | Where-Object { $_.User -like '*@*' } | Export-Csv -Path C:\IT\365groups.csv -NoTypeInformation

#Alternate Method for 365 and Distro Groups and dumps csv to user's desktop
$Groups = Get-UnifiedGroup -ResultSize Unlimited

$Groups | ForEach-Object {
$group = $_
Get-UnifiedGroupLinks -Identity $group.Name -LinkType Members -ResultSize Unlimited | ForEach-Object {
      New-Object -TypeName PSObject -Property @{
       Group = $group.DisplayName
       Member = $_.Name
       EmailAddress = $_.PrimarySMTPAddress
       RecipientType= $_.RecipientType
}}} | Export-CSV "$env:USERPROFILE\Desktop\Office365GroupMembers.csv" -NoTypeInformation -Encoding UTF8
