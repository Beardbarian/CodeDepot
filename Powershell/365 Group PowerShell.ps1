#SMTP=Primary - smtp=Secondary

#365 Group
Set-UnifiedGroup "365 Group Name" -emailaddresses @{Add='smtp:#email1@email.com','smtp:#email2@email.com'}

#Mailbox
`Set-Mailbox "Finance" -EmailAddresses @{Add='SMTP:Finance@domain.org','smtp:user@domain.org','smtp:finance@domain.onmicrosoft.com'}`

#Set GUID
`Get-Mailbox user | fl ExchangeGuid`
`Set-RemoteMailbox "Mailbox Name" -ExchangeGuid GUID`

#Set forwarding
`Enable-RemoteMailbox "Mailbox Name" -RemoteRoutingAddress email@domain.mail.onmicrosoft.com`

#Set Username
`Set-Mailbox -identity Finance -WindowsEmailAddress Finance@domain.org`
