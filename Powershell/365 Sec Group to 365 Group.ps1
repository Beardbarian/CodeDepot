#Type name of security group
$Group1 = ""
#enter email alias without spaces or special car 
$mailname = ""
$group1ObjectID = Get-AzureADGroup -Filter "Displayname eq '$group1'" | Select objectid -ExpandProperty ObjectID 
#Create new Office 365 group
New-AzureADMSGroup -DisplayName $Group1 -MailNickname $mailname -GroupTypes "Unified" -MailEnabled $false -SecurityEnabled $True 
$group2 = Get-AzureADMSGroup -Filter "MailNickname eq '$mailname'" 
$membersGroup1 = Get-AzureADGroupMember -ObjectId $group1ObjectID -All $true
foreach($member in $membersGroup1)
{
    $currentuser = Get-AzureADUser -ObjectId $member.ObjectId | select objectid
    Add-AzureADGroupMember -ObjectId $group2.ID -RefObjectId $currentuser.objectid
}
#press enter
Get-AzureADGroupMember -ObjectId $group2.ID -All $true
