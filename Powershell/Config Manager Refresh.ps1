#Application Deployment Evaluation Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000121}" /NOINTERACTIVE
#Discovery Data Collection Cycle	
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000003}" /NOINTERACTIVE
#Hardware Inventory Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000001}" /NOINTERACTIVE
#Machine Policy Retrieval Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000021}" /NOINTERACTIVE
#Software Metering Usage Report Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000031}" /NOINTERACTIVE
#Software Updates Assignments Evaluation Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000108}" /NOINTERACTIVE
#Software Update Scan Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000113}" /NOINTERACTIVE
#User Policy Retrieval Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000026}" /NOINTERACTIVE
#User Policy Evaluation Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000027}" /NOINTERACTIVE
#Windows Installers Source List Update Cycle
WMIC /namespace:\\root\ccm path sms_client CALL TriggerSchedule "{00000000-0000-0000-0000-000000000032}" /NOINTERACTIVE


#2nd attempt WIP
-----------------------------------------------
$ComputerName = $env:COMPUTERNAME # Pulls PC Name, can replace with the target computer's name to run remotely

$ActionGUIDs = @(
    "{00000000-0000-0000-0000-000000000121}", # Application Deployment Evaluation Cycle
    "{00000000-0000-0000-0000-000000000003}", # Discovery Data Collection Cycle
    "{00000000-0000-0000-0000-000000000001}", # Hardware Inventory Cycle
    "{00000000-0000-0000-0000-000000000022}", # Machine Policy Retrieval & Evaluation Cycle
    "{00000000-0000-0000-0000-000000000031}", # Software Metering Usage Report Cycle
    "{00000000-0000-0000-0000-000000000108}", # Software Updates Deployment Evaluation Cycle
    "{00000000-0000-0000-0000-000000000113}", # Software Updates Scan Cycle
    "{00000000-0000-0000-0000-000000000021}", # User Policy Retrieval & Evaluation Cycle
    "{00000000-0000-0000-0000-000000000032}"  # Windows Installer Source List Update Cycle
)

foreach ($guid in $ActionGUIDs) {
    Write-Host "Triggering action with GUID: $guid on $ComputerName"
    Invoke-WmiMethod -Namespace "root\ccm" -Class "SMS_Client" -Name "TriggerSchedule" -ArgumentList $guid -ComputerName $ComputerName
}
