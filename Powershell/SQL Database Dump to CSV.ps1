$databaseName="Database"
$instanceName="Instance"
$baseExportPath="C:\Exports\"
$query = "SELECT name FROM sys.Tables"
$tableNames = Invoke-SqlCmd –ServerInstance $instanceName -Database $databaseName –Query $query

New-Item -Force $baseExportPath -type directory

foreach($dataRow in $tableNames)
{ 
    $exportFileName=$baseExportPath + "\\" + $dataRow.get_Item(0).ToString() + ".csv"
   $tableSpecificQuery="select * from " + $dataRow.get_Item(0).ToString()
   Invoke-SqlCmd –ServerInstance $instanceName -Database $databaseName –Query $tableSpecificQuery | Export-Csv -Path $exportFileName -NoTypeInformation
}
