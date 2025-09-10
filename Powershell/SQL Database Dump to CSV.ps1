# Define variables
$databaseName = "Database"
$instanceName = "Instance"
$baseExportPath = "C:\Exports"

# Ensure the export directory exists
if (-not (Test-Path -Path $baseExportPath -PathType Container)) {
    Write-Host "Creating export directory: $baseExportPath"
    New-Item -Path $baseExportPath -Type Directory -Force | Out-Null
}

try {
    # Get a list of all table names from the specified database
    Write-Host "Retrieving table names from database '$databaseName' on instance '$instanceName'..."
    $tableNames = Invoke-SqlCmd -ServerInstance $instanceName -Database $databaseName -Query "SELECT name FROM sys.tables"

    # Export each table to a CSV file
    foreach ($table in $tableNames) {
        $tableName = $table.name
        $exportFilePath = Join-Path -Path $baseExportPath -ChildPath "$tableName.csv"
        
        Write-Host "Exporting table '$tableName' to '$exportFilePath'..."
        $query = "SELECT * FROM [$tableName]"
        
        Invoke-SqlCmd -ServerInstance $instanceName -Database $databaseName -Query $query | Export-Csv -Path $exportFilePath -NoTypeInformation
    }

    Write-Host "All tables have been successfully exported."
}
catch {
    Write-Error "An error occurred: $($_.Exception.Message)"
}
