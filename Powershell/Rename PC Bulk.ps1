[System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic') | Out-Null
$DomainAdminCredentials = Get-Credential -Message "Enter Domain Admin Credentials to rename computers"

Function Get-FileName($initialDirectory)
{   
 [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = "csv files (*.csv)| *.csv"
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
     }

$ComputerList = Get-FileName -initialDirectory $env:userprofile\documents
$computers = Import-CSV $ComputerList

$renamedcomputers = @() 
$unavailablecomputers = @()

Foreach ($computer in $computers){
$PingTest = Test-Connection -ComputerName $computer.CurrentName -Count 1 -quiet
If ($PingTest){
    Write-Host "Renaming $($computer.currentname) to $($computer.NewName)"
        Rename-Computer -ComputerName $Computer.CurrentName -NewName $Computer.NewName -DomainCredential $DomainAdminCredentials -Confirm:$false -Force
            $renamedcomputers += $computer.CurrentName
     
    }

Else{
        Write-Warning "Failed to connect to computer $($computer.currentname)"
            $unavailablecomputers += $computer.CurrentName

    }
}

$renamedcomputers | Out-File $env:userprofile\documents\renamedcomputers.txt
$unavailablecomputers | Out-File $env:userprofile\documents\unavailablecomputers.txt
$Error | Out-File $env:userprofile\documents\renamedcomputererrors.txt
