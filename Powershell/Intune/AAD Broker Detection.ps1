$registryPath = "HKLM\SYSTEM\CurrentControlSet\Control"
 
$principal = "APPLICATION PACKAGE AUTHORITY\ALL APPLICATION PACKAGES"

# Get the ACL for the registry key
$acl = Get-Acl -Path "Registry::$registryPath"

# Check for read permissions
$hasReadAccess = $false
foreach ($access in $acl.Access) {
    if ($access.IdentityReference -eq $principal -and
($access.RegistryRights -band [System.Security.AccessControl.RegistryRights]::ReadKey)) {
        $hasReadAccess = $true
        break
    }
}

if ($hasReadAccess) {
    Write-Output "$principal has read access to '$registryPath'."
    exit 0
} else {
    Write-Output "$principal does NOT have read access to '$registryPath'."
    exit 1
}
