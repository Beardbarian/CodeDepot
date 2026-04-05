# Adds "ALL APPLICATION PACKAGES" with Read permissions to HKLM\SYSTEM\CurrentControlSet\Control using SID

$registryPath = "HKLM\SYSTEM\CurrentControlSet\Control"
$principalSid = New-Object System.Security.Principal.SecurityIdentifier("S-1-15-2-1")

# Get the current ACL
$acl = Get-Acl -Path "Registry::$registryPath"

# Define the access rule
$rule = New-Object System.Security.AccessControl.RegistryAccessRule(
    $principalSid,
    [System.Security.AccessControl.RegistryRights]::ReadKey,
    [System.Security.AccessControl.InheritanceFlags]::ContainerInherit,
    [System.Security.AccessControl.PropagationFlags]::None,
    [System.Security.AccessControl.AccessControlType]::Allow
)

# Add the rule if it doesn't already exist
$exists = $false
foreach ($access in $acl.Access) {
    if ($access.IdentityReference -eq $principalSid -and
        ($access.RegistryRights -band [System.Security.AccessControl.RegistryRights]::ReadKey)) {
        $exists = $true
        break
    }
}

if (-not $exists) {
    $acl.SetAccessRule($rule)
    Set-Acl -Path "Registry::$registryPath" -AclObject $acl
    Write-Output "Added Read permission for 'ALL APPLICATION PACKAGES' to '$registryPath'."
} else {
    Write-Output "'ALL APPLICATION PACKAGES' already has Read permission on '$registryPath'."
}
