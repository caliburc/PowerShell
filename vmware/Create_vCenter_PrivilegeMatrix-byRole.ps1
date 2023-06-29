$date = Get-Date -UFormat %Y%m%d_%H%M
$outputCSVPath = $PSScriptRoot + "\outputdata\" + $date + "_vCenter_Role_Privilege_Matrix.csv"
$privilegeIDs = Get-VIPrivilege

# Get all roles using Get-VIRole

$roles = Get-VIRole | Where-Object { $_.Name -like 'HEDC*' } | Sort-Object -Property Name
Write-Host "Gathering Privilges for the following roles:"
    foreach ($role in $roles) {
        Write-Host $role
        }

# Create an empty hashtable to store the data
$rolePermissions = @{}

# Iterate through each role
foreach ($role in $roles) {
    $roleName = $role.Name
    $rolePermissions[$roleName] = @{}

    # Iterate through each privilege ID
    foreach ($privilegeID in $privilegeIDs) {
        $hasPrivilege = Get-VIPrivilege -Role $role -Id $privilegeID.id -ErrorAction SilentlyContinue

        # Store the true/false value for the privilege in the rolePermissions hashtable
        $rolePermissions[$roleName][$privilegeID] = $hasPrivilege
    }
}

# Create an array to hold the output data
$outputData = @()

# Iterate through each privilege ID
foreach ($privilegeID in $privilegeIDs) {
    # Create a new object for each permission
    $permissionObject = [PSCustomObject]@{
        'Privilege Parent Group' = $privilegeID.parentgroup
        'Privilege Parent GroupID' = $privilegeID.parentgroupID
        'Privilege Name' = $privilegeID.Name
        'Privilege ID' = $privilegeID.id
        'Privilege Description' = $privilegeID.description
    }

    # Iterate through each role
    foreach ($role in $roles) {
        $roleName = $role.Name
        $hasPrivilege = $rolePermissions[$roleName][$privilegeID]

        # Add a new column for each role and store the true/false value
        #$permissionObject | Add-Member -MemberType NoteProperty -Name $roleName -Value $hasPrivilege
        if ($hasPrivilege) {
            $permissionObject | Add-Member -MemberType NoteProperty -Name $roleName -Value "True"
        } else {
            $permissionObject | Add-Member -MemberType NoteProperty -Name $roleName -Value "False"
        }
    }

    # Add the permission object to the output array
    $outputData += $permissionObject
}

# Export the output data to a CSV file
$outputData | Export-Csv -Path $outputCSVPath -NoTypeInformation

Write-Host "Exported CSV to $outputCSVPath"