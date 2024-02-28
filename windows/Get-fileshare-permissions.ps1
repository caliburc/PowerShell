$csvPath = "C:\Temp\sa07-share-permissions"  # Replace with your desired file path
$csvContent = "ServerName#DiskShareName#IdentityReference#FileSystemRights"

# Iterate over server names from krsm-fs-sa0701v to krsm-fs-sa0720v
for ($i = 1; $i -le 20; $i++) {
    $serverNumber = $i.ToString("D2")  # Format the number as 2-digit
    $serverName = "krsm-fs-sa07${serverNumber}v"  # Construct the server name
    
    $sharedResources = net view $serverName
    # Filter the list to get only lines containing "Disk" and extract share names
    $diskShares = $sharedResources | Where-Object { $_ -match "Disk\s+" } | ForEach-Object {
        $lineParts = $_ -split '\s+'
        $lineParts[0]
    }

    # Iterate over each disk share
    foreach ($diskShare in $diskShares) {
        $diskAcl = Get-Acl "\\$serverName\$diskShare"
        # Iterate over each access rule in the ACL and add it to our csv
        foreach ($accessRule in $diskAcl.Access) {
            $identityReference = $accessRule.IdentityReference
            $fileSystemRights = $accessRule.FileSystemRights
            $csvLine = "$serverName`#$diskShare`#$identityReference`#$fileSystemRights"
            $csvContent += "`n$csvLine" 
        }
    }
}

$csvContent | Out-File -FilePath $csvPath -Encoding utf8
Write-Host "CSV export complete."
