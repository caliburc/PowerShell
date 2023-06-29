# Connect to vCenter
Connect-VIServer -Server apvcsa1p.idm.hedc.mil

$filedate = Get-Date -UFormat %Y%m%d_%H%M

# Get all powered on Windows VMs
#$VMs = Get-VM | Where {$_.PowerState -eq "PoweredOn" -and $_.Guest.OSFullName -match "Windows"}

# Get all powered on VMs
$VMs = Get-VM | Where {$_.PowerState -eq "PoweredOn"}

# Initialize an array to store the snapshot data
$SnapshotData = @()

# Loop through each VM
foreach ($VM in $VMs) {
    # Get all snapshots for the VM
    $Snapshots = Get-Snapshot -VM $VM

    # Loop through each snapshot
    foreach ($Snapshot in $Snapshots) {
        # Check if the snapshot is older than 2 days
        if (((Get-Date) - $Snapshot.Created).TotalDays -gt 2) {
            # Add the snapshot data to the array
            $SnapshotData += New-Object PSObject -Property @{
                VM = $VM.Name
                SnapshotName = $Snapshot.Name
                Description = $snapshot.Description
                Created = $Snapshot.Created
                SizeMB = $Snapshot.SizeMB
            }
        }
    }
}

# Export the snapshot data to a CSV file
$SnapshotData | Export-Csv -Path "\\fshill\hedc\depot\Infrastructure\Technical\scripts\vmware\outputdata\$($filedate)_VMs_Snapshots_older2days.csv" -NoTypeInformation
