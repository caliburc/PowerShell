
# Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter`n"
    }else {
    Write-Host "Connecting to vCenter`n"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}

$myvm = Read-Host "Enter VM Name (as shown in vCenter)"

$consolecons = (Get-VM -name $myvm).ExtensionData.QueryConnections()
$consoleconscount = $consolecons.Count
Write-Host "Found $consoleconscount Connections, Removing them..."

foreach ($con in $consolecons) {

    $connection = New-Object VMware.Vim.VirtualMachineConnection
    $connection.Label = $con.Label
    $connection.Client = $con.Client
    $connection.UserName = $con.UserName

    (Get-VM -name $myvm).ExtensionData.DropConnections(@($connection))
}
