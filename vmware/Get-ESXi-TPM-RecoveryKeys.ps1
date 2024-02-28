$date = Get-Date -Format "yyyyMMdd"
$filename = "$date" + "_ESXi_TPM_Recovery_Keys.csv"
$filepath = "\\path\vmware"

# Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter"
    }else {
    Write-Host "Connecting to vCenter"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}

$VMHosts = Get-VMHost | Sort-Object
$VMHostKeys = @()
foreach ($VMHost in $VMHosts) {
    $esxcli = Get-EsxCli -VMHost $VMHost -V2
    try {
        $encryption = $esxcli.system.settings.encryption.get.Invoke()
        if ($encryption.Mode -eq "TPM")
        {
            $key = $esxcli.system.settings.encryption.recovery.list.Invoke()
            $hostKey = [pscustomobject]@{
                Host = $VMHost.Name
                Cluster = $VMhost.Parent.Name
                EncryptionMode = $encryption.Mode
                RequireExecutablesOnlyFromInstalledVIBs = $encryption.RequireExecutablesOnlyFromInstalledVIBs
                RequireSecureBoot = $encryption.RequireSecureBoot
                RecoveryID = $key.RecoveryID
                RecoveryKey = $key.Key
            }
            $VMHostKeys += $hostKey
        }
        else
        {
            $hostKey = [pscustomobject]@{
                Host = $VMHost.Name
                Cluster = $VMhost.Parent.Name
                EncryptionMode = $encryption.Mode
                RequireExecutablesOnlyFromInstalledVIBs = $encryption.RequireExecutablesOnlyFromInstalledVIBs
                RequireSecureBoot = $encryption.RequireSecureBoot
                RecoveryID = $null
                RecoveryKey = $null
            }
            $VMHostKeys += $hostKey
        }
    }
    catch {
        $VMHost.Name + $_
    }
}
$VMHostKeys | Export-Csv "$filepath\$filename"
