# Connect to vcenter
# Change cluster and adjust "where-object" filter to find what you're looking fo (ex. the name of the VM's, current code filters for test and dev machines)
# Will print out the top 25 VM's with the highest Average IO and the the datastore the VM is on.
<#
Get-Cluster G2012-HEDC-WIN-2012-ONLY | Get-VM | Where-Object {$_.Name -like "*tv*" -or $_.Name -like "*dv*"} | Select-Object Name, 
    @{N="AverageIO"; 
      E={[math]::Round(($_ | Get-Stat -Stat disk.usage.average -Start (Get-Date).adddays(-1) | Measure-Object -Average -Property Value).Average, 2)}},
    @{N="Datastore"; 
      E={$_.ExtensionData.Config.DatastoreUrl.Name -split ' ' -join "`n"}} | `
    Sort-Object -Property AverageIO -Descending | Select-Object Name, AverageIO, Datastore -First 25
    #>

    #Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter"
    }else {
    Write-Host "Connecting to vCenter"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}

$cluster = "G00-HEDC-ADMIN-891"
$VMInfo = Get-Cluster $cluster | Get-VM <#| Where-Object {$_.Name -like "*tv*" -or $_.Name -like "*dv*"}#>| Select-Object Name, 
    @{N="AverageIO"; 
      E={[math]::Round(($_ | Get-Stat -Stat disk.usage.average -Start (Get-Date).adddays(-1) | Measure-Object -Average -Property Value).Average, 2)}},
    @{N="Datastore"; 
      E={$_.ExtensionData.Config.DatastoreUrl.Name -split ' ' -join "`n"}} | `
    Sort-Object -Property AverageIO -Descending | Select-Object Name, AverageIO, Datastore -First 25

$VMInfo | Format-Table -AutoSize
#$VMInfo | Out-GridView -Title "Top 25 VMs by Average IO on Cluster:$cluster"

