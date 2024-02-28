Import-Module Posh-SSH

Write-Host "This script resets the last changed date for the root password" -BackgroundColor Yellow -ForegroundColor Black
Write-Host "You will be connected to vCenter first, pay attention to the prompts" -BackgroundColor Yellow -ForegroundColor Black
Write-Host "Enter your vCenter Creds first, then the ESXi credentials.`n" -BackgroundColor Yellow -ForegroundColor Black
Write-Host ">>>>> Press Enter to Continue <<<<<<"
Read-Host " "

Write-Host "------------- START vCenter Auth -------------`n"
    #Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter"
    }else {
    Write-Host "Connecting to vCenter"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}
Write-Host "`n------------- END vCenter Auth -------------`n"

# Define your list of ESXi hosts (use FQDN)
$myhosts = @("v2g04u22.vmn.infra.hedc","v2g05u22.vmn.infra.hedc")

# Provide SSH credentials
$esxicred = Get-Credential -Message "Input Credential for ESXi"

foreach($esxi in $myhosts){
    # Enable SSH service on Host
	  Get-VMHostService -VMHost $esxi | Where-Object {$_.Key -eq "TSM-SSH" } | Start-VMHostService -confirm:$false 
    $session = New-SSHSession -ComputerName $esxi -Credential $esxicred
    # Set the variable for the command to run the lastdate change command and ouput the results from the shadow file
    $setpwchangedate = '/usr/lib/vmware/auth/bin/chage root --lastday=$(($(date +%s) / 86400 - 2)) && newchangedate=$(date -d "@$(( $(cat /etc/shadow | grep "^root" | cut -d: -f3) * 86400 ))"); echo "PW Changed date set to: $newchangedate"'
    Invoke-SSHCommand -SessionId $session.SessionId  -Command $setpwchangedate -ShowStandardOutputStream -ShowErrorOutputStream
    Remove-SSHSession -SessionId $session.SessionId
}
