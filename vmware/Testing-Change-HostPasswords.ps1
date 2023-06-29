<#
Synopsis
    Creates a new random password and changes the Root user password on the ESXi Host

  #>

$vcenter = "apvcsa1p.idm.hedc.mil"

#This section generates a random password between 14 and 18 characters, with a minium of 6 non alphanumeric values (aka symbols)
Add-Type -AssemblyName 'System.Web'
$minLength = 14 ## characters
$maxLength = 18 ## characters
$length = Get-Random -Minimum $minLength -Maximum $maxLength
$nonAlphaChars = 6
$newPassword = [System.Web.Security.Membership]::GeneratePassword($length, $nonAlphaChars)

#Get the current password from the keypass
$kpscriptget = \\fshill\hedc\depot\projects\Keepass\Infrastructure\KPScript.exe -c:GetEntryString "\\fshill\hedc\depot\projects\Keepass\Infrastructure\HEDCInfraAutomation.kdbx" -keyfile:"\\fshill\hedc\depot\projects\Keepass\Infrastructure\HEDCInfraAutomation.keyx" -refx-Group:VMware -ref-Title:"ESXi Root" -Field:Password
#the command above creates some extra junk, we just want the first line from the output, grab it here
$currentpass = $kpscriptget.Split([Environment]::NewLine) | Select -First 1

# Define a log file
$DateString = Get-Date -UFormat %Y%m%d_%H%M
$LogFile = "$DateSTring.Change-HostPasswords.csv"
# Rename the old log file, if it exists
if(Test-Path $LogFile) {
	$DateString = Get-Date((Get-Item $LogFile).LastWriteTIme) -UFormat %Y%m%d_%H%M
	Move-Item $LogFile "$LogFile.old.csv" -Force -Confirm:$false
}
# Add some CSV headers to the log file
Add-Content $Logfile "Date,Location,Host,Result"

# Hide the warnings for certificates (or better, install valid ones!)
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false | Out-Null


$newPasswordSec = ConvertTo-SecureString -String $newPassword -AsPlainText -Force
$currentPasswordSec = ConvertTo-SecureString -String $currentpass -AsPlainText -Force

# Create credential objects using the supplied passwords
$RootCredential = New-Object System.Management.Automation.PSCredential("root",$currentPasswordSec)
$NewRootCredential = New-Object System.Management.Automation.PSCredential("root",$newPasswordSec)


# Connect to the vCenter server
Connect-VIServer $vCenter | Out-Null

# Create an object for the root account with the new pasword
$RootAccount = New-Object VMware.Vim.HostPosixAccountSpec
$RootAccount.id = "root"
$RootAccount.password = $newPassword
$RootAccount.shellAccess = "/bin/bash"

$VMHosts = Get-VMHost "v1c17u08.vmn.infra.hedc"
# Get the hosts from the Location and for each host

$VMHosts | % {
	# Disconnect any connected sessions - prevents errors getting multiple ServiceInstances
	$global:DefaultVIServers | Disconnect-VIServer -Confirm:$false
	Write-Debug ($_.Name + " - attempting to connect")
	# Create a direct connection to the host
	$VIServer = Connect-VIServer $_.Name -Credential $RootCredential -ErrorAction SilentlyContinue
	# If it's connected
	if($VIServer.IsConnected -eq $True) {
		Write-Debug ($_.Name + " - connected")
		$VMHost = $_
		# Attempt to update the Root user object using the account object we created before
		# Catch any errors in a try/catch block to log any failures.
		try {
			$ServiceInstance = Get-View ServiceInstance
			$AccountManager = Get-View -Id $ServiceInstance.content.accountManager 
			$AccountManager.UpdateUser($RootAccount)
			Write-Debug ($VMHost.Name + " - password changed")
			Add-Content $Logfile ((get-date -Format "dd/MM/yy HH:mm")+","+$VMHost.Parent+","+$VMHost.Name+",Success")
            $success = true
		}
		catch {
			Write-Debug ($VMHost.Name + " - password change failed")
			Write-Debug $_
			Add-Content $Logfile ((get-date -Format "dd/MM/yy HH:mm")+","+$VMHost.Parent+","+$VMHost.Name+",Failed (Password Change)")
		}
		# Disconnect from the server
		Disconnect-VIServer -Server $VMHost.Name -Confirm:$false -ErrorAction SilentlyContinue
		Write-Debug ($VMHost.Name + " - disconnected")
	} else {
		# Log any connection failures
		Write-Debug ($_.Name+" - unable to connect")
		Add-Content $Logfile ((get-date -Format "dd/MM/yy HH:mm")+","+$_.Parent+","+$_.Name+",Failed (Connection)")
	}
}

if ($success -eq $true) {
    \\fshill\hedc\depot\projects\Keepass\Infrastructure\KPScript.exe -c:EditEntry "\\fshill\hedc\depot\projects\Keepass\Infrastructure\HEDCInfraAutomation.kdbx" -keyfile:"\\fshill\hedc\depot\projects\Keepass\Infrastructure\HEDCInfraAutomation.keyx" -refx-Group:VMware -ref-Title:"ESXi Root" -UserName:root -set-Password:$newPassword
}

Remove-Variable * -Force -ErrorAction 'SilentlyContinue'