Write-Host -ForegroundColor Cyan "This script will connect to vCenter and collect the Name, OS, IP Address, and Program Name for all VMs that are powered on"
Read-Host "If you want to continue, press Enter. If you want to cancel, press Ctrl+C"

# Get current date
$Date = Get-Date -Format "yyyyMMdd"

# Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter"
    }else {
    Write-Host "Connecting to vCenter"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}

# Check if user is a Windows Admin
$username = whoami
if ($username -like ".*adf" -or $username -like ".*adm") {
    $windowsadmin = $true
    Write-Host "User is a Windows Admin"
    }else {
    $windowsadmin = $false
    Write-Host "User is not a Windows Admin"
}

# Get all VMs that are powered on
$vmList = Get-VM | Where-Object {$_.PowerState -eq "PoweredOn"} | Sort-Object -Property Name

# Set the default values for the progress bar
$vmCount = $vmList.Count
$progress = 0
$vmNum = 1

# Create an empty array to store the results
$VMInfo = @()

# Loop through each VM and get the operating system and IP address
foreach ($vmName in $vmList) {
    $VM = Get-VM $vmName
    # Update Progress bar
    $progress = [Math]::Round($vmNum / $vmCount * 100)
    Write-Progress -Activity "Setting Values for $vmName" -Status "Processing VM $vmNum of $vmCount" -PercentComplete $progress

    $OS = $VM.ExtensionData.Summary.Config.GuestFullName
    # VMware is not always correct with what version of Windows, this checks the OS version from within the server itself.
    # Sometimes we might not have permission to check this or the machine is not on the domain, these catch's check for that and use the VMware OS value if that's the case
    # we check if you're a windows admin before doing this, to avoid trying every single server and just failing
    if ($windowsadmin){
        if ($OS -like "Microsoft Windows Server*") {
            try {
            $osCaption = (Get-WmiObject Win32_OperatingSystem -ComputerName $vmName -ErrorAction Stop).Caption
            $OS = $osCaption
            }
            catch [System.Runtime.InteropServices.COMException] {
             if ($_.Exception.Message -like "*The RPC server is unavailable*") {
                    Write-Warning "RPC Server is unavailable when getting OS info for $vmName - Using VMware OS information instead"
                }
                else {
                    Write-Warning "Error geting OS infromation for $vmName : $($_.Exception.Message) - Using VMware OS information instead"
                }
            }
            catch {
                Write-Warning "Error getting OS information on VM $vmName : $($_.Exception.Message) - Using VMware OS information instead"
            }
        }
    }

    # Get all the IP's of a VM and join them into a single string separated by a semicolon
    $IP = $VM.Guest.IPAddress -join '; '

    # Get the Tag Category named "Program_Name" from the VM and return the Tag Name for that Tag
    $Program = (Get-TagAssignment -Entity $VM -Category "Program_Name").Tag.Name

    $VMInfoObj = New-Object -TypeName PSObject
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name VMName -Value $VM.Name
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name OS -Value $OS
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $IP
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name Program -Value $Program

    $VMInfo += $VMInfoObj
    $vmNum++
}

$csvFile = "\\fshill\hedc\depot\tmp\" + $Date + "_VMExport.csv"
$VMInfo | Export-Csv $csvFile -NoTypeInformation
Write-Host -ForegroundColor Cyan "Saving CSV file to " -NoNewline
Write-Host -ForegroundColor Yellow "$csvFile"
Read-Host "Press Enter to continue"


