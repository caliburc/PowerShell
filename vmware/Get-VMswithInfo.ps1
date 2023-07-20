#using powercli, get all VMs that are powered on, their operating system, and IP address
#and export to a CSV file
#Get current date
$Date = Get-Date -Format "yyyyMMdd"

#Connect to vCenter
# Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter"
    }else {
    Write-Host "Connecting to vCenter"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}

#Get all VMs that are powered on
$vmList = Get-VM | Where-Object {$_.PowerState -eq "PoweredOn"} | Sort-Object -Property Name

#set the default values for the progress bar
$vmCount = $vmList.Count
$progress = 0
$vmNum = 1

#Create an empty array to store the results
$VMInfo = @()

#Loop through each VM and get the operating system and IP address
foreach ($vmName in $vmList) {
    $VM = Get-VM $vmName
    # Update Progress bar
    $progress = [Math]::Round($vmNum / $vmCount * 100)
    Write-Progress -Activity "Setting Values for $vmName" -Status "Processing VM $vmNum of $vmCount" -PercentComplete $progress

    $OS = $VM.ExtensionData.Summary.Config.GuestFullName
    # VMware is not always correct with what version of Windows, this checks the OS version from within the server itself.
    # Sometimes we might not have permission to check this or the machine is not on the domain, these catch's check for that and use the VMware OS value if that's the case
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
    $IP = $VM.Guest.IPAddress[0]
    #get the Tag Category named "Program_Name" from the VM and return the Tag Name for that Tag
    $Program = (Get-TagAssignment -Entity $VM -Category "Program_Name").Tag.Name

    #get the full VM folder path, including subfolders
    $VMFolder = $VM.Folder.FullPath

    #Create a custom object to store the results
    $VMInfoObj = New-Object -TypeName PSObject
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name VMName -Value $VM.Name
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name OS -Value $OS
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name IPAddress -Value $I
    $VMInfoObj | Add-Member -MemberType NoteProperty -Name Program -Value $Program


    #Add the custom object to the array
    $VMInfo += $VMInfoObj
    $vmNum++
}

#Export the results to a CSV file, with the filename being "${Date}_VMExport.csv" e.g. "20180101_VMExport.csv"  
$VMInfo | Export-Csv -Path "\\fshill\hedc\depot\tmp\$Date" + "_VMExport.csv" -NoTypeInformation

#Disconnect from vCenter
#Disconnect-VIServer -Server vcenter.domain.com -Confirm:$false

