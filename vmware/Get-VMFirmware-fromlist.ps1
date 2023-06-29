$date = Get-Date -UFormat %Y%m%d_%H%M

$vm_list = Get-Content "$PSScriptRoot\inputdata\20230309_2016-2019_WinVMs.txt"

$output = @()

# Retrieve firmware type for each virtual machine in the list
foreach ($vm_name in $vm_list) {
    $vm = Get-VM $vm_name
    $firmware_type = $vm.ExtensionData.Config.Firmware
    $output += [PSCustomObject]@{
        ServerName = $vm.Name
        FirmwareType = $firmware_type
    }
}
$csvFile = "$PSScriptRoot\outputdata\" + $date + "_WindowsServer2016-2019_firmware.csv"

$output | Export-Csv "$csvFile" -NoTypeInformation