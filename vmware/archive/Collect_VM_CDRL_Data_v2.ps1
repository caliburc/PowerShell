Remove-Variable * -Force -ErrorAction silentlycontinue

# Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter"
    }else {
    Write-Host "Connecting to vCenter"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}
# Define Customer VMs 
$projectName = Read-Host "Enter Project Name (Defined by VMware Folder Name)"

# Import the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Open the Template CDRL Workbook
$workbook = $excel.Workbooks.Open("\\fshill\hedc\depot\Infrastructure\compliance\CDRL\CDRL A007 Configuration Baseline Development - Template.xlsx")

# Select the worksheet
$worksheet = $workbook.Worksheets.Item(2)

#Get the VM's from the input projectName provided and sort them
$vmList = Get-VM -Location (Get-Folder -Name $projectName) | Sort

#set the default values for the progress bar
$vmCount = $vmList.Count
$progress = 0
$vmNum = 1

#select the starting row range to copy from for formatting
$startRow = 4
$itemID = 1
$startRange = $worksheet.Range("A4:AA4")

foreach ($vmName in $vmList) {
    # Update Progress bar
    $progress = [Math]::Round($vmNum / $vmCount * 100)
    Write-Progress -Activity "Setting Values for $vmName" -Status "Processing VM $vmNum of $vmCount" -PercentComplete $progress
    Write-Host "Setting Values for VM $vmName"
    $vm = Get-VM $vmName
    $os = $vm.ExtensionData.Summary.Config.GuestFullName
    $esxihost = Get-VMHost -VM $vmName
    # VMware is not always correct with what version of Windows, this checks the OS version from within the server itself
    if ($os -like "Microsoft Windows Server*") {
        $osCaption = (Get-WmiObject Win32_OperatingSystem -ComputerName $vmName).Caption
        $os = $osCaption
    }
    $vcpuCount = $vm.NumCpu
    $memoryCount = $vm.MemoryGB
    $storage = $vm.ProvisionedSpaceGB.ToString("N2")

    # Populate the cells with the virtual machine information
    $worksheet.Cells.Item($startRow, "D").Value2 = "x"
    $worksheet.Cells.Item($startRow, "G").Value2 = $itemID
    $worksheet.Cells.Item($startRow, "H").Value2 = "$vmName"
    $worksheet.Cells.Item($startRow, "I").Value2 = "1"
    $worksheet.Cells.Item($startRow, "J").Value2 = "VM"
    $worksheet.Cells.Item($startRow, "K").Value2 = "1"
    $worksheet.Cells.Item($startRow, "L").Value2 = "$os"
    $worksheet.Cells.Item($startRow, "M").Value2 = "$vcpuCount"
    $worksheet.Cells.Item($startRow, "N").Value2 = "$memoryCount"
    $worksheet.Cells.Item($startRow, "O").Value2 = "$storage"
    $worksheet.Cells.Item($startRow, "P").Value2 = "$esxihost"
    $worksheet.Cells.Item($startRow, "Q").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "R").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "S").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "T").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "U").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "V").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "W").Value2 = "PBL"

    # Copy The Formatting from the first row and paste it to this row, also autofit the columns
    $startRange.Copy() | Out-Null
    $worksheet.Range("A$($startRow):AA$($startRow)").PasteSpecial(-4122) | Out-Null
    $worksheet.Range("A$($startRow):AA$($startRow)").EntireColumn.AutoFit() | Out-Null

    # Move to the next row
    $startRow++
    # Itterate ID
    $itemID++
    # Increment Progress
    $vmNum++
}


$date = Get-Date -Format "yyyyMMdd"
$filename = "$date" + " $projectName" + " CDRL A007 Configuration Baseline.xlsx"
$filepath = "\\fshill\hedc\depot\Infrastructure\compliance\CDRL"
Write-Host "Saving file to $filepath\$filename"
$worksheet.SaveAs("$filePath\$filename")
$workbook.Close()
$excel.Quit()
sleep 8

Remove-Variable * -Force -ErrorAction silentlycontinue