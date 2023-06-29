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

# Open the workbook
$workbook = $excel.Workbooks.Open("\\fshill\hedc\depot\Infrastructure\compliance\CDRL\CDRL A007 Configuration Baseline Development - Template2.xlsx")

# Select the worksheet and starting cell
$worksheet = $workbook.Worksheets.Item(2)
$startRow = 4
$startRange = $worksheet.Range("A4:Y4")
#$startFormat = $startRange.Style 

$vmList = Get-VM -Location (Get-Folder -Name $projectName) | Sort

# Read the list of virtual machine names from file
#$vmNames = Get-Content "\\fshill\hedc\depot\Infrastructure\Technical\scripts\vmware\inputdata\CDRL_Test_A10.txt"

$itemID = 1

# Loop through the virtual machines and populate the cells
foreach ($vmName in $vmList) {
    # Get the virtual machine information using PowerCLI
    Write-Host "Setting Values for VM $vmName"
    $vm = Get-VM $vmName
    #$os = $vm.Guest.OSFullName
    $os = $vm.ExtensionData.Summary.Config.GuestFullName

    if ($os -like "Microsoft Windows Server*") {
        $osCaption = (Get-WmiObject Win32_OperatingSystem -ComputerName $vmName).Caption
        $os = $osCaption
    }

    $vcpuCount = $vm.NumCpu
    $memoryCount = $vm.MemoryGB
    #$storage = $vm.ExtensionData.Summary.Storage.Uncommitted
    $storage = $vm.ProvisionedSpaceGB.ToString("N2")

    #Write-Host "VM: $vmName, OS: $os, Memory: $memoryCount, Storage: $storage"

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
    $worksheet.Cells.Item($startRow, "P").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "Q").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "R").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "S").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "T").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "U").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "V").Value2 = "PBL"

    
    $startRange.Copy() | Out-Null
    $worksheet.Range("A$($startRow):Y$($startRow)").PasteSpecial(-4122) | Out-Null
    $worksheet.Range("A$($startRow):Y$($startRow)").EntireColumn.AutoFit() | Out-Null

    # Move to the next row
    $startRow++
    # Itterate ID
    $itemID++


}


$date = Get-Date -Format "yyyyMMdd"
$filename = "$date" + " $projectName" + " CDRL A007 Configuration Baseline.xlsx"
$filepath = "\\fshill\hedc\depot\Infrastructure\compliance\CDRL"
Write-Host "Saving file to $filepath\$filename"
$worksheet.SaveAs("$filePath\$filename")
$workbook.Close()
$excel.Quit()
