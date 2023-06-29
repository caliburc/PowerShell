#Remove-Variable * -Force -ErrorAction silentlycontinue

# Connect to vCenter
if (($global:DefaultVIServers).Name -like "apvcsa1p*") {
    Write-Host "Already Connected to vCenter"
    }else {
    Write-Host "Connecting to vCenter"
    Connect-VIServer apvcsa1p.idm.hedc.mil
}
# Define Customer VMs 
$projectName = Read-Host "Enter Unique Identifyer for this sheet (ex. `"HEDC`")"

# Import the Excel COM object
$excel = New-Object -ComObject Excel.Application

# Open the Template CDRL Workbook
$workbook = $excel.Workbooks.Open("\\fshill\hedc\depot\Infrastructure\compliance\CDRL\CDRL A007 Configuration Baseline Development - Template V3.xlsx")

# Select the worksheet
$worksheet = $workbook.Worksheets.Item(2)

#Get the VM's from the input projectName provided and sort them
$vmList = Get-VM | Where-Object {$_.PowerState -eq "PoweredOn"} | Sort-Object -Property ID

#set the default values for the progress bar
$vmCount = $vmList.Count
$progress = 0
$vmNum = 1

$startRow = 4
#select the starting row range to copy from for formatting
$startRange = $worksheet.Range("A4:AD4")
# the dumb thing wasn't autoformatting correct, this sets the column wide, and then at the end of the script it resizes the whole sheet
$worksheet.Columns.Item("N").ColumnWidth = 70

foreach ($vmName in $vmList) {
    #Write-Host "Setting Values for VM $vmName"
    $vm = Get-VM $vmName

    # Update Progress bar
    $progress = [Math]::Round($vmNum / $vmCount * 100)
    Write-Progress -Activity "Setting Values for $vmName" -Status "Processing VM $vmNum of $vmCount" -PercentComplete $progress

    $itemID = $vm.id
    $os = $vm.ExtensionData.Summary.Config.GuestFullName
    # VMware is not always correct with what version of Windows, this checks the OS version from within the server itself.
    # Sometimes we might not have permission to check this or the machine is not on the domain, these catch's check for that and use the VMware OS value if that's the case
    if ($os -like "Microsoft Windows Server*") {
        try {
        $osCaption = (Get-WmiObject Win32_OperatingSystem -ComputerName $vmName -ErrorAction Stop).Caption
        $os = $osCaption
        }
        catch [System.Runtime.InteropServices.COMException] {
            if ($_.Exception.Message -like "*The RPC server is unavailable*") {
                Write-Warning "RPC Server is unavailable when getting OS info for $vmName - Using VMware OS information instead"
            }
            else {
                Write-Warning "Error geting OS infromation for $vmName : $($_.Exception.Message) - Using VMware OS information isntead"
            }
        }
        catch {
            Write-Warning "Error getting OS information on VM $vmName : $($_.Exception.Message) - Using VMware OS information instead"
        }
    }
    $esxihost = $vm.VMHost
    $cluster = (Get-VMHost $esxihost).Parent.Name
    $vcpuCount = $vm.NumCpu
    $memoryCount = $vm.MemoryGB
    #Uses GB values to the second decimal place
    $storage = $vm.ProvisionedSpaceGB.ToString("N2")
    #if the VM uses multiple datstores, it splits them and puts them into a list
    $datastore = $vm.ExtensionData.Config.DatastoreUrl.Name -split ' ' -join "`n"


    # Populate the cells with the virtual machine information
    $worksheet.Cells.Item($startRow, "D").Value2 = "x"
    $worksheet.Cells.Item($startRow, "G").Value2 = "$itemID"
    $worksheet.Cells.Item($startRow, "H").Value2 = "$vmName"
    $worksheet.Cells.Item($startRow, "I").Value2 = "1"
    $worksheet.Cells.Item($startRow, "J").Value2 = "VM"
    $worksheet.Cells.Item($startRow, "N").Value2 = "$datastore"
    $worksheet.Cells.Item($startRow, "P").Value2 = "1"
    $worksheet.Cells.Item($startRow, "Q").Value2 = "$os"
    $worksheet.Cells.Item($startRow, "R").Value2 = "$vcpuCount"
    $worksheet.Cells.Item($startRow, "S").Value2 = "$memoryCount"
    $worksheet.Cells.Item($startRow, "T").Value2 = "$storage"
    $worksheet.Cells.Item($startRow, "U").Value2 = "$esxihost"
    $worksheet.Cells.Item($startRow, "V").Value2 = "$cluster"
    $worksheet.Cells.Item($startRow, "W").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "X").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "Y").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "Z").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "AA").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "AB").Value2 = "N/A"
    $worksheet.Cells.Item($startRow, "AC").Value2 = "PBL"

    # Copy The Formatting from the first row and paste it to this row, also autofit the columns
    $startRange.Copy() | Out-Null
    $worksheet.Range("A$($startRow):AD$($startRow)").PasteSpecial(-4122) | Out-Null
    $worksheet.Range("A$($startRow):AD$($startRow)").EntireColumn.AutoFit() | Out-Null
    $worksheet.Rows.Item($startRow).EntireRow.AutoFit() | Out-Null

    # Move to the next row
    $startRow++
    # Increment Progress
    $vmNum++
}

#autofit the entire sheet
$worksheet.Columns.AutoFit() | Out-Null
$worksheet.Rows.Autofit() | Out-Null

$date = Get-Date -Format "yyyyMMdd"
$filename = "$date" + " $projectName" + " CDRL A007 Configuration Baseline.xlsx"
$filepath = "\\fshill\hedc\depot\Infrastructure\compliance\CDRL"
Write-Host "Saving file to $filepath\$filename"
$worksheet.SaveAs("$filePath\$filename")
$workbook.Close()
$excel.Quit()
sleep 8

#Remove-Variable * -Force -ErrorAction silentlycontinue