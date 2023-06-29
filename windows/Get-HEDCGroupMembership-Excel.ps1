Remove-Variable * -Force -ErrorAction silentlycontinue

# Read list of group names from text file
$groupNames = Get-Content -Path "$PSScriptRoot\Input\groupsmanagedby_HEDC_Manage_HEDC_Access.txt"

$date = Get-Date -UFormat %Y%m%d_%H%M

# Create new Excel workbook
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()

#Create a Table of Contents
$tocSheet = $workbook.Worksheets.Add()
$tocSheet.Name = "Table of Contents"
$tocSheet.Cells.Item(1,1) = "Group Sheet Links"
$tocSheetRow = 2


# Loop through each group and add to Excel workbook
foreach ($groupName in $groupNames) {
    # Get group object
    $group = Get-ADGroup -Identity "$groupName" -Properties Name, Mail, Member
    
    Write-Host $group.Name
    # Add new worksheet to workbook with group name as tab name
    $groupNamedTrimmed = $group.Name.Substring(0, [Math]::Min($group.Name.Length, 31))
    $worksheet = $workbook.Worksheets.Add()
    $worksheet.Name = $groupNamedTrimmed
    
    # Add return to Table of Contents Link on each sheet
    $worksheet.Hyperlinks.Add($worksheet.Cells.Item(1,1), "", "'Table of Contents'!A1", "Return To Table of Contents", "Return to Table of Contents") | Out-Null
   

     # Add group name and email to worksheet
    $worksheet.Cells.Item(2,1) = "Group Name:"
    $worksheet.Cells.Item(2,2) = $group.Name
    $worksheet.Cells.Item(3,1) = "Email Address:"
    $worksheet.Cells.Item(3,2) = $group.Mail

    $worksheet.Cells.Item(5,1) = "Name"
    $worksheet.Cells.Item(5,2) = "Email"
    
    # Get group members and add to worksheet
    $members = $group.Member | Get-ADObject -Properties Name, Mail
    $rowCount = 6
    foreach ($member in $members) {
        $worksheet.Cells.Item($rowCount,1) = $member.Name
        $worksheet.Cells.Item($rowCount,2) = $member.Mail
        $rowCount++
    }

    # Add the group to the TOC
    $tocSheet.Cells.Item($tocSheetRow,1) = $group.Name
    $hyperlink = $tocSheet.Hyperlinks.Add($tocSheet.Cells.Item($tocSheetRow,1), "", "'$($group.Name)'!A1", "")
    $hyperlink | Out-Null
    $tocSheetRow++
    $worksheet.Columns.AutoFit() | Out-Null
}

# Save and close Excel workbook
$tocSheet.Move($workbook.Sheets.Item(1))
$tocSheet.Columns.AutoFit() | Out-Null
$filename = "$PSScriptRoot\Groups-Output\$date`_groupsmanagedby_HEDC_Manage_HEDC_Access.xlsx"
$workbook.SaveAs($filename)
$excel.Quit()
Write-Host "Saved file to $filename"

Remove-Variable * -Force -ErrorAction silentlycontinue
