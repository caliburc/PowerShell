Remove-Variable * -Force -ErrorAction silentlycontinue

$date = Get-Date -UFormat %Y%m%d_%H%M

$groupname = ""
while ($groupname -eq "") {
    $groupname = Read-Host "Provide Group Name"
    }

$group = Get-ADGroup -Identity $groupname -Properties Name, Mail, Member

#Write-Host ""
#Write-Host "Group Name: $($group.Name)"
#Write-Host "Group Email: $($group.Mail)"

$members = $group.Member | Get-ADObject -Properties Name, Mail
$table = $members | Select-Object Name, Mail | Sort-Object Name | Format-Table -AutoSize
$tableString = $table | Out-String
$header = "Group Name: $($group.Name)`nGroup Email: $($group.Mail)`n"
$output = $header + $tableString

$filename = "$date" + "_$groupname`_GroupMembers.txt"
$outputPath = Join-Path -Path "$PSScriptRoot\Groups-Output\" -ChildPath $filename 
if (!(Test-Path $outputPath)) {
    New-Item -ItemType File -Path $outputPath
    }

$output | Out-File -FilePath $outputPath -Append

Remove-Variable * -Force -ErrorAction silentlycontinue