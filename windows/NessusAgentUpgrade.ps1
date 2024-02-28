
<#################################################################################
    Author: Jason Johnson / 1165077764
    Date: 2022-01-06
    Purpose: Install Nessus Agent or Upgrade Agent to Desired Version
 #################################################################################
                                How to Use  
    Prerequisites:
    1.) Change the installerfilename to the current Agent MSI file name
        Ex. CM-285654_Nessu.sAgent-10.2.1-x64.msi
    2.) Change the target version the full version number of the version installing -
        This number can be found by just installing the agent on a machine manually 
         and then looking at the Installed Programs on the machine (ex. 10.3.2.20006)
    3.) Make sure the installer is at 
        \\fshill\hedc\depot\software\ACAS\NessusAgent\MS\.
         Or change the installerpath variable to your desired location

    Running:
    When running, input the group or groups of the systems you will be adding.
    This can be a comma seperated list and the groups can have spaces
    Then put in a file path for to a text document with a list of servers

################################################################################>

CLS
Remove-Variable * -Force -ErrorAction silentlycontinue
$logDate = Get-Date -UFormat %Y%m%d_%H%M

########  Modify for version #########
$installerfilename = "CM-299986_NessusAgent-10.4.4-x64.msi"
$targetversion = "10.4.4.20007"
######################################

$logPath = "C:\Temp\" + "$logDate" + "_NessusInstall_Log.txt"
$msipath = "C:\Temp\" + $installerfilename
$nessusserver = "137.241.12.22:8934"
$nessuskey = "3fd5af3fa780c560e21a7425b64511f815273fbf9d70e9141ea9aa61239f091c" #obtained from nessus manager "linked agent" screen
$InstallAgent = $false 

do {
  Write-Host "[1] Windows 2012"
  Write-Host "[2] Windows 2016"
  Write-Host "[3] Windows 2019"
  Write-Host "[4] Windows 10"
  Write-Host "[5] Windows 2012 (PMO-JET)"
  Write-Host "[6] Windows 2016 (PMO-Geobase)"
  $userinput = Read-Host "Select an option (1-6)"
  
  switch ($userinput) {
    "1" { $nessusgroup = "Windows2012" }
    "2" { $nessusgroup = "Windows2016" }
    "3" { $nessusgroup = "Windows2019" }
    "4" { $nessusgroup = "Windows10" }
    "5" { $nessusgroup = "Windows2012,PMO - JET" }
    "6" { $nessusgroup = "Windows2016,PMO - Geobase" }
    Default { Write-Warning "`nInvalid option selected. Try again you silly goose." }
    } 
} until ($nessusgroup)

if ($nessusgroup -eq "Windows2012,PMO - JET") {
    $defaultList = "\\fshill\hedc\depot\Infrastructure\Technical\NessusAgent\AgentLists\Windows2012-JET-Serverlist.txt"
} elseif ($nessusgroup -eq "Windows2016,PMO - Geobase") {
    $defaultList = "\\fshill\hedc\depot\Infrastructure\Technical\NessusAgent\AgentLists\Windows2016-Geobase-Serverlist.txt"
} else {
    $defaultList = "\\fshill\hedc\depot\Infrastructure\Technical\NessusAgent\AgentLists\${nessusgroup}-Serverlist.txt"
}
do {
    Write-Host "The default server list for this group(s) is: " -NoNewLine;Write-Host -ForegroundColor Green $defaultList -NoNewline
    $confirmDefault = Read-Host "`nDo you want to use this list? (Y/N)"
    $confirmDefault = $confirmDefault.ToLower()
} until ($confirmDefault -eq "y" -or $confirmDefault -eq "n")

if ($confirmDefault -eq "y" -or $confirmDefault -eq "yes") {
    Write-Host -ForegroundColor Cyan "Using default serverlist: $defaultList"
    $serverList = $defaultList
    if((test-path $serverlist) -eq $false) {
        do {
            Write-Host -ForegroundColor Red "`nCouldn't load the default server list!`n"
            $serverList = (Read-Host "Provide path to list of servers")
        } until ((test-path $serverlist) -eq $true)
    }
} else {
    $serverList = Read-Host "Provide path to list of servers"
    if((test-path $serverlist) -eq $false) {
        do {
           Write-Host -ForegroundColor Red "`nCouldn't load your server list, check path and try again!`n"
           $serverList = (Read-Host "Provide path to list of servers")
        } until ((test-path $serverlist) -eq $true)
    }
}

$servers = get-content $serverList

Write-Host -ForegroundColor Cyan "Writing Log to $logPath"
Add-Content -Path $logpath -Value "Script being ran by: $env:USERNAME`n"
Add-Content -Path $logpath -Value "Running against group $nessusgroup`n"

Write-Host $servers

foreach ($server in $servers) {
    $count = 0
    $installerpath = "\\fshill\HEDC\depot\software\ACAS\NessusAgent\MS\" + $installerfilename
    $remotepath = "\\" + $server + ".area52.afnoapps.usaf.mil\C$\Temp\" + $installerfilename
    $remotepathtest = "\\" + $server + ".area52.afnoapps.usaf.mil\C$\Temp"
    $nessuscli = "\\" + $server + ".area52.afnoapps.usaf.mil\C$\Program Files\Tenable\Nessus Agent\nessuscli.exe"
    $InstallAgent = $false
    Try {
        # Check if Nessus is installed and print out the current version installed
        Write-Host "`nChecking Nessus Version on" $server 
        $NessusCurrentInstall = Get-WmiObject -Class Win32_Product -ComputerName ($server + ".area52.afnoapps.usaf.mil") | where name -Like "Nessus*"
        Write-Host "Nessus Agent " -NoNewline;Write-Host -ForegroundColor Green $($NessusCurrentInstall.Version) -NoNewline;Write-Host  " is currently installed on $server"
        Add-Content -Path $logpath -Value "Nessus Agent $($NessusCurrentInstall.Version) is currently installed on $server"
    }
    Catch [System.Runtime.InteropServices.COMException] {
      if ($_.Exception.Message -like "*The RPC server is unavailable*") {
        Write-Warning "RPC Server is unavailable - skipping $server"
        Continue
        }
    }
    If(!$NessusCurrentInstall){
        # If the agent is not installed, we will install it using the required arguments
       $InstallAgent = $true
       } 
       Else {
            if($NessusCurrentInstall.Version -ne $targetversion){
                # If the agent is no the same as our target version, we will install/upgrade to our target version
                Write-Host "Agent needs to be upgraded on $server"
                Add-Content -Path $logpath -Value "Agent needs to be upgraded on $server"
                $UpgradeAgent = $true
            } 
            else {
                # The agent already matches the target version, do nothing
                Write-Host "Nothing to do on $server"
                Add-Content -Path $logpath -Value "Nothing to do on $server"
                $InstallAgent = $false
            }
        }
     if($InstallAgent){
            # Install from a blank slate, use msi parameters to link to nessus
            Write-Host "Nessus Agent not installed on $server `n Installing Agent Version " -NoNewLine;Write-Host -ForegroundColor Red $targetversion
            Add-Content -Path $logpath -Value "Nessus Agent not installed on $server - Installing Agent $targetversion"
            #Check if Temp directory exists on sever
            if(-not(Test-Path $remotepathtest)){
                New-Item -Path $remotepathtest -ItemType Directory
                }
            #Copy-Item -Path "\\fshill\HEDC\depot\software\ACAS\NessusAgent\MS\CM290173_NessusAgent-10.3.2-x64.msi" -Destination `"$remotepath`"
            Copy-Item -Path $installerpath -Destination $remotepath
            #This never seems to link correctly, saving for continuity
            #Invoke-WmiMethod win32_process -name create -ComputerName ($server + ".area52.afnoapps.usaf.mil") -ArgumentList "msiexec /i $msipath NESSUS_GROUPS=`"$nessusgroup`" NESSUS_SERVER=`"$nessusserver`" NESSUS_KEY=$nessuskey /qn"
            Invoke-WmiMethod win32_process -name create -ComputerName ($server + ".area52.afnoapps.usaf.mil") -ArgumentList "msiexec /i $msipath /qn"
            sleep 5
            #Following loop checks to see if the nessuscli.exe is installed - then runs the linking command.
            do {
                if(-not(Test-Path "$nessuscli" -PathType Leaf)){
                    $count++
                    Write-Host "Waiting for Nessus Agent to be installed for linking, try number ${count}/5, standby ..."
                    sleep 12
                    } else {
                    write-host "Installed"
                    #psexec -nobanner \\($server + ".area52.afnoapps.usaf.mil") -h -s "C:\Program Files\Tenable\Nessus Agent\nessuscli.exe" agent link --host=137.241.12.22 --port=8934 --key=3fd5af3fa780c560e21a7425b64511f815273fbf9d70e9141ea9aa61239f091c \ --groups=`"$nessusgroup`" 2>$null
                    $count = 5
                    }
             }until ($count -eq 5)
}
     
     if($UpgradeAgent){
            # Just upgrade the agent - existing agent linking should stay in place
            Write-Host "Upgrading Nessus Agent version on $server"
            if(-not(Test-Path $remotepathtest)){
                New-Item -Path $remotepathtest -ItemType Directory
            }
            Copy-Item -Path $installerpath -Destination $remotepath
            $result = Invoke-WmiMethod win32_process -name create -ComputerName ($server + ".area52.afnoapps.usaf.mil") -ArgumentList "msiexec /i $msipath /qn"
     }
}

Write-Host -ForegroundColor Green "Log written to $logPath"
Write-Host -ForegroundColor Green "Done!"


Remove-Variable * -Force -ErrorAction silentlycontinue
