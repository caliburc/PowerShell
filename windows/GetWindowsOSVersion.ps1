$date = Get-Date -UFormat %Y%m%d_%H%M

# Set the path to the text file containing the list of servers
#$serverList = "C:\temp\WindowsServerswithnoNessus.txt"
$serverlist = Read-Host "Provide path to list of servers:"
if((test-path $serverlist) -eq $false) {
    do {
        Write-Host "`nInvalid path!`n"
        $serverlist = (Read-Host "Provide path to list of servers")
    } until ((test-path $serverlist) -eq $true)
}

# Create an empty array to store the results
$results = @()

# Read the contents of the text file into an array
$servers = Get-Content $serverlist

# Loop through the array of servers
foreach ($server in $servers) {
   # Use the Get-WmiObject cmdlet to retrieve the operating system version
   $os = Get-WmiObject -Class Win32_OperatingSystem -ComputerName $server

   # Create a new object to store the server name and operating system version
   $obj = New-Object PSObject -Property @{
      Server = $server
      OperatingSystem = $os.Caption
   }

   # Add the object to the array of results
   $results += $obj
}

# Set the path to the CSV file that you want to create
$csvFile = "C:\temp\" + $date + "_WindowsServerOSVersion.csv"

# Export the array of results to the CSV file
$results | Export-Csv $csvFile -NoTypeInformation
