Write-Output "Xfinity Network Outage Logging Shell`n"
Write-Output "Purpose: To Track the Intermittent Disconnection of Internet Connectivity`n"
Write-Output "Location: North Charleston SC`n"
Write-Output "Router Used: Netgear Nighthawk AX1800 WiFi Router RAX10 - Firmware: 1.0.15.146 As of $(Get-Date)`n"
Write-Output "DNS Servers: Get Dynamically From ISP`n"
Write-Output "Connection Method: Ethernet`n"

# MAIN FUNCTION 0: GATHER DYNAMICALLY OBTAINED DOMAIN NAME SERVERS
Write-Output "GATHERING DYNAMICALLY OBTAINED DOMAIN NAME SERVERS USING TRACERT AT $(Get-Date)`n"
tracert 8.8.8.8 | Select-String -Pattern "comcast"
Write-Output ""

# MAIN FUNCTION 1: LOOP TEST-CONNECTION 
Write-Output "Testing Connection With 5 Second Delay From G'S MACHINE With IP Address: 10.0.0.2 to Netgear NightHawk Router With IP Address: 10.0.0.1`n"
for ($i = 1; ; $i++) {
    try {
        # PING TARGET
        $pingResult = Test-Connection -ComputerName 10.0.0.1 -Count 1 -ErrorAction Stop

        # SUCCESSFUL PING REPORTS DETAILED INFORMATION
        $pingResult | Select-Object @{Name='Time';Expression={Get-Date}}, Source, Address, Status | Format-Table -AutoSize -Wrap
    }
    catch {
        # FAILED PING REPORTS DATE, TIME & STATUS
        Write-Output "Ping Failed At $(Get-Date) - Status: Unsuccessful...Retrying"
    }
       # FIVE SECOND DELAY
       Start-Sleep -Seconds 5
}