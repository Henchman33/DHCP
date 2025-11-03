<# Change line #10 your DHCP server FQDN Name
   Change line #20 for logfile path you want
   Change line #58 for whatif statement to -whatif - Will not make any changes, provides a "dryrun" before commenting it out. Comment it out when you are ready to make the change # -whatif  
   again
   -whatif is to run without changing anything
   #-whatif runs the script and makes the changes
Thanks to JNS!!!!
#>
# DHCP server name
$dhcpServer = "DHCPXXX.com"
 
# Old → New DNS mapping
$dnsMap = @{
    "10.X.X.X" = "10.X.X.X"  #<--- left IP is OLD IP - right IP is new IP :)
    "10.X.X.X" = "10.X.X.X"
    "10.X.X.X" = "10.X.X.X"
}
 
# CSV log file
$logFile = "C:\Temp\DHCP_DNS_Update_Log_usrnopdhci02.myigt.com.csv"
$logEntries = @()  # Collect logs in memory
$changedCount = 0
$skippedCount = 0
 
# Get all scopes
$scopes = Get-DhcpServerv4Scope -ComputerName $dhcpServer
 
foreach ($scope in $scopes) {
    $scopeId = $scope.ScopeId
    $current = (Get-DhcpServerv4OptionValue -ComputerName $dhcpServer -ScopeId $scopeId -OptionId 006 -ErrorAction SilentlyContinue).Value
 
    # Skip if no DNS option configured
    if (-not $current) {
        $skippedCount++
        continue
    }
 
    # Replace only old DNS entries
    $newDns = $current | ForEach-Object {
        if ($dnsMap.ContainsKey($_)) {
            $dnsMap[$_]
        } else {
            $_
        }
    }
 
    # Compare joined strings to detect real changes
    if (($newDns -join ',') -eq ($current -join ',')) {
        $skippedCount++
        continue
    }
 
    # Simulate applying changes ONLY for changed scopes
    Set-DhcpServerv4OptionValue `
        -ComputerName $dhcpServer `
        -ScopeId $scopeId `
        -DnsServer ([string[]]$newDns) `
        -WhatIf
 
    # Add log entry as an object for proper CSV export
    $logEntries += [PSCustomObject]@{
        ScopeId    = $scopeId
        OldDNS     = ($current -join ", ")
        NewDNS     = ($newDns -join ", ")
        Timestamp  = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    }
 
    $changedCount++
}
 
# Export all logs at once with header (overwrite file)
$logEntries | Export-Csv -Path $logFile -NoTypeInformation -Force
 
# Summary
Write-Host "Summary:"
Write-Host "Scopes processed: $($scopes.Count)"
Write-Host "Scopes changed: $changedCount"
Write-Host "Scopes skipped: $skippedCount"
Write-Host "Log file: $logFile"
