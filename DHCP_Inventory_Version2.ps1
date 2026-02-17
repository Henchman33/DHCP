<#
.SYNOPSIS
  DHCP Inventory / Audit Script
.DESCRIPTION
  Collects DHCP Server, Scope, and Option data for auditing and inventory.
  Generates both CSV and searchable HTML report with full scope details.
  Compatible with Windows PowerShell 5.0+ (ISE).
#>

# -------------------------------
# CONFIGURATION
# -------------------------------
$DhcpServers = Get-DhcpServerInDC | Select-Object -ExpandProperty DnsName  # auto-detect DHCP servers from AD
# Or manually specify:
#$DhcpServers = "arbuepdhcpi01.myigt.com","atupspdhcphuba1.myigt.com","ausydpdhcphuba1.myigt.com","camntpdhcpi02.myigt.com","cnpekpdhcp01.myigt.com","gbmanpdhcphuba1.myigt.com","lasp-dhci01.myigt.com","rnop-dhci03.myigt.com","usrnopdhci01.myigt.com","usrnopdhci02.myigt.com","usrnopdhcphuba1.myigt.com","usrnopdhcphuba2.myigt.com"

# Create timestamped folder under user profile
$timestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")
$ReportFolder = Join-Path $env:USERPROFILE\Desktop ("DHCP Report $timestamp")
New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null

$CsvPath  = Join-Path $ReportFolder "DHCP_Inventory_$timestamp.csv"
$HtmlPath = Join-Path $ReportFolder "DHCP_Inventory_$timestamp.html"

Write-Host "Collecting DHCP Inventory..." -ForegroundColor Cyan

# -------------------------------
# FUNCTION
# -------------------------------
function Get-DhcpInventory {
    param([string[]]$Servers)

    $results = @()

    foreach ($serverName in $Servers) {
        Write-Host "Querying DHCP server: $serverName" -ForegroundColor Cyan

        try {
            $serverOptions = Get-DhcpServerv4OptionValue -ComputerName $serverName -ErrorAction Stop
            $server_dns    = ($serverOptions | Where-Object {$_.OptionId -eq 6}).Value
            $server_domain = ($serverOptions | Where-Object {$_.OptionId -eq 15}).Value
            $server_ntp    = ($serverOptions | Where-Object {$_.OptionId -eq 42}).Value
        } catch {
            Write-Warning "Unable to retrieve server-level options for ${serverName}: $_"
            $server_dns = $server_domain = $server_ntp = $null
        }

        try {
            $scopes = Get-DhcpServerv4Scope -ComputerName $serverName -ErrorAction Stop
        } catch {
            Write-Warning "Could not enumerate scopes on ${serverName}: $_"
            continue
        }

        foreach ($scope in $scopes) {
            $scopeId = $scope.ScopeId
            try {
                $scopeOptions = Get-DhcpServerv4OptionValue -ComputerName $serverName -ScopeId $scopeId -ErrorAction Stop
                $scope_dns    = ($scopeOptions | Where-Object {$_.OptionId -eq 6}).Value
                $scope_domain = ($scopeOptions | Where-Object {$_.OptionId -eq 15}).Value
                $scope_ntp    = ($scopeOptions | Where-Object {$_.OptionId -eq 42}).Value
                $scope_router = ($scopeOptions | Where-Object {$_.OptionId -eq 3}).Value
            } catch {
                Write-Warning "Unable to get options for ${serverName} scope ${scopeId}: $_"
                $scope_dns = $scope_domain = $scope_ntp = $scope_router = $null
            }

            $results += [PSCustomObject]@{
                DHCPServer     = $serverName
                ScopeName      = $scope.Name
                ScopeId        = $scope.ScopeId
                StartRange     = $scope.StartRange
                EndRange       = $scope.EndRange
                SubnetMask     = $scope.SubnetMask
                LeaseDuration  = $scope.LeaseDuration
                Scope_DNS      = ($scope_dns -join ",")
                Scope_Domain   = ($scope_domain -join ",")
                Scope_Router   = ($scope_router -join ",")
                Scope_NTP      = ($scope_ntp -join ",")
                Server_DNS     = ($server_dns -join ",")
                Server_Domain  = ($server_domain -join ",")
                Server_NTP     = ($server_ntp -join ",")
            }
        }
    }
    return $results
}

# -------------------------------
# RUN INVENTORY
# -------------------------------
$inventory = Get-DhcpInventory -Servers $DhcpServers
$inventory | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvPath

# -------------------------------
# GENERATE HTML REPORT
# -------------------------------
$style = @"
<style>
body { font-family: Arial; margin: 20px; }
h1 { color: #004085; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 5px; font-size: 12px; }
th { background: #eaeaea; position: sticky; top: 0; }
tr:nth-child(even) { background: #f9f9f9; }
#searchInput { padding: 8px; width: 40%; margin-bottom: 10px; }
</style>
<script>
function filterTable() {
 var input = document.getElementById('searchInput');
 var filter = input.value.toLowerCase();
 var table = document.getElementById('dhcpTable');
 var trs = table.tBodies[0].getElementsByTagName('tr');
 for (var i=0;i<trs.length;i++) {
  trs[i].style.display = trs[i].textContent.toLowerCase().indexOf(filter)>-1?'':'none';
 }
}
</script>
"@

$table = $inventory | ConvertTo-Html -Head $style -Title "DHCP Inventory Report" `
    -PreContent "<h1>DHCP Inventory Report</h1><p>Author: Stephen McKee - IGTPLC</p><input id='searchInput' onkeyup='filterTable()' placeholder='Search...' /><table id='dhcpTable'>" `
    -PostContent "</table><p>Generated: $(Get-Date)</p>"

$table | Out-File -FilePath $HtmlPath -Encoding UTF8

Write-Host "Inventory Completed!"
Write-Host "CSV:  $CsvPath"
Write-Host "HTML: $HtmlPath"
