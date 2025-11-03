<#
.SYNOPSIS
  DHCP Inventory Script with IP Highlighting and Excel Export
.DESCRIPTION
  Collects DHCP servers, scopes, and option details.
  Highlights scopes or server options containing specific IPs.
  Exports to CSV, HTML, and Excel formats.
  Author: Stephen McKee - IGTPLC
#>

# -------------------------------
# CONFIGURATION
# -------------------------------
$timestamp     = (Get-Date).ToString("yyyyMMdd_HHmmss")
$ReportFolder  = Join-Path $env:USERPROFILE ("DHCP Report $timestamp")
$CsvReport     = Join-Path $ReportFolder "DHCP_Inventory_$timestamp.csv"
$HtmlReport    = Join-Path $ReportFolder "DHCP_Inventory_$timestamp.html"
$ExcelReport   = Join-Path $ReportFolder "DHCP_Inventory_$timestamp.xlsx"

New-Item -Path $ReportFolder -ItemType Directory -Force | Out-Null

# Target IPs to highlight
$OldIPs = @("10.210.254.240","10.210.254.241","10.210.254.242")

# -------------------------------
# ENUMERATE DHCP SERVERS
# -------------------------------
try {
    $DhcpServers = Get-DhcpServerInDC | Select-Object -ExpandProperty DnsName
} catch {
    Write-Error "Unable to enumerate DHCP servers from AD. Run PowerShell as admin."
    exit
}

$results = @()

# -------------------------------
# INVENTORY COLLECTION
# -------------------------------
foreach ($server in $DhcpServers) {
    Write-Host "`nQuerying DHCP Server: $server" -ForegroundColor Cyan

    try {
        $serverOptions = Get-DhcpServerv4OptionValue -ComputerName $server -ErrorAction Stop
        $scopes        = Get-DhcpServerv4Scope -ComputerName $server -ErrorAction Stop
    } catch {
        Write-Warning "⚠️  Failed to query {$server}: $_"
        continue
    }

    # ---- SERVER-LEVEL OPTIONS ----
    $dnsServer = $serverOptions.DnsServer
    $ntpOption = ($serverOptions | Where-Object {$_.OptionId -eq 42}).Value
    $dnsDomain = $serverOptions.DnsDomain

    $serverMatchesOldIP = ($OldIPs | Where-Object { $dnsServer -contains $_ -or $ntpOption -contains $_ })

    $results += [PSCustomObject]@{
        Type          = "ServerOption"
        DHCPServer    = $server
        ScopeId       = ""
        ScopeName     = ""
        DNS           = ($dnsServer -join ",")
        NTP           = ($ntpOption -join ",")
        Router        = ""
        Domain        = $dnsDomain
        ContainsOldIP = if ($serverMatchesOldIP) { "YES" } else { "" }
    }

    # ---- SCOPE-LEVEL OPTIONS ----
    foreach ($scope in $scopes) {
        $scopeId   = $scope.ScopeId
        $scopeName = $scope.Name

        try {
            $scopeOptions = Get-DhcpServerv4OptionValue -ComputerName $server -ScopeId $scopeId -ErrorAction Stop
        } catch {
            Write-Warning "  ⚠️  Unable to get options for scope $scopeId on {$server}: $_"
            continue
        }

        $dnsOpt    = ($scopeOptions | Where-Object { $_.OptionId -eq 6 }).Value
        $routerOpt = ($scopeOptions | Where-Object { $_.OptionId -eq 3 }).Value
        $ntpOpt    = ($scopeOptions | Where-Object { $_.OptionId -eq 42 }).Value
        $domainOpt = ($scopeOptions | Where-Object { $_.OptionId -eq 15 }).Value

        $containsOld = $false
        foreach ($ip in $OldIPs) {
            if ($dnsOpt -contains $ip -or $routerOpt -contains $ip -or $ntpOpt -contains $ip) {
                $containsOld = $true
                break
            }
        }

        $results += [PSCustomObject]@{
            Type          = "Scope"
            DHCPServer    = $server
            ScopeId       = $scopeId
            ScopeName     = $scopeName
            DNS           = ($dnsOpt -join ",")
            NTP           = ($ntpOpt -join ",")
            Router        = ($routerOpt -join ",")
            Domain        = ($domainOpt -join ",")
            ContainsOldIP = if ($containsOld) { "YES" } else { "" }
        }
    }
}

# -------------------------------
# EXPORT TO CSV
# -------------------------------
$results | Export-Csv -NoTypeInformation -Encoding UTF8 -Path $CsvReport

# -------------------------------
# EXPORT TO HTML
# -------------------------------
$style = @"
<style>
body { font-family: Arial; margin:20px; }
h1 { color:#004085; }
table { border-collapse:collapse; width:100%; }
th,td { border:1px solid #ccc; padding:5px; font-size:12px; }
th { background:#eaeaea; position:sticky; top:0; }
tr:nth-child(even){background:#f9f9f9;}
tr.highlight { background-color:#ffeeba !important; }
#searchInput{padding:8px;width:40%;margin-bottom:10px;}
</style>
<script>
function filterTable(){
 var i=document.getElementById('searchInput'),f=i.value.toLowerCase();
 var t=document.getElementById('dhcpTable'),r=t.tBodies[0].getElementsByTagName('tr');
 for(var x=0;x<r.length;x++){r[x].style.display=r[x].textContent.toLowerCase().indexOf(f)>-1?'':'none';}
}
</script>
"@

$htmlBody = @()
$htmlBody += "<h1>DHCP Inventory Report</h1>"
$htmlBody += "<p>Author: Stephen McKee - IGTPLC</p>"
$htmlBody += "<input id='searchInput' onkeyup='filterTable()' placeholder='Search...' />"
$htmlBody += "<table id='dhcpTable'><thead><tr><th>Type</th><th>DHCP Server</th><th>Scope ID</th><th>Scope Name</th><th>DNS</th><th>NTP</th><th>Router</th><th>Domain</th><th>ContainsOldIP</th></tr></thead><tbody>"

foreach ($r in $results) {
    $class = if ($r.ContainsOldIP -eq "YES") { "class='highlight'" } else { "" }
    $htmlBody += "<tr $class><td>$($r.Type)</td><td>$($r.DHCPServer)</td><td>$($r.ScopeId)</td><td>$($r.ScopeName)</td><td>$($r.DNS)</td><td>$($r.NTP)</td><td>$($r.Router)</td><td>$($r.Domain)</td><td>$($r.ContainsOldIP)</td></tr>"
}

$htmlBody += "</tbody></table><p>Generated: $(Get-Date)</p>"
$fullHtml = "<html><head>$style</head><body>$($htmlBody -join '')</body></html>"
$fullHtml | Out-File -Encoding UTF8 -FilePath $HtmlReport

# -------------------------------
# EXPORT TO EXCEL (using ImportExcel)
# -------------------------------
try {
    Import-Module ImportExcel -ErrorAction Stop

    $ExcelData = $results | Sort-Object DHCPServer, Type, ScopeId
    $ExcelData | Export-Excel -Path $ExcelReport -AutoSize -BoldTopRow -FreezeTopRow -Title "DHCP Inventory Report - $timestamp" `
        -WorksheetName 'DHCP Inventory' -AutoFilter -TableName 'DHCPInventory' -ClearSheet

    # Add conditional formatting to highlight old IP rows
    $excelPkg = Open-ExcelPackage -Path $ExcelReport
    $ws = $excelPkg.Workbook.Worksheets["DHCP Inventory"]
    $rowCount = $ws.Dimension.End.Row
    Add-ConditionalFormatting -Worksheet $ws -Range "I2:I$rowCount" -RuleType ContainsText -ConditionValue "YES" -BackgroundColor Yellow
    Close-ExcelPackage $excelPkg
} catch {
    Write-Warning "⚠️ Excel export skipped or ImportExcel module not available: $_"
}

# -------------------------------
# SUMMARY
# -------------------------------
Write-Host "✅ DHCP inventory complete!"
Write-Host "CSV Report:   $CsvReport"
Write-Host "HTML Report:  $HtmlReport"
Write-Host "Excel Report: $ExcelReport"
Write-Host "Highlighted entries indicate presence of old IPs ($($OldIPs -join ', '))."
