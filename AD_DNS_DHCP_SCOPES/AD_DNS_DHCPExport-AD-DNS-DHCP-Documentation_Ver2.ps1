# ============================================================
# Enterprise AD DNS DHCP Documentation Script
# Author: Stephen McKee - Everi - IGT Server Administrator 2
# Version: Enterprise + Inactive Scope Highlighting
# ============================================================

#region PATH SETUP

$DesktopPath = [Environment]::GetFolderPath("Desktop")
$OutputFolder = Join-Path $DesktopPath "AD_DNS_DHCP_Documentation"

if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$ErrorLog = Join-Path $OutputFolder "AD_DNS_DHCP_ERRORS.log"

function Write-ErrorLog {
    param ($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $ErrorLog -Value "$TimeStamp - $Message"
}

#endregion

#region MODULES

try { Import-Module ActiveDirectory -ErrorAction Stop } catch { Write-ErrorLog $_ }
try { Import-Module DnsServer -ErrorAction Stop } catch { Write-ErrorLog $_ }
try { Import-Module DhcpServer -ErrorAction Stop } catch { Write-ErrorLog $_ }

#endregion

#region DISCOVERY

try {
    $DnsServers = Get-ADComputer -Filter {ServicePrincipalName -like "*DNS*"} -Properties DNSHostName |
                  Select-Object Name, DNSHostName
} catch { Write-ErrorLog $_ }

try {
    $DhcpServers = Get-DhcpServerInDC
} catch { Write-ErrorLog $_ }

#endregion

#region DNS DATA COLLECTION

$DnsZones = @()
$DnsRecords = @()
$DnsScavenging = @()
$ConditionalForwarders = @()
$DnsServerSettings = @()

foreach ($Server in $DnsServers) {

    try {

        $Settings = Get-DnsServer -ComputerName $Server.Name
        $DnsServerSettings += [PSCustomObject]@{
            Server = $Server.Name
            IsADIntegrated = $Settings.IsDsIntegrated
            Forwarders = ($Settings.Forwarders -join ", ")
            ScavengingInterval = $Settings.ScavengingInterval
        }

        $Zones = Get-DnsServerZone -ComputerName $Server.Name

        foreach ($Zone in $Zones) {

            $DnsZones += [PSCustomObject]@{
                Server = $Server.Name
                ZoneName = $Zone.ZoneName
                ZoneType = $Zone.ZoneType
                DynamicUpdate = $Zone.DynamicUpdate
                AgingEnabled = $Zone.AgingEnabled
            }

            if ($Zone.AgingEnabled) {
                $Aging = Get-DnsServerZoneAging -ComputerName $Server.Name -Name $Zone.ZoneName
                $DnsScavenging += [PSCustomObject]@{
                    Server = $Server.Name
                    Zone = $Zone.ZoneName
                    NoRefreshInterval = $Aging.NoRefreshInterval
                    RefreshInterval = $Aging.RefreshInterval
                }
            }

            if ($Zone.ZoneType -eq "Forwarder") {
                $ConditionalForwarders += [PSCustomObject]@{
                    Server = $Server.Name
                    ZoneName = $Zone.ZoneName
                    MasterServers = ($Zone.MasterServers -join ", ")
                }
            }

            $Records = Get-DnsServerResourceRecord -ComputerName $Server.Name -ZoneName $Zone.ZoneName

            foreach ($Record in $Records) {
                $DnsRecords += [PSCustomObject]@{
                    Server = $Server.Name
                    Zone = $Zone.ZoneName
                    HostName = $Record.HostName
                    RecordType = $Record.RecordType
                    TimeStamp = $Record.Timestamp
                }
            }
        }

    } catch { Write-ErrorLog $_ }
}

#endregion

#region DHCP DATA COLLECTION

$DhcpScopes = @()
$DhcpReservations = @()
$DhcpLeases = @()
$InactiveScopes = @()

foreach ($Server in $DhcpServers) {

    try {

        $Scopes = Get-DhcpServerv4Scope -ComputerName $Server.DnsName

        foreach ($Scope in $Scopes) {

            $DhcpScopes += [PSCustomObject]@{
                Server = $Server.DnsName
                ScopeName = $Scope.Name
                ScopeID = $Scope.ScopeId
                StartRange = $Scope.StartRange
                EndRange = $Scope.EndRange
                State = $Scope.State
            }

            # 🚨 Detect Inactive Scopes
            if ($Scope.State -ne "Active") {
                $InactiveScopes += [PSCustomObject]@{
                    Server = $Server.DnsName
                    ScopeName = $Scope.Name
                    ScopeID = $Scope.ScopeId
                    State = $Scope.State
                }
            }

            $Reservations = Get-DhcpServerv4Reservation -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId
            foreach ($Res in $Reservations) {
                $DhcpReservations += [PSCustomObject]@{
                    Server = $Server.DnsName
                    ScopeID = $Scope.ScopeId
                    IPAddress = $Res.IPAddress
                    ClientID = $Res.ClientId
                    Name = $Res.Name
                }
            }

            $Leases = Get-DhcpServerv4Lease -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId
            foreach ($Lease in $Leases) {
                $DhcpLeases += [PSCustomObject]@{
                    Server = $Server.DnsName
                    ScopeID = $Scope.ScopeId
                    IPAddress = $Lease.IPAddress
                    HostName = $Lease.HostName
                    AddressState = $Lease.AddressState
                }
            }
        }

    } catch { Write-ErrorLog $_ }
}

#endregion

#region EXECUTIVE SUMMARY

$ExecutiveSummary = [PSCustomObject]@{
    TotalDNSServers = $DnsServers.Count
    TotalDNSZones = $DnsZones.Count
    TotalDNSRecords = $DnsRecords.Count
    ConditionalForwarders = $ConditionalForwarders.Count
    TotalDHCPServers = $DhcpServers.Count
    TotalScopes = $DhcpScopes.Count
    InactiveScopes = $InactiveScopes.Count
    TotalReservations = $DhcpReservations.Count
    ActiveLeases = $DhcpLeases.Count
}

#endregion

#region EXPORT CSV

$DnsZones | Export-Csv "$OutputFolder\DNS_Zones.csv" -NoTypeInformation
$DnsRecords | Export-Csv "$OutputFolder\DNS_Records.csv" -NoTypeInformation
$DnsScavenging | Export-Csv "$OutputFolder\DNS_Scavenging.csv" -NoTypeInformation
$ConditionalForwarders | Export-Csv "$OutputFolder\DNS_ConditionalForwarders.csv" -NoTypeInformation
$DhcpScopes | Export-Csv "$OutputFolder\DHCP_Scopes.csv" -NoTypeInformation
$DhcpReservations | Export-Csv "$OutputFolder\DHCP_Reservations.csv" -NoTypeInformation
$DhcpLeases | Export-Csv "$OutputFolder\DHCP_Leases.csv" -NoTypeInformation
$InactiveScopes | Export-Csv "$OutputFolder\DHCP_Inactive_Scopes.csv" -NoTypeInformation

#endregion

#region EXPORT XLSX (ImportExcel Required)

try {
    Import-Module ImportExcel -ErrorAction Stop
    $ExcelPath = "$OutputFolder\AD_DNS_DHCP_Enterprise.xlsx"

    $ExecutiveSummary | Export-Excel $ExcelPath -WorksheetName "Executive Summary" -AutoSize
    $DnsZones | Export-Excel $ExcelPath -WorksheetName "DNS Zones" -AutoSize
    $DnsScavenging | Export-Excel $ExcelPath -WorksheetName "DNS Scavenging" -AutoSize
    $ConditionalForwarders | Export-Excel $ExcelPath -WorksheetName "Conditional Forwarders" -AutoSize
    $DhcpScopes | Export-Excel $ExcelPath -WorksheetName "DHCP Scopes" -AutoSize
    $InactiveScopes | Export-Excel $ExcelPath -WorksheetName "Inactive Scopes" -AutoSize
    $DhcpReservations | Export-Excel $ExcelPath -WorksheetName "DHCP Reservations" -AutoSize
    $DhcpLeases | Export-Excel $ExcelPath -WorksheetName "DHCP Leases" -AutoSize

} catch {
    Write-ErrorLog "ImportExcel module missing."
}

#endregion

#region PROFESSIONAL HTML DASHBOARD

$HtmlPath = "$OutputFolder\AD_DNS_DHCP_Enterprise.html"

$WarningBanner = ""
if ($InactiveScopes.Count -gt 0) {
    $WarningBanner = "<div style='background:#f8d7da;color:#721c24;padding:15px;border-radius:5px;margin-bottom:15px;'>
    <b>WARNING:</b> $($InactiveScopes.Count) Inactive DHCP Scope(s) Detected
    </div>"
}

$HtmlHeader = @"
<html>
<head>
<title>AD DNS DHCP Documentation</title>
<style>
body {font-family:Segoe UI;margin:20px;background-color:#f4f6f9;}
h1 {color:#1f4e79;}
.card {background:white;padding:15px;margin:15px 0;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,.1);}
button {background:#1f4e79;color:white;padding:8px;border:none;cursor:pointer;margin-bottom:10px;}
.content {display:none;}
table {border-collapse:collapse;width:100%;}
th,td {border:1px solid #ccc;padding:6px;font-size:12px;}
th {background:#e9ecef;}
</style>
<script>
function toggle(id){
 var x=document.getElementById(id);
 x.style.display=(x.style.display==="none")?"block":"none";
}
function searchTable(){
 var input=document.getElementById("search").value.toLowerCase();
 var rows=document.getElementsByTagName("tr");
 for(var i=1;i<rows.length;i++){
  rows[i].style.display=rows[i].innerText.toLowerCase().includes(input)?"":"none";
 }
}
</script>
</head>
<body>

<h1>AD DNS DHCP Documentation</h1>
<h3>Created by Stephen McKee - Everi - IGT Server Administrator 2</h3>
$WarningBanner

<div class='card'>
<b>Executive Summary</b><br><br>
Total DNS Servers: $($ExecutiveSummary.TotalDNSServers)<br>
Total DNS Zones: $($ExecutiveSummary.TotalDNSZones)<br>
Conditional Forwarders: $($ExecutiveSummary.ConditionalForwarders)<br>
Total DHCP Servers: $($ExecutiveSummary.TotalDHCPServers)<br>
Total Scopes: $($ExecutiveSummary.TotalScopes)<br>
Inactive Scopes: $($ExecutiveSummary.InactiveScopes)<br>
</div>

<input type="text" id="search" onkeyup="searchTable()" placeholder="Search all results..." style="width:100%;padding:8px;margin-bottom:15px;">

"@

$Sections = @{
    "DNS Zones" = $DnsZones
    "DNS Scavenging" = $DnsScavenging
    "Conditional Forwarders" = $ConditionalForwarders
    "DHCP Scopes" = $DhcpScopes
    "Inactive DHCP Scopes" = $InactiveScopes
    "DHCP Reservations" = $DhcpReservations
    "DHCP Leases" = $DhcpLeases
}

$SectionID = 1
$HtmlBody = ""

foreach ($Section in $Sections.GetEnumerator()) {
    $HtmlBody += "<div class='card'>"
    $HtmlBody += "<button onclick='toggle(""sec$SectionID"")'>$($Section.Key)</button>"
    $HtmlBody += "<div class='content' id='sec$SectionID' style='display:none;'>"
    $HtmlBody += ($Section.Value | ConvertTo-Html -Fragment)
    $HtmlBody += "</div></div>"
    $SectionID++
}

$HtmlFooter = "</body></html>"

($HtmlHeader + $HtmlBody + $HtmlFooter) | Out-File $HtmlPath -Encoding UTF8

#endregion

Write-Host "`nEnterprise Documentation Complete." -ForegroundColor Green
Write-Host "Saved to: $OutputFolder"
