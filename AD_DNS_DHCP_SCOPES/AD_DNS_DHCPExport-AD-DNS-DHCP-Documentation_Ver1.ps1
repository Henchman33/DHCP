# ============================================================
# Script: Export-AD-DNS-DHCP-Documentation.ps1
# Author: Stephen McKee - Everi - IGT Server Administrator 2
# Purpose: Document all DNS and DHCP Servers and Scopes
# ============================================================

# ----------------------------
# VARIABLES & PATH SETUP
# ----------------------------

$DesktopPath = [Environment]::GetFolderPath("Desktop")
$OutputFolder = Join-Path $DesktopPath "AD_DNS_DHCP_Documentation"

# Create folder if not exists
if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

$ErrorLog = Join-Path $OutputFolder "AD_DNS_DHCP_ERRORS.log"

function Write-ErrorLog {
    param ($Message)
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $ErrorLog -Value "$TimeStamp - $Message"
}

# ----------------------------
# IMPORT REQUIRED MODULES
# ----------------------------

try { Import-Module ActiveDirectory -ErrorAction Stop } catch { Write-ErrorLog $_ }
try { Import-Module DnsServer -ErrorAction Stop } catch { Write-ErrorLog $_ }
try { Import-Module DhcpServer -ErrorAction Stop } catch { Write-ErrorLog $_ }

# ----------------------------
# DISCOVER DNS SERVERS
# ----------------------------

$DnsServers = @()
try {
    $DnsServers = Get-ADComputer -Filter {ServicePrincipalName -like "*DNS*"} |
                  Select-Object Name, DNSHostName
} catch { Write-ErrorLog $_ }

# ----------------------------
# DISCOVER DHCP SERVERS
# ----------------------------

$DhcpServers = @()
try {
    $DhcpServers = Get-DhcpServerInDC
} catch { Write-ErrorLog $_ }

# ----------------------------
# COLLECT DNS DATA
# ----------------------------

$DnsZones = @()
$DnsRecords = @()

foreach ($Server in $DnsServers) {
    try {
        $Zones = Get-DnsServerZone -ComputerName $Server.Name
        foreach ($Zone in $Zones) {
            $DnsZones += [PSCustomObject]@{
                Server = $Server.Name
                ZoneName = $Zone.ZoneName
                ZoneType = $Zone.ZoneType
                DynamicUpdate = $Zone.DynamicUpdate
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

# ----------------------------
# COLLECT DHCP DATA
# ----------------------------

$DhcpScopes = @()
$DhcpScopeOptions = @()

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

            $Options = Get-DhcpServerv4OptionValue -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId
            foreach ($Option in $Options) {
                $DhcpScopeOptions += [PSCustomObject]@{
                    Server = $Server.DnsName
                    ScopeID = $Scope.ScopeId
                    OptionID = $Option.OptionId
                    Name = $Option.Name
                    Value = ($Option.Value -join ", ")
                }
            }
        }
    } catch { Write-ErrorLog $_ }
}

# ----------------------------
# EXPORT CSV FILES
# ----------------------------

$DnsZones | Export-Csv "$OutputFolder\DNS_Zones.csv" -NoTypeInformation
$DnsRecords | Export-Csv "$OutputFolder\DNS_Records.csv" -NoTypeInformation
$DhcpScopes | Export-Csv "$OutputFolder\DHCP_Scopes.csv" -NoTypeInformation
$DhcpScopeOptions | Export-Csv "$OutputFolder\DHCP_Scope_Options.csv" -NoTypeInformation

# ----------------------------
# EXPORT XLSX (Requires ImportExcel Module)
# Install-Module ImportExcel if needed
# ----------------------------

try {
    Import-Module ImportExcel -ErrorAction Stop

    $ExcelPath = "$OutputFolder\AD_DNS_DHCP_Documentation.xlsx"

    $DnsZones | Export-Excel $ExcelPath -WorksheetName "DNS Zones" -AutoSize
    $DnsRecords | Export-Excel $ExcelPath -WorksheetName "DNS Records" -AutoSize
    $DhcpScopes | Export-Excel $ExcelPath -WorksheetName "DHCP Scopes" -AutoSize
    $DhcpScopeOptions | Export-Excel $ExcelPath -WorksheetName "DHCP Scope Options" -AutoSize

} catch {
    Write-ErrorLog "ImportExcel Module missing."
}

# ----------------------------
# GENERATE PROFESSIONAL HTML REPORT
# ----------------------------

$HtmlPath = "$OutputFolder\AD_DNS_DHCP_Documentation.html"

$HtmlHeader = @"
<html>
<head>
<title>AD DNS DHCP Documentation</title>
<style>
body {font-family: Arial; margin: 20px;}
h1 {color: #2E4053;}
.section {margin-bottom: 20px;}
button {background-color:#2E86C1;color:white;border:none;padding:10px;cursor:pointer;}
.content {display:none; padding:10px; border:1px solid #ccc;}
table {border-collapse: collapse; width:100%;}
th, td {border:1px solid #ccc; padding:5px;}
th {background-color:#f2f2f2;}
</style>
<script>
function toggle(id){
 var x=document.getElementById(id);
 if(x.style.display==="none"){x.style.display="block";}
 else{x.style.display="none";}
}
function searchTable(){
 var input=document.getElementById("search");
 var filter=input.value.toLowerCase();
 var tables=document.getElementsByTagName("table");
 for(var t=0;t<tables.length;t++){
  var tr=tables[t].getElementsByTagName("tr");
  for(var i=1;i<tr.length;i++){
    tr[i].style.display=tr[i].innerText.toLowerCase().includes(filter)?"":"none";
  }
 }
}
</script>
</head>
<body>

<h1>AD DNS DHCP Documentation</h1>
<h3>Created by Stephen McKee - Everi - IGT Server Administrator 2</h3>

<input type="text" id="search" onkeyup="searchTable()" placeholder="Search...">

"@

$HtmlBody = ""

$Sections = @{
    "DNS Zones" = $DnsZones
    "DNS Records" = $DnsRecords
    "DHCP Scopes" = $DhcpScopes
    "DHCP Scope Options" = $DhcpScopeOptions
}

$SectionID = 1

foreach ($Section in $Sections.GetEnumerator()) {
    $HtmlBody += "<div class='section'>"
    $HtmlBody += "<button onclick='toggle(""sec$SectionID"")'>$($Section.Key)</button>"
    $HtmlBody += "<div class='content' id='sec$SectionID' style='display:none;'>"
    $HtmlBody += ($Section.Value | ConvertTo-Html -Fragment)
    $HtmlBody += "</div></div>"
    $SectionID++
}

$HtmlFooter = "</body></html>"

$FullHtml = $HtmlHeader + $HtmlBody + $HtmlFooter
$FullHtml | Out-File $HtmlPath -Encoding UTF8

Write-Host "Documentation Complete. Files saved to $OutputFolder" -ForegroundColor Green
