# ============================================================
# AD DNS DHCP Ultimate Infrastructure Health Report
# Author: Stephen McKee - Everi - IGT Server Administrator 2
# Version: Ultimate Enterprise Edition (Printer Section Removed)
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

Import-Module ActiveDirectory
Import-Module DnsServer
Import-Module DhcpServer

#endregion

#region DISCOVERY

$DnsServers = Get-ADComputer -Filter {ServicePrincipalName -like "*DNS*"} -Properties DNSHostName |
              Select-Object Name

$DhcpServers = Get-DhcpServerInDC

#endregion

#region DHCP HEALTH ANALYSIS

$DhcpScopes = @()
$InactiveScopes = @()
$HighUtilization = @()
$NoLeases = @()
$NoReservations = @()
$DnsUpdatesDisabled = @()

$RiskScore = 0

foreach ($Server in $DhcpServers) {

    $Scopes = Get-DhcpServerv4Scope -ComputerName $Server.DnsName

    foreach ($Scope in $Scopes) {

        $Leases = Get-DhcpServerv4Lease -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId |
                  Where-Object {$_.AddressState -eq "Active"}

        $Reservations = Get-DhcpServerv4Reservation -ComputerName $Server.DnsName -ScopeId $Scope.ScopeId

        $TotalAddresses = ([ipaddress]$Scope.EndRange).Address - ([ipaddress]$Scope.StartRange).Address + 1
        $ActiveCount = $Leases.Count

        if ($TotalAddresses -gt 0) {
            $Utilization = [math]::Round(($ActiveCount / $TotalAddresses) * 100,2)
        } else {
            $Utilization = 0
        }

        $ScopeObject = [PSCustomObject]@{
            Server = $Server.DnsName
            ScopeName = $Scope.Name
            ScopeID = $Scope.ScopeId
            UtilizationPercent = $Utilization
            ActiveLeases = $ActiveCount
            TotalAddresses = $TotalAddresses
            State = $Scope.State
        }

        $DhcpScopes += $ScopeObject

        # Risk Conditions

        if ($Scope.State -ne "Active") {
            $InactiveScopes += $ScopeObject
            $RiskScore += 3
        }

        if ($Utilization -gt 90) {
            $HighUtilization += $ScopeObject
            $RiskScore += 3
        }

        if ($ActiveCount -eq 0) {
            $NoLeases += $ScopeObject
            $RiskScore += 2
        }

        if ($Reservations.Count -eq 0) {
            $NoReservations += $ScopeObject
            $RiskScore += 1
        }

        if ($Scope.DynamicUpdates -eq "Never") {
            $DnsUpdatesDisabled += $ScopeObject
            $RiskScore += 2
        }
    }
}

#endregion

#region RISK STATUS

if ($RiskScore -le 3) {
    $OverallStatus = "GREEN"
    $StatusColor = "#28a745"
}
elseif ($RiskScore -le 7) {
    $OverallStatus = "YELLOW"
    $StatusColor = "#ffc107"
}
else {
    $OverallStatus = "RED"
    $StatusColor = "#dc3545"
}

#endregion

#region EXPORT CSV

$DhcpScopes | Export-Csv "$OutputFolder\DHCP_Scope_Health.csv" -NoTypeInformation
$InactiveScopes | Export-Csv "$OutputFolder\DHCP_Inactive_Scopes.csv" -NoTypeInformation
$HighUtilization | Export-Csv "$OutputFolder\DHCP_High_Utilization.csv" -NoTypeInformation
$NoLeases | Export-Csv "$OutputFolder\DHCP_No_Active_Leases.csv" -NoTypeInformation
$NoReservations | Export-Csv "$OutputFolder\DHCP_No_Reservations.csv" -NoTypeInformation
$DnsUpdatesDisabled | Export-Csv "$OutputFolder\DHCP_DNS_Updates_Disabled.csv" -NoTypeInformation

#endregion

#region HTML DASHBOARD WITH UTILIZATION BARS

$HtmlPath = "$OutputFolder\AD_DNS_DHCP_Ultimate.html"

$Html = @"
<html>
<head>
<title>AD DNS DHCP Documentation</title>
<style>
body {font-family:Segoe UI;margin:20px;background:#f4f6f9;}
.card {background:white;padding:15px;margin:15px 0;border-radius:6px;box-shadow:0 2px 6px rgba(0,0,0,.1);}
.bar {height:20px;background:#ddd;border-radius:4px;}
.fill {height:100%;border-radius:4px;}
</style>
</head>
<body>

<h1>AD DNS DHCP Documentation</h1>
<h3>Created by Stephen McKee - Everi - IGT Server Administrator 2</h3>

<div class='card'>
<b>Overall Infrastructure Status:</b>
<span style='color:$StatusColor;font-weight:bold;'>$OverallStatus</span><br>
Risk Score: $RiskScore
</div>

<div class='card'>
<b>Scope Utilization Dashboard</b><br><br>
"@

foreach ($Scope in $DhcpScopes) {

    if ($Scope.UtilizationPercent -lt 80) {
        $Color = "#28a745"
    }
    elseif ($Scope.UtilizationPercent -lt 90) {
        $Color = "#ffc107"
    }
    else {
        $Color = "#dc3545"
    }

    $Html += "
    <b>$($Scope.ScopeName) ($($Scope.UtilizationPercent)%)</b>
    <div class='bar'>
        <div class='fill' style='width:$($Scope.UtilizationPercent)%;background:$Color;'></div>
    </div><br>
    "
}

$Html += "</div></body></html>"

$Html | Out-File $HtmlPath -Encoding UTF8

#endregion

#region EMAIL ALERT (Optional - Configure SMTP)

if ($OverallStatus -eq "RED") {

    $SmtpServer = "smtp.yourdomain.com"
    $To = "admin@yourdomain.com"
    $From = "report@yourdomain.com"

    Send-MailMessage -To $To -From $From -Subject "High Risk DHCP Infrastructure Detected" `
        -Body "Risk Score: $RiskScore. Review report immediately." `
        -SmtpServer $SmtpServer
}

#endregion

Write-Host "`nUltimate Infrastructure Health Report Complete." -ForegroundColor Green
Write-Host "Saved to: $OutputFolder"
