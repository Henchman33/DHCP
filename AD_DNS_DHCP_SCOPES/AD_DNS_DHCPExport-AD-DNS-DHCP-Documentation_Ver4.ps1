# Requires PowerShell - Version 5.1
<#
.SYNOPSIS
    AD DHCP DNS Report - Enterprise Documentation Script
    Created by: Stephen McKee - Server Administrator 2 - IGT - Everi

.DESCRIPTION
    Comprehensively documents all DHCP Scopes, IPv4 Addresses, Reservations, Leases,
    VLANs, DNS Zones, Records, and Active Directory DNS/DHCP server configuration
    across all Domain Controllers in the environment.

    Exports findings to:
      - CSV files (per data category)
      - XLSX workbook (multi-tab)
      - Enterprise-grade searchable HTML report

.NOTES
    Run from any Domain Controller or workstation with:
      - RSAT: DHCP Server Tools
      - RSAT: DNS Server Tools
      - RSAT: Active Directory Domain Services Tools
      - ImportExcel module (auto-installed if missing)
	  
	   Right-click > Run with PowerShell, OR in PowerShell ISE:
	   Make sure you're running as Administrator
Set-ExecutionPolicy -Scope process -ExecutionPolicy Bypass
.\AD_DHCP_DNS_Report.ps1
#>

#region --- INITIALIZATION ---
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"

$ScriptTitle   = "AD DHCP DNS Report"
$ScriptAuthor  = "Stephen McKee - Server Administrator 2 - IGT - Everi"
$RunDate       = Get-Date
$DateTimeStamp = $RunDate.ToString("yyyy-MM-dd_HH-mm-ss")
$DateDisplay   = $RunDate.ToString("MMMM dd, yyyy hh:mm:ss tt")

# Output folder on current user Desktop
$DesktopPath   = [Environment]::GetFolderPath("Desktop")
$OutputFolder  = Join-Path $DesktopPath "AD_DHCP_DNS_Report_$DateTimeStamp"
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  $ScriptTitle" -ForegroundColor Yellow
Write-Host "  $ScriptAuthor" -ForegroundColor Yellow
Write-Host "  Started: $DateDisplay" -ForegroundColor Gray
Write-Host "  Output:  $OutputFolder" -ForegroundColor Gray
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

# Check/Install ImportExcel
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[*] ImportExcel module not found. Attempting to install..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[+] ImportExcel installed successfully." -ForegroundColor Green
    } catch {
        Write-Warning "Could not install ImportExcel. XLSX export will be skipped. Error: $_"
        $SkipXLSX = $TRUE
    }
}
Import-Module ImportExcel -ErrorAction SilentlyContinue

#endregion

#region --- HELPER FUNCTIONS ---

function Write-Section {
    param([string]$Message)
    Write-Host ""
    Write-Host "[>>] $Message" -ForegroundColor Cyan
}

function Write-OK    { param([string]$m) Write-Host "  [+] $m" -ForegroundColor Green }
function Write-Warn  { param([string]$m) Write-Host "  [!] $m" -ForegroundColor Yellow }
function Write-Fail  { param([string]$m) Write-Host "  [-] $m" -ForegroundColor Red }
function Write-Info  { param([string]$m) Write-Host "  [i] $m" -ForegroundColor Gray }

function Safe-Export-CSV {
    param([object[]]$Data, [string]$Path, [string]$Label)
    if ($Data -and $Data.Count -gt 0) {
        $Data | Export-Csv -Path $Path -NoTypeInformation -Force
        Write-OK "$Label exported to CSV ($($Data.Count) records)"
    } else {
        Write-Warn "$Label - No data to export"
    }
}

#endregion

#region --- COLLECT AD DOMAIN INFO ---

Write-Section "Collecting Active Directory Domain Information"
$DomainInfo    = $null
$DomainControllers = @()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $DomainInfo = Get-ADDomain -ErrorAction Stop
    $ForestInfo = Get-ADForest -ErrorAction Stop
    $DomainControllers = Get-ADDomainController -filter * -ErrorAction Stop | Sort-Object HostName
    Write-OK "Domain: $($DomainInfo.DNSRoot) | Forest: $($ForestInfo.Name) | DCs found: $($DomainControllers.Count)"
} catch {
    Write-Fail "Could not retrieve AD Domain info: $_"
    $DomainControllers = @([PSCustomObject]@{ HostName = $env:COMPUTERNAME; Site = "Unknown"; IPv4Address = "Unknown"; IsGlobalCatalog = $FALSE; OperationMasterRoles = @() })
}

$ADSummaryData = if ($DomainInfo) {
    [PSCustomObject]@{
        DomainName          = $DomainInfo.DNSRoot
        NetBIOSName         = $DomainInfo.NetBIOSName
        ForestName          = $ForestInfo.Name
        ForestMode          = $ForestInfo.ForestMode
        DomainMode          = $DomainInfo.DomainMode
        PDCEmulator         = $DomainInfo.PDCEmulator
        RIDMaster           = $DomainInfo.RIDMaster
        InfrastructureMaster= $DomainInfo.InfrastructureMaster
        SchemaMaster        = $ForestInfo.SchemaMaster
        DomainNamingMaster  = $ForestInfo.DomainNamingMaster
        DomainControllers   = ($DomainControllers.HostName -join "; ")
        TotalDCs            = $DomainControllers.Count
        ReportGeneratedBy   = $env:USERNAME
        ReportDate          = $DateDisplay
    }
} else { @() }

$DCDetailData = $DomainControllers | foreach-Object {
    [PSCustomObject]@{
        Hostname            = $_.HostName
        Site                = $_.Site
        IPv4Address         = $_.IPv4Address
        IsGlobalCatalog     = $_.IsGlobalCatalog
        OperationMasterRoles= ($_.OperationMasterRoles -join "; ")
        IsReadOnly          = $_.IsReadOnly
        OperatingSystem     = $_.OperatingSystem
        OperatingSystemVersion = $_.OperatingSystemVersion
    }
}

Safe-Export-CSV -Data @($ADSummaryData) -Path "$OutputFolder\01_AD_Summary.csv" -Label "AD Summary"
Safe-Export-CSV -Data $DCDetailData -Path "$OutputFolder\02_DomainControllers.csv" -Label "Domain Controllers"

#endregion

#region --- COLLECT DHCP SERVER LIST ---

Write-Section "Discovering DHCP Servers"

$DHCPServers = @()
try {
    $DHCPServers = Get-DhcpServerInDC -ErrorAction Stop | Sort-Object DnsName
    Write-OK "DHCP Servers found in AD: $($DHCPServers.Count)"
} catch {
    Write-Warn "Could not query DHCP servers from AD. Falling back to Domain Controllers list."
    $DHCPServers = $DomainControllers | Select-Object @{N="DnsName";E={$_.HostName}}, @{N="IPAddress";E={$_.IPv4Address}}
}

$DHCPServerData = $DHCPServers | foreach-Object {
    $srv = $_
    $stats = $null
    try { $stats = Get-DhcpServerSetting -ComputerName $srv.DnsName -ErrorAction Stop } catch {}
    [PSCustomObject]@{
        ServerName          = $srv.DnsName
        IPAddress           = $srv.IPAddress
        IsAuthorized        = $TRUE
        DynamicDNSEnabled   = if ($stats) { $stats.DynamicDnsQueueLength } else { "N/A" }
        ConflictDetectionAttempts = if ($stats) { $stats.ConflictDetectionAttempts } else { "N/A" }
        NapEnabled          = if ($stats) { $stats.NapEnabled } else { "N/A" }
    }
}
Safe-Export-CSV -Data $DHCPServerData -Path "$OutputFolder\03_DHCP_Servers.csv" -Label "DHCP Servers"

#endregion

#region --- COLLECT DHCP SCOPES, LEASES, RESERVATIONS ---

Write-Section "Collecting DHCP Scopes, Leases, and Reservations"

$AllScopes       = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllLeases       = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllReservations = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllExclusions   = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllScopeOptions = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllServerOptions= [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($server in $DHCPServers) {
    $srv = $server.DnsName
    Write-Info "Processing DHCP Server: $srv"

    # Server-level options
    try {
        $srvOptions = Get-DhcpServerv4OptionValue -ComputerName $srv -ErrorAction Stop
        foreach ($opt in $srvOptions) {
            $AllServerOptions.Add([PSCustomObject]@{
                Server      = $srv
                OptionId    = $opt.OptionId
                Name        = $opt.Name
                Type        = $opt.Type
                Value       = ($opt.Value -join "; ")
                VendorClass = $opt.VendorClass
                PolicyName  = $opt.PolicyName
            })
        }
    } catch { Write-Warn "  Server options failed on $srv : $_" }

    # Scopes
    try {
        $scopes = Get-DhcpServerv4Scope -ComputerName $srv -ErrorAction Stop
        Write-Info "  Scopes found: $($scopes.Count)"

        foreach ($scope in $scopes) {
            # Try to infer VLAN from scope name or description
            $vlanGuess = "N/A"
            if ($scope.Name -match "(?i)vlan\s*(\d+)") { $vlanGuess = "VLAN $($Matches[1])" }
            elseif ($scope.Description -match "(?i)vlan\s*(\d+)") { $vlanGuess = "VLAN $($Matches[1])" }

            # Scope statistics
            $stats = $null
            try { $stats = Get-DhcpServerv4ScopeStatistics -ComputerName $srv -ScopeId $scope.ScopeId -ErrorAction Stop } catch {}

            $AllScopes.Add([PSCustomObject]@{
                Server             = $srv
                ScopeId            = $scope.ScopeId
                Name               = $scope.Name
                Description        = $scope.Description
                SubnetMask         = $scope.SubnetMask
                StartRange         = $scope.StartRange
                EndRange           = $scope.EndRange
                State              = $scope.State
                LeaseDuration      = $scope.LeaseDuration
                NapEnable          = $scope.NapEnable
                NapProfile         = $scope.NapProfile
                SuperscopeName     = $scope.SuperscopeName
                MaxBootpClients    = $scope.MaxBootpClients
                ActivatePolicies   = $scope.ActivatePolicies
                VLAN_Inferred      = $vlanGuess
                TotalAddresses     = if ($stats) { $stats.AddressesFree + $stats.AddressesInUse } else { "N/A" }
                InUse              = if ($stats) { $stats.AddressesInUse } else { "N/A" }
                Free               = if ($stats) { $stats.AddressesFree } else { "N/A" }
                PercentInUse       = if ($stats) { "$([math]::Round($stats.PercentageInUse,2))%" } else { "N/A" }
                Reserved           = if ($stats) { $stats.ReservedAddress } else { "N/A" }
                Pending            = if ($stats) { $stats.PendingOffers } else { "N/A" }
            })

            # Leases
            try {
                $leases = Get-DhcpServerv4Lease -ComputerName $srv -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($lease in $leases) {
                    $AllLeases.Add([PSCustomObject]@{
                        Server        = $srv
                        ScopeId       = $scope.ScopeId
                        ScopeName     = $scope.Name
                        IPAddress     = $lease.IPAddress
                        ClientId      = $lease.ClientId
                        Hostname      = $lease.HostName
                        AddressState  = $lease.AddressState
                        LeaseExpiryTime = $lease.LeaseExpiryTime
                        ClientType    = $lease.ClientType
                        Description   = $lease.Description
                        NapCapable    = $lease.NapCapable
                        NapStatus     = $lease.NapStatus
                        ProbationEnds = $lease.ProbationEnds
                        DNS_RR_Name   = $lease.DnsRR
                        DNS_Registration = $lease.DnsRegistration
                        ServerIP      = $lease.ServerIPAddress
                        VLAN_Inferred = $vlanGuess
                    })
                }
            } catch { Write-Warn "  Leases failed for scope $($scope.ScopeId) on $srv : $_" }

            # Reservations
            try {
                $reservations = Get-DhcpServerv4Reservation -ComputerName $srv -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($res in $reservations) {
                    $AllReservations.Add([PSCustomObject]@{
                        Server        = $srv
                        ScopeId       = $scope.ScopeId
                        ScopeName     = $scope.Name
                        IPAddress     = $res.IPAddress
                        ClientId      = $res.ClientId
                        Name          = $res.Name
                        Description   = $res.Description
                        Type          = $res.Type
                        VLAN_Inferred = $vlanGuess
                    })
                }
            } catch { Write-Warn "  Reservations failed for scope $($scope.ScopeId) on $srv : $_" }

            # Exclusions
            try {
                $exclusions = Get-DhcpServerv4ExclusionRange -ComputerName $srv -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($ex in $exclusions) {
                    $AllExclusions.Add([PSCustomObject]@{
                        Server    = $srv
                        ScopeId   = $scope.ScopeId
                        ScopeName = $scope.Name
                        StartRange= $ex.StartRange
                        EndRange  = $ex.EndRange
                    })
                }
            } catch {}

            # Scope Options
            try {
                $scopeOpts = Get-DhcpServerv4OptionValue -ComputerName $srv -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($opt in $scopeOpts) {
                    $AllScopeOptions.Add([PSCustomObject]@{
                        Server      = $srv
                        ScopeId     = $scope.ScopeId
                        ScopeName   = $scope.Name
                        OptionId    = $opt.OptionId
                        Name        = $opt.Name
                        Type        = $opt.Type
                        Value       = ($opt.Value -join "; ")
                        VendorClass = $opt.VendorClass
                        PolicyName  = $opt.PolicyName
                    })
                }
            } catch {}
        }
    } catch { Write-Fail "  Could not retrieve scopes from $srv : $_" }
}

Safe-Export-CSV -Data $AllScopes.ToArray()       -Path "$OutputFolder\04_DHCP_Scopes.csv"       -Label "DHCP Scopes"
Safe-Export-CSV -Data $AllLeases.ToArray()       -Path "$OutputFolder\05_DHCP_Leases.csv"       -Label "DHCP Leases"
Safe-Export-CSV -Data $AllReservations.ToArray() -Path "$OutputFolder\06_DHCP_Reservations.csv" -Label "DHCP Reservations"
Safe-Export-CSV -Data $AllExclusions.ToArray()   -Path "$OutputFolder\07_DHCP_Exclusions.csv"   -Label "DHCP Exclusions"
Safe-Export-CSV -Data $AllScopeOptions.ToArray() -Path "$OutputFolder\08_DHCP_ScopeOptions.csv" -Label "DHCP Scope Options"
Safe-Export-CSV -Data $AllServerOptions.ToArray()-Path "$OutputFolder\09_DHCP_ServerOptions.csv"-Label "DHCP Server Options"

#endregion

#region --- COLLECT DNS DATA ---

Write-Section "Collecting DNS Zones and Records"

$AllDNSZones   = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDNSRecords = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDNSForwarders = [System.Collections.Generic.List[PSCustomObject]]::new()

# Use all DCs for DNS collection
$DNSServers = $DomainControllers | Select-Object -ExpandProperty HostName

foreach ($dnsServer in $DNSServers) {
    Write-Info "Processing DNS Server: $dnsServer"

    # Forwarders
    try {
        $fwd = Get-DnsServerForwarder -ComputerName $dnsServer -ErrorAction Stop
        foreach ($f in $fwd.IPAddress) {
            $AllDNSForwarders.Add([PSCustomObject]@{
                Server          = $dnsServer
                ForwarderIP     = $f
                UseRootHint     = $fwd.UseRootHint
                Timeout         = $fwd.Timeout
                EnableReordering= $fwd.EnableReordering
            })
        }
    } catch { Write-Warn "  Forwarders failed on $dnsServer : $_" }

    # Zones
    try {
        $zones = Get-DnsServerZone -ComputerName $dnsServer -ErrorAction Stop
        Write-Info "  DNS Zones found: $($zones.Count)"

        foreach ($zone in $zones) {
            $AllDNSZones.Add([PSCustomObject]@{
                Server              = $dnsServer
                ZoneName            = $zone.ZoneName
                ZoneType            = $zone.ZoneType
                IsDsIntegrated      = $zone.IsDsIntegrated
                IsReverseLookupZone = $zone.IsReverseLookupZone
                IsAutoCreated       = $zone.IsAutoCreated
                IsPaused            = $zone.IsPaused
                IsShutdown          = $zone.IsShutdown
                IsReadOnly          = $zone.IsReadOnly
                DynamicUpdate       = $zone.DynamicUpdate
                ReplicationScope    = $zone.ReplicationScope
                DirectoryPartitionName = $zone.DirectoryPartitionName
                MasterServers       = ($zone.MasterServers -join "; ")
                NotifyServers       = ($zone.NotifyServers -join "; ")
                SecureSecondaries   = $zone.SecureSecondaries
                SecondaryServers    = ($zone.SecondaryServers -join "; ")
            })

            # Records (only for primary/primary-ds zones to avoid huge duplicates across DCs)
            if ($zone.ZoneType -in @("Primary","Stub") -and $dnsServer -eq $DNSServers[0]) {
                try {
                    $records = Get-DnsServerResourceRecord -ComputerName $dnsServer -ZoneName $zone.ZoneName -ErrorAction Stop
                    foreach ($rec in $records) {
                        $recData = ""
                        try {
                            switch ($rec.RecordType) {
                                "A"     { $recData = $rec.RecordData.IPv4Address.IPAddressToString }
                                "AAAA"  { $recData = $rec.RecordData.IPv6Address.IPAddressToString }
                                "CNAME" { $recData = $rec.RecordData.HostNameAlias }
                                "MX"    { $recData = "$($rec.RecordData.MailExchange) Pref=$($rec.RecordData.Preference)" }
                                "NS"    { $recData = $rec.RecordData.NameServer }
                                "PTR"   { $recData = $rec.RecordData.PtrDomainName }
                                "SOA"   { $recData = "NS=$($rec.RecordData.PrimaryServer) Serial=$($rec.RecordData.SerialNumber)" }
                                "SRV"   { $recData = "$($rec.RecordData.DomainName):$($rec.RecordData.Port) Pri=$($rec.RecordData.Priority) Wt=$($rec.RecordData.Weight)" }
                                "TXT"   { $recData = ($rec.RecordData.DescriptiveText -join " ") }
                                default { $recData = $rec.RecordData | Out-String -Width 200 | foreach-Object { $_.Trim() } | Select-Object -First 1 }
                            }
                        } catch { $recData = "ParseError" }

                        $AllDNSRecords.Add([PSCustomObject]@{
                            Zone        = $zone.ZoneName
                            Name        = $rec.HostName
                            RecordType  = $rec.RecordType
                            TTL         = $rec.TimeToLive
                            Data        = $recData
                            Timestamp   = $rec.Timestamp
                            AgeingEnabled = $rec.TimeStamp -ne $null
                            Server      = $dnsServer
                        })
                    }
                } catch { Write-Warn "  Records failed for zone $($zone.ZoneName) on $dnsServer : $_" }
            }
        }
    } catch { Write-Fail "  DNS Zones failed on $dnsServer : $_" }
}

Safe-Export-CSV -Data $AllDNSZones.ToArray()      -Path "$OutputFolder\10_DNS_Zones.csv"      -Label "DNS Zones"
Safe-Export-CSV -Data $AllDNSRecords.ToArray()    -Path "$OutputFolder\11_DNS_Records.csv"    -Label "DNS Records"
Safe-Export-CSV -Data $AllDNSForwarders.ToArray() -Path "$OutputFolder\12_DNS_Forwarders.csv" -Label "DNS Forwarders"

#endregion

#region --- EXPORT XLSX ---

if (-not $SkipXLSX) {
    Write-Section "Exporting Excel Workbook (.xlsx)"
    $XLSXPath = "$OutputFolder\AD_DHCP_DNS_Report_$DateTimeStamp.xlsx"

    $xlParams = @{ Path = $XLSXPath; AutoSize = $TRUE; FreezeTopRow = $TRUE; BoldTopRow = $TRUE; TableStyle = "Medium9" }

    function Add-Sheet {
        param($Data, [string]$SheetName)
        if ($Data -and @($Data).Count -gt 0) {
            @($Data) | Export-Excel @xlParams -WorksheetName $SheetName -Append
            Write-OK "Sheet '$SheetName' added ($(@($Data).Count) rows)"
        } else {
            Write-Warn "Sheet '$SheetName' skipped - no data"
        }
    }

    Add-Sheet -Data @($ADSummaryData)            -SheetName "AD_Summary"
    Add-Sheet -Data $DCDetailData                -SheetName "Domain_Controllers"
    Add-Sheet -Data $DHCPServerData              -SheetName "DHCP_Servers"
    Add-Sheet -Data $AllScopes.ToArray()         -SheetName "DHCP_Scopes"
    Add-Sheet -Data $AllLeases.ToArray()         -SheetName "DHCP_Leases"
    Add-Sheet -Data $AllReservations.ToArray()   -SheetName "DHCP_Reservations"
    Add-Sheet -Data $AllExclusions.ToArray()     -SheetName "DHCP_Exclusions"
    Add-Sheet -Data $AllScopeOptions.ToArray()   -SheetName "DHCP_ScopeOptions"
    Add-Sheet -Data $AllServerOptions.ToArray()  -SheetName "DHCP_ServerOptions"
    Add-Sheet -Data $AllDNSZones.ToArray()       -SheetName "DNS_Zones"
    Add-Sheet -Data $AllDNSRecords.ToArray()     -SheetName "DNS_Records"
    Add-Sheet -Data $AllDNSForwarders.ToArray()  -SheetName "DNS_Forwarders"

    Write-OK "Excel workbook saved: $XLSXPath"
} else {
    Write-Warn "XLSX export skipped (ImportExcel not available)"
}

#endregion

#region --- BUILD HTML REPORT ---

Write-Section "Building Enterprise HTML Report"

function ConvertTo-HtmlTable {
    param(
        [object[]]$Data,
        [string]$TableId,
        [string]$Caption
    )
    if (-not $Data -or $Data.Count -eq 0) {
        return "<p class='no-data'>No data available for this section.</p>"
    }
    $headers = $Data[0].PSObject.Properties.Name
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<div class='table-wrapper'>")
    [void]$sb.Append("<input type='text' class='search-box' placeholder='Search $Caption...' onkeyup=""filterTable(this, '$TableId')"" />")
    [void]$sb.Append("<div class='table-scroll'><table id='$TableId' class='data-table'>")
    [void]$sb.Append("<thead><tr>")
    foreach ($h in $headers) { [void]$sb.Append("<th onclick=""sortTable('$TableId', $([array]::IndexOf($headers,$h)))"">$h <span class='sort-icon'>⇅</span></th>") }
    [void]$sb.Append("</tr></thead><tbody>")
    foreach ($row in $Data) {
        [void]$sb.Append("<tr>")
        foreach ($h in $headers) {
            $val = $row.$h
            if ($null -eq $val) { $val = "" }
            $valStr = [System.Web.HttpUtility]::HtmlEncode($val.ToString())
            # Color-code scope state
            $tdClass = ""
            if ($h -eq "State") {
                if ($val -eq "Active") { $tdClass = " class='state-active'" }
                elseif ($val -eq "InActive") { $tdClass = " class='state-inactive'" }
            }
            if ($h -eq "AddressState") {
                if ($val -like "*Active*") { $tdClass = " class='state-active'" }
                elseif ($val -like "*Expired*") { $tdClass = " class='state-expired'" }
            }
            if ($h -eq "PercentInUse") {
                $pct = [double]($val -replace '%','') 
                if ($pct -ge 90) { $tdClass = " class='pct-critical'" }
                elseif ($pct -ge 70) { $tdClass = " class='pct-warn'" }
                elseif ($pct -gt 0)  { $tdClass = " class='pct-ok'" }
            }
            [void]$sb.Append("<td$tdClass>$valStr</td>")
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table></div></div>")
    return $sb.ToString()
}

# Build summary cards
$totalScopes       = $AllScopes.Count
$activeScopes      = ($AllScopes | Where-Object { $_.State -eq "Active" }).Count
$totalLeases       = $AllLeases.Count
$totalReservations = $AllReservations.Count
$totalDNSZones     = ($AllDNSZones | Select-Object -Property ZoneName,IsReverseLookupZone -Unique).Count
$totalDNSRecords   = $AllDNSRecords.Count
$totalDHCPServers  = $DHCPServers.Count
$totalDCs          = $DomainControllers.Count

$HTMLPath = "$OutputFolder\AD_DHCP_DNS_Report_$DateTimeStamp.html"

# Load System.Web for HtmlEncode
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue

$navLinks = @(
    @{id="sec-ad-summary";   label="AD Summary"},
    @{id="sec-dcs";          label="Domain Controllers"},
    @{id="sec-dhcp-servers"; label="DHCP Servers"},
    @{id="sec-scopes";       label="DHCP Scopes"},
    @{id="sec-leases";       label="DHCP Leases"},
    @{id="sec-reservations"; label="DHCP Reservations"},
    @{id="sec-exclusions";   label="Exclusions"},
    @{id="sec-scope-opts";   label="Scope Options"},
    @{id="sec-srv-opts";     label="Server Options"},
    @{id="sec-dns-zones";    label="DNS Zones"},
    @{id="sec-dns-records";  label="DNS Records"},
    @{id="sec-dns-fwd";      label="DNS Forwarders"}
)

$navHtml = ($navLinks | foreach-Object { "<a href='#$($_.id)'>$($_.label)</a>" }) -join "`n"

$tAD_Summary     = ConvertTo-HtmlTable -Data @($ADSummaryData)            -TableId "tbl_adsum"    -Caption "AD Summary"
$tDCs            = ConvertTo-HtmlTable -Data $DCDetailData                -TableId "tbl_dcs"      -Caption "Domain Controllers"
$tDHCPSrv        = ConvertTo-HtmlTable -Data $DHCPServerData              -TableId "tbl_dhcpsrv"  -Caption "DHCP Servers"
$tScopes         = ConvertTo-HtmlTable -Data $AllScopes.ToArray()         -TableId "tbl_scopes"   -Caption "DHCP Scopes"
$tLeases         = ConvertTo-HtmlTable -Data $AllLeases.ToArray()         -TableId "tbl_leases"   -Caption "DHCP Leases"
$tReservations   = ConvertTo-HtmlTable -Data $AllReservations.ToArray()   -TableId "tbl_res"      -Caption "DHCP Reservations"
$tExclusions     = ConvertTo-HtmlTable -Data $AllExclusions.ToArray()     -TableId "tbl_excl"     -Caption "DHCP Exclusions"
$tScopeOpts      = ConvertTo-HtmlTable -Data $AllScopeOptions.ToArray()   -TableId "tbl_scopeopts"-Caption "DHCP Scope Options"
$tSrvOpts        = ConvertTo-HtmlTable -Data $AllServerOptions.ToArray()  -TableId "tbl_srvopts"  -Caption "DHCP Server Options"
$tDNSZones       = ConvertTo-HtmlTable -Data $AllDNSZones.ToArray()       -TableId "tbl_dnszones" -Caption "DNS Zones"
$tDNSRecords     = ConvertTo-HtmlTable -Data $AllDNSRecords.ToArray()     -TableId "tbl_dnsrec"   -Caption "DNS Records"
$tDNSFwd         = ConvertTo-HtmlTable -Data $AllDNSForwarders.ToArray()  -TableId "tbl_dnsfwd"   -Caption "DNS Forwarders"

$HTML = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8" />
<meta name="viewport" content="width=device-width, initial-scale=1.0"/>
<title>AD DHCP DNS Report - Stephen McKee</title>
<style>
  :root {
    --bg: #0d1117; --surface: #161b22; --surface2: #1c2128; --border: #30363d;
    --accent: #2563eb; --accent2: #1d4ed8; --text: #e6edf3; --muted: #7d8590;
    --green: #3fb950; --yellow: #d29922; --red: #f85149; --purple: #a371f7;
    --orange: #db6d28; --cyan: #39d3f2;
    --header-bg: #0d1117; --nav-bg: #161b22;
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  html { scroll-behavior: smooth; }
  body { background: var(--bg); color: var(--text); font-family: 'Segoe UI', system-ui, sans-serif; font-size: 13px; }

  /* TOP HEADER */
  .top-header {
    background: linear-gradient(135deg, #0d1117 0%, #161b22 50%, #0d1117 100%);
    border-bottom: 2px solid var(--accent);
    padding: 24px 32px 20px;
    position: sticky; top: 0; z-index: 100;
    display: flex; justify-content: space-between; align-items: flex-end;
  }
  .top-header .title-block h1 {
    font-size: 22px; font-weight: 700; color: #fff; letter-spacing: 0.5px;
  }
  .top-header .title-block h1 span { color: var(--accent); }
  .top-header .title-block .subtitle {
    font-size: 11px; color: var(--muted); margin-top: 3px; letter-spacing: 0.3px;
  }
  .top-header .meta { text-align: right; font-size: 11px; color: var(--muted); line-height: 1.7; }
  .top-header .meta strong { color: var(--cyan); }

  /* NAV */
  .sidenav {
    position: fixed; top: 0; left: 0; width: 210px; height: 100vh;
    background: var(--nav-bg); border-right: 1px solid var(--border);
    overflow-y: auto; padding-top: 90px; z-index: 90;
  }
  .sidenav .nav-label { font-size: 10px; color: var(--muted); text-transform: uppercase;
    letter-spacing: 1px; padding: 14px 16px 6px; }
  .sidenav a {
    display: block; padding: 7px 16px; color: var(--muted); text-decoration: none;
    font-size: 12px; border-left: 3px solid transparent; transition: all 0.15s;
  }
  .sidenav a:hover { color: var(--text); background: var(--surface2); border-left-color: var(--accent); }

  /* MAIN CONTENT */
  .main { margin-left: 210px; padding: 24px 28px; }

  /* GLOBAL SEARCH */
  .global-search-bar {
    background: var(--surface); border: 1px solid var(--border); border-radius: 8px;
    padding: 14px 18px; margin-bottom: 20px; display: flex; align-items: center; gap: 12px;
  }
  .global-search-bar input {
    background: var(--surface2); border: 1px solid var(--border); color: var(--text);
    padding: 8px 14px; border-radius: 6px; font-size: 13px; width: 380px;
    outline: none; transition: border-color 0.2s;
  }
  .global-search-bar input:focus { border-color: var(--accent); }
  .global-search-bar label { color: var(--muted); font-size: 12px; }
  .global-match-count { font-size: 12px; color: var(--cyan); min-width: 100px; }

  /* SUMMARY CARDS */
  .cards-grid { display: grid; grid-template-columns: repeat(auto-fill, minmax(155px,1fr)); gap: 14px; margin-bottom: 24px; }
  .card {
    background: var(--surface); border: 1px solid var(--border); border-radius: 10px;
    padding: 16px; position: relative; overflow: hidden; cursor: default;
  }
  .card::before { content:''; position:absolute; top:0; left:0; right:0; height:3px; }
  .card.blue::before { background: var(--accent); }
  .card.green::before { background: var(--green); }
  .card.yellow::before { background: var(--yellow); }
  .card.red::before { background: var(--red); }
  .card.purple::before { background: var(--purple); }
  .card.cyan::before { background: var(--cyan); }
  .card.orange::before { background: var(--orange); }
  .card .card-val { font-size: 32px; font-weight: 700; line-height: 1.1; color: #fff; }
  .card .card-lbl { font-size: 11px; color: var(--muted); margin-top: 4px; }

  /* SECTIONS */
  .section { margin-bottom: 32px; }
  .section-header {
    display: flex; align-items: center; gap: 10px;
    border-bottom: 1px solid var(--border); padding-bottom: 10px; margin-bottom: 14px;
  }
  .section-header h2 { font-size: 16px; font-weight: 600; color: #fff; }
  .section-badge {
    background: var(--accent); color: #fff; font-size: 10px; font-weight: 700;
    padding: 2px 8px; border-radius: 20px;
  }

  /* TABLE */
  .table-wrapper { position: relative; }
  .search-box {
    background: var(--surface2); border: 1px solid var(--border); color: var(--text);
    padding: 6px 12px; border-radius: 5px; font-size: 12px; margin-bottom: 8px;
    min-width: 280px; outline: none;
  }
  .search-box:focus { border-color: var(--accent); }
  .table-scroll { overflow-x: auto; border-radius: 8px; border: 1px solid var(--border); }
  .data-table { width: 100%; border-collapse: collapse; white-space: nowrap; }
  .data-table thead { background: var(--surface2); position: sticky; top: 0; }
  .data-table th {
    padding: 9px 12px; text-align: left; font-size: 11px; font-weight: 600;
    color: var(--muted); text-transform: uppercase; letter-spacing: 0.5px;
    border-bottom: 1px solid var(--border); cursor: pointer; user-select: none;
  }
  .data-table th:hover { color: var(--text); }
  .sort-icon { opacity: 0.4; font-size: 10px; }
  .data-table td {
    padding: 7px 12px; border-bottom: 1px solid #21262d; font-size: 12px;
    max-width: 320px; overflow: hidden; text-overflow: ellipsis;
  }
  .data-table tbody tr:hover { background: var(--surface2); }
  .data-table tbody tr:last-child td { border-bottom: none; }
  .data-table tr.hidden { display: none; }

  /* State badges */
  td.state-active  { color: var(--green); font-weight: 600; }
  td.state-inactive{ color: var(--muted); }
  td.state-expired { color: var(--red); }
  td.pct-critical  { color: var(--red); font-weight: 700; }
  td.pct-warn      { color: var(--yellow); font-weight: 600; }
  td.pct-ok        { color: var(--green); }
  .no-data { color: var(--muted); font-style: italic; padding: 12px 0; }

  /* Footer */
  .footer {
    margin-left: 210px; padding: 16px 28px; border-top: 1px solid var(--border);
    color: var(--muted); font-size: 11px; display: flex; justify-content: space-between;
  }

  /* Scrollbar */
  ::-webkit-scrollbar { width: 6px; height: 6px; }
  ::-webkit-scrollbar-track { background: var(--bg); }
  ::-webkit-scrollbar-thumb { background: var(--border); border-radius: 3px; }
  ::-webkit-scrollbar-thumb:hover { background: var(--muted); }

  @media print {
    .sidenav, .global-search-bar, .search-box, .top-header { display: none !important; }
    .main, .footer { margin-left: 0 !important; }
  }
</style>
</head>
<body>

<header class="top-header">
  <div class="title-block">
    <h1><span>AD</span> DHCP DNS Report</h1>
    <div class="subtitle">Created by $ScriptAuthor</div>
  </div>
  <div class="meta">
    <div>Generated: <strong>$DateDisplay</strong></div>
    <div>Generated By: <strong>$($env:USERNAME)</strong></div>
    <div>Domain: <strong>$(if($DomainInfo){$DomainInfo.DNSRoot}else{"N/A"})</strong></div>
  </div>
</header>

<nav class="sidenav">
  <div class="nav-label">Navigation</div>
  $navHtml
</nav>

<main class="main">

  <!-- GLOBAL SEARCH -->
  <div class="global-search-bar">
    <label>&#128269; Global Search:</label>
    <input type="text" id="globalSearch" placeholder="Search all tables..." oninput="globalFilter()" />
    <span class="global-match-count" id="globalMatchCount"></span>
  </div>

  <!-- SUMMARY CARDS -->
  <div class="cards-grid">
    <div class="card blue"><div class="card-val">$totalDCs</div><div class="card-lbl">Domain Controllers</div></div>
    <div class="card green"><div class="card-val">$totalDHCPServers</div><div class="card-lbl">DHCP Servers</div></div>
    <div class="card cyan"><div class="card-val">$totalScopes</div><div class="card-lbl">DHCP Scopes</div></div>
    <div class="card green"><div class="card-val">$activeScopes</div><div class="card-lbl">Active Scopes</div></div>
    <div class="card yellow"><div class="card-val">$totalLeases</div><div class="card-lbl">DHCP Leases</div></div>
    <div class="card orange"><div class="card-val">$totalReservations</div><div class="card-lbl">Reservations</div></div>
    <div class="card purple"><div class="card-val">$totalDNSZones</div><div class="card-lbl">DNS Zones</div></div>
    <div class="card blue"><div class="card-val">$totalDNSRecords</div><div class="card-lbl">DNS Records</div></div>
  </div>

  <!-- AD SUMMARY -->
  <div class="section" id="sec-ad-summary">
    <div class="section-header"><h2>Active Directory Domain Summary</h2><span class="section-badge">AD</span></div>
    $tAD_Summary
  </div>

  <!-- DOMAIN CONTROLLERS -->
  <div class="section" id="sec-dcs">
    <div class="section-header"><h2>Domain Controllers</h2><span class="section-badge">$totalDCs</span></div>
    $tDCs
  </div>

  <!-- DHCP SERVERS -->
  <div class="section" id="sec-dhcp-servers">
    <div class="section-header"><h2>DHCP Servers</h2><span class="section-badge">$totalDHCPServers</span></div>
    $tDHCPSrv
  </div>

  <!-- DHCP SCOPES -->
  <div class="section" id="sec-scopes">
    <div class="section-header"><h2>DHCP Scopes (IPv4)</h2><span class="section-badge">$totalScopes</span></div>
    $tScopes
  </div>

  <!-- DHCP LEASES -->
  <div class="section" id="sec-leases">
    <div class="section-header"><h2>DHCP Leases (Active &amp; Expired)</h2><span class="section-badge">$totalLeases</span></div>
    $tLeases
  </div>

  <!-- DHCP RESERVATIONS -->
  <div class="section" id="sec-reservations">
    <div class="section-header"><h2>DHCP Reservations</h2><span class="section-badge">$totalReservations</span></div>
    $tReservations
  </div>

  <!-- DHCP EXCLUSIONS -->
  <div class="section" id="sec-exclusions">
    <div class="section-header"><h2>DHCP Exclusion Ranges</h2></div>
    $tExclusions
  </div>

  <!-- SCOPE OPTIONS -->
  <div class="section" id="sec-scope-opts">
    <div class="section-header"><h2>DHCP Scope Options</h2></div>
    $tScopeOpts
  </div>

  <!-- SERVER OPTIONS -->
  <div class="section" id="sec-srv-opts">
    <div class="section-header"><h2>DHCP Server Options</h2></div>
    $tSrvOpts
  </div>

  <!-- DNS ZONES -->
  <div class="section" id="sec-dns-zones">
    <div class="section-header"><h2>DNS Zones</h2><span class="section-badge">$totalDNSZones</span></div>
    $tDNSZones
  </div>

  <!-- DNS RECORDS -->
  <div class="section" id="sec-dns-records">
    <div class="section-header"><h2>DNS Resource Records</h2><span class="section-badge">$totalDNSRecords</span></div>
    $tDNSRecords
  </div>

  <!-- DNS FORWARDERS -->
  <div class="section" id="sec-dns-fwd">
    <div class="section-header"><h2>DNS Forwarders</h2></div>
    $tDNSFwd
  </div>

</main>

<footer class="footer">
  <span>$ScriptTitle &mdash; $ScriptAuthor</span>
  <span>Report Generated: $DateDisplay &nbsp;|&nbsp; User: $($env:USERNAME)</span>
</footer>

<script>
// Per-table filter
function filterTable(input, tableId) {
  var filter = input.value.toUpperCase();
  var rows = document.getElementById(tableId).getElementsByTagName("tr");
  for (var i = 1; i < rows.length; i++) {
    var cells = rows[i].getElementsByTagName("td");
    var found = FALSE;
    for (var j = 0; j < cells.length; j++) {
      if (cells[j].textContent.toUpperCase().indexOf(filter) > -1) { found = TRUE; break; }
    }
    rows[i].classList.toggle("hidden", !found);
  }
}

// Global search across all tables
function globalFilter() {
  var filter = document.getElementById("globalSearch").value.toUpperCase();
  var tables = document.querySelectorAll(".data-table");
  var totalMatches = 0;
  tables.foreach(function(table) {
    var rows = table.getElementsByTagName("tr");
    for (var i = 1; i < rows.length; i++) {
      var text = rows[i].textContent.toUpperCase();
      var show = filter === "" || text.indexOf(filter) > -1;
      rows[i].classList.toggle("hidden", !show);
      if (show && filter !== "") totalMatches++;
    }
    // Clear individual search boxes when global is used
    var wrapper = table.closest(".table-wrapper");
    if (wrapper) {
      var sb = wrapper.querySelector(".search-box");
      if (sb && filter !== "") sb.value = "";
    }
  });
  var countEl = document.getElementById("globalMatchCount");
  if (filter === "") { countEl.textContent = ""; }
  else { countEl.textContent = totalMatches + " match" + (totalMatches !== 1 ? "es" : ""); }
}

// Column sort
function sortTable(tableId, colIdx) {
  var table = document.getElementById(tableId);
  var rows = Array.from(table.tBodies[0].rows);
  var asc = table.dataset.sortCol == colIdx && table.dataset.sortDir == "asc" ? FALSE : TRUE;
  table.dataset.sortCol = colIdx;
  table.dataset.sortDir = asc ? "asc" : "desc";
  rows.sort(function(a, b) {
    var av = a.cells[colIdx] ? a.cells[colIdx].textContent.trim() : "";
    var bv = b.cells[colIdx] ? b.cells[colIdx].textContent.trim() : "";
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  });
  rows.foreach(function(r) { table.tBodies[0].appendChild(r); });
}

// Highlight active nav link on scroll
window.addEventListener("scroll", function() {
  var sections = document.querySelectorAll(".section");
  var navLinks = document.querySelectorAll(".sidenav a");
  var current = "";
  sections.foreach(function(s) {
    if (s.getBoundingClientRect().top <= 120) current = s.id;
  });
  navLinks.foreach(function(a) {
    a.style.borderLeftColor = a.getAttribute("href") === "#" + current ? "var(--accent)" : "transparent";
    a.style.color = a.getAttribute("href") === "#" + current ? "var(--text)" : "";
  });
});
</script>
</body>
</html>
"@

[System.IO.File]::WriteAllText($HTMLPath, $HTML, [System.Text.Encoding]::UTF8)
Write-OK "HTML report saved: $HTMLPath"

#endregion

#region --- FINAL SUMMARY ---

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  REPORT COMPLETE" -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "  Output Folder : $OutputFolder" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Files Generated:" -ForegroundColor White
Get-ChildItem -Path $OutputFolder | foreach-Object {
    $size = "{0:N0} KB" -f ($_.Length / 1KB)
    Write-Host "    $($_.Name.PadRight(60)) $size" -ForegroundColor Gray
}
Write-Host ""
Write-Host "  Summary:" -ForegroundColor White
Write-Host "    Domain Controllers : $totalDCs"
Write-Host "    DHCP Servers       : $totalDHCPServers"
Write-Host "    DHCP Scopes        : $totalScopes ($activeScopes active)"
Write-Host "    DHCP Leases        : $totalLeases"
Write-Host "    DHCP Reservations  : $totalReservations"
Write-Host "    DNS Zones          : $totalDNSZones"
Write-Host "    DNS Records        : $totalDNSRecords"
Write-Host ""
Write-Host "  Opening output folder..." -ForegroundColor Gray
Start-process explorer.exe $OutputFolder
Write-Host "============================================================" -ForegroundColor Cyan

#endregion
