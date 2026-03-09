#Requires -Version 5.1
<#
.SYNOPSIS
    AD_DHCP_Report - Enterprise DHCP Documentation Script
    Created by: Stephen McKee - Server Administrator 2 - IGT - Everi

.DESCRIPTION
    Comprehensively documents ALL DHCP configuration across every authorized
    DHCP server in the Active Directory environment, including:

      - AD Domain & DHCP Server Inventory
      - DHCP Server Settings & Configuration (per server)
      - DHCP Server Health & Statistics
      - IPv4 & IPv6 Scopes (all properties)
      - Scope Health & Utilization (free/in-use/expired leases)
      - Scope Statistics per server
      - Superscopes
      - Multicast Scopes
      - DHCP Failover Relationships
      - DHCP Policies (server-level and scope-level)
      - Reservations (all scopes)
      - Exclusion Ranges
      - DHCP Options (server, scope, reservation, policy level)
      - Active Leases
      - DHCP Audit Log Settings
      - DHCP Database Settings
      - AD-Authorized DHCP Servers

    Exports to:
      - Multiple CSV files (one per category)
      - Multi-tab XLSX workbook
      - Enterprise-grade dark-themed searchable HTML report

.NOTES
    Requirements:
      - PowerShell 5.1+
      - RSAT: DHCP Server Tools
      - RSAT: Active Directory Domain Services Tools
      - ImportExcel module (auto-installed from PSGallery if missing)
      - Run as Administrator or Domain Admin
#>

#region --- INITIALIZATION ---
Set-StrictMode -Version Latest
$ErrorActionPreference = "Continue"
$SkipXLSX = $false

$ScriptTitle   = "AD DHCP Report"
$ScriptAuthor  = "Stephen McKee - Server Administrator 2 - IGT - Everi"
$RunDate       = Get-Date
$DateTimeStamp = $RunDate.ToString("yyyy-MM-dd_HH-mm-ss")
$DateDisplay   = $RunDate.ToString("MMMM dd, yyyy hh:mm:ss tt")

$DesktopPath  = [Environment]::GetFolderPath("Desktop")
$OutputFolder = Join-Path $DesktopPath "AD_DHCP_Report_$DateTimeStamp"
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null

Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  $ScriptTitle" -ForegroundColor Yellow
Write-Host "  $ScriptAuthor" -ForegroundColor Yellow
Write-Host "  Started : $DateDisplay" -ForegroundColor Gray
Write-Host "  Output  : $OutputFolder" -ForegroundColor Gray
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host ""

# Auto-install ImportExcel
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "[*] ImportExcel not found. Installing..." -ForegroundColor Yellow
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
        Write-Host "[+] ImportExcel installed." -ForegroundColor Green
    } catch {
        Write-Warning "ImportExcel install failed. XLSX will be skipped. Error: $_"
        $SkipXLSX = $true
    }
}
Import-Module ImportExcel -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Web -ErrorAction SilentlyContinue
#endregion

#region --- HELPER FUNCTIONS ---
function Write-Section { param([string]$m) Write-Host ""; Write-Host "[>>] $m" -ForegroundColor Cyan }
function Write-OK      { param([string]$m) Write-Host "  [+] $m" -ForegroundColor Green }
function Write-Warn    { param([string]$m) Write-Host "  [!] $m" -ForegroundColor Yellow }
function Write-Fail    { param([string]$m) Write-Host "  [-] $m" -ForegroundColor Red }
function Write-Info    { param([string]$m) Write-Host "  [i] $m" -ForegroundColor Gray }

# Safe HTML encoder - falls back to manual replace if System.Web not available
function Safe-HtmlEncode {
    param([string]$str)
    try {
        return [System.Web.HttpUtility]::HtmlEncode($str)
    } catch {
        return $str -replace '&','&amp;' -replace '<','&lt;' -replace '>','&gt;' -replace '"','&quot;' -replace "'","&#39;"
    }
}

function Safe-Export-CSV {
    param([object[]]$Data, [string]$Path, [string]$Label)
    if ($Data -and $Data.Count -gt 0) {
        $Data | Export-Csv -Path $Path -NoTypeInformation -Force
        Write-OK "$Label -> CSV ($($Data.Count) rows)"
    } else {
        Write-Warn "$Label - No data"
    }
}

function ConvertTo-HtmlTable {
    param([object[]]$Data, [string]$TableId, [string]$Caption)
    if (-not $Data -or $Data.Count -eq 0) {
        return "<p class='no-data'>No data available.</p>"
    }
    $headers = $Data[0].PSObject.Properties.Name
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.Append("<div class='table-wrapper'>")
    [void]$sb.Append("<input type='text' class='search-box' placeholder='Search $Caption...' onkeyup=""filterTable(this,'$TableId')"" />")
    [void]$sb.Append("<div class='table-scroll'><table id='$TableId' class='data-table'>")
    [void]$sb.Append("<thead><tr>")
    $idx = 0
    foreach ($h in $headers) {
        [void]$sb.Append("<th onclick=""sortTable('$TableId',$idx)"">$h <span class='sort-icon'>&#8645;</span></th>")
        $idx++
    }
    [void]$sb.Append("</tr></thead><tbody>")
    foreach ($row in $Data) {
        [void]$sb.Append("<tr>")
        foreach ($h in $headers) {
            $val = $row.$h; if ($null -eq $val) { $val = "" }
            $vs  = Safe-HtmlEncode($val.ToString())
            $cls = ""
            # Scope state coloring
            if ($h -eq "State" -or $h -eq "ScopeState") {
                if ($val -eq "Active")   { $cls = " class='s-active'" }
                elseif ($val -eq "InActive" -or $val -eq "Inactive") { $cls = " class='s-inactive'" }
            }
            # Health / utilization coloring
            if ($h -eq "PercentInUse") {
                $pct = 0; [double]::TryParse($val, [ref]$pct) | Out-Null
                if ($pct -ge 90)     { $cls = " class='util-crit'" }
                elseif ($pct -ge 75) { $cls = " class='util-warn'" }
                else                 { $cls = " class='util-ok'" }
            }
            if ($h -eq "FailoverState" -or $h -eq "State") {
                if ($val -like "*Normal*" -or $val -eq "Active") { $cls = " class='bool-yes'" }
                elseif ($val -like "*Err*" -or $val -like "*Comm*") { $cls = " class='bool-no'" }
            }
            if ($h -eq "IsAuthorized") {
                if ($val -eq "True")  { $cls = " class='bool-yes'" }
                else                  { $cls = " class='bool-no'" }
            }
            [void]$sb.Append("<td$cls>$vs</td>")
        }
        [void]$sb.Append("</tr>")
    }
    [void]$sb.Append("</tbody></table></div></div>")
    return $sb.ToString()
}

function Safe-HtmlTableBuild {
    param($Data, [string]$TableId, [string]$Caption)
    try {
        return ConvertTo-HtmlTable -Data $Data -TableId $TableId -Caption $Caption
    } catch {
        Write-Warn "Table build failed for $Caption : $_"
        return "<p class='no-data'>Table generation failed for $Caption.</p>"
    }
}
#endregion

#region --- AD DOMAIN & DHCP SERVER INVENTORY ---
Write-Section "Collecting Active Directory Domain & DHCP Server Information"

$DomainInfo     = $null
$ForestInfo     = $null
$AuthDHCPServers = @()

try {
    Import-Module ActiveDirectory -ErrorAction Stop
    $DomainInfo  = Get-ADDomain  -ErrorAction Stop
    $ForestInfo  = Get-ADForest  -ErrorAction Stop
    Write-OK "Domain: $($DomainInfo.DNSRoot) | Forest: $($ForestInfo.Name)"
} catch {
    Write-Fail "AD query failed: $_"
}

# Get AD-authorized DHCP servers
try {
    Import-Module DhcpServer -ErrorAction Stop
    $AuthDHCPServers = Get-DhcpServerInDC -ErrorAction Stop | Sort-Object DnsName
    Write-OK "AD-Authorized DHCP Servers found: $($AuthDHCPServers.Count)"
} catch {
    Write-Fail "DHCP module / Get-DhcpServerInDC failed: $_"
    # Fall back to local machine
    $AuthDHCPServers = @([PSCustomObject]@{ DnsName = $env:COMPUTERNAME; IPAddress = $env:COMPUTERNAME })
}

$ADSummaryData = if ($DomainInfo) {
    [PSCustomObject]@{
        DomainName           = $DomainInfo.DNSRoot
        NetBIOSName          = $DomainInfo.NetBIOSName
        ForestName           = $ForestInfo.Name
        ForestMode           = $ForestInfo.ForestMode
        DomainMode           = $DomainInfo.DomainMode
        PDCEmulator          = $DomainInfo.PDCEmulator
        TotalDHCPServers     = $AuthDHCPServers.Count
        DHCPServers          = ($AuthDHCPServers.DnsName -join "; ")
        ReportGeneratedBy    = $env:USERNAME
        ReportDate           = $DateDisplay
    }
} else { @() }

$AuthServerData = $AuthDHCPServers | ForEach-Object {
    [PSCustomObject]@{
        DnsName   = $_.DnsName
        IPAddress = $_.IPAddress
    }
}

Safe-Export-CSV -Data @($ADSummaryData) -Path "$OutputFolder\01_AD_Domain_Summary.csv"      -Label "AD Domain Summary"
Safe-Export-CSV -Data $AuthServerData   -Path "$OutputFolder\02_DHCP_Authorized_Servers.csv" -Label "AD Authorized DHCP Servers"
#endregion

#region --- PER-SERVER DHCP DATA COLLECTION ---
Write-Section "Collecting DHCP Data from All Authorized Servers"

$AllServerSettings   = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllServerStats      = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllServerHealth     = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllAuditLog         = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllDatabase         = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllIPv4Scopes       = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllIPv6Scopes       = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllScopeStats       = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllScopeHealth      = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllSuperscopes      = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllMulticastScopes  = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllFailover         = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllPolicies         = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllReservations     = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllExclusions       = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllServerOptions    = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllScopeOptions     = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllLeases           = [System.Collections.Generic.List[PSCustomObject]]::new()

foreach ($srv in $AuthDHCPServers) {
    $dhcpServer = $srv.DnsName
    Write-Info "Processing DHCP Server: $dhcpServer"

    #-- Server Settings
    try {
        $cfg = Get-DhcpServerSetting -ComputerName $dhcpServer -ErrorAction Stop
        $AllServerSettings.Add([PSCustomObject]@{
            Server                      = $dhcpServer
            IsAuthorized                = $cfg.IsAuthorized
            IsDomainJoined              = $cfg.IsDomainJoined
            DynamicBootp                = $cfg.DynamicBootp
            RestoreStatus               = $cfg.RestoreStatus
            IsConflictDetectionEnabled  = $cfg.IsConflictDetectionEnabled
            ConflictDetectionAttempts   = $cfg.ConflictDetectionAttempts
            NapEnabled                  = $cfg.NapEnabled
            NpsUnreachableAction        = $cfg.NpsUnreachableAction
            ActivatePolicies            = $cfg.ActivatePolicies
        })
        Write-OK "  Server settings collected"
    } catch { Write-Warn "  Server settings failed on $dhcpServer : $_" }

    #-- Server Statistics
    try {
        $stats = Get-DhcpServerv4Statistics -ComputerName $dhcpServer -ErrorAction Stop
        $AllServerStats.Add([PSCustomObject]@{
            Server              = $dhcpServer
            TotalScopes         = $stats.TotalScopes
            ScopesWithDelayOffer = $stats.ScopesWithDelayOffer
            TotalAddresses      = $stats.TotalAddresses
            InUse               = $stats.InUse
            Available           = $stats.Available
            PercentInUse        = [math]::Round($stats.PercentInUse, 2)
            PercentAvailable    = [math]::Round($stats.PercentAvailable, 2)
            Discovers           = $stats.Discovers
            Offers              = $stats.Offers
            Requests            = $stats.Requests
            Acks                = $stats.Acks
            Naks                = $stats.Naks
            Declines            = $stats.Declines
            Releases            = $stats.Releases
            ServerStartTime     = $stats.ServerStartTime
        })
        Write-OK "  Server statistics collected"
    } catch { Write-Warn "  Server statistics failed on $dhcpServer : $_" }

    #-- Audit Log Settings
    try {
        $audit = Get-DhcpServerAuditLog -ComputerName $dhcpServer -ErrorAction Stop
        $AllAuditLog.Add([PSCustomObject]@{
            Server      = $dhcpServer
            Enable      = $audit.Enable
            Path        = $audit.Path
            MaxMBFileSize = $audit.MaxMBFileSize
            MinMBDiskSpace = $audit.MinMBDiskSpace
            DiskCheckInterval = $audit.DiskCheckInterval
        })
        Write-OK "  Audit log settings collected"
    } catch { Write-Warn "  Audit log settings failed on $dhcpServer : $_" }

    #-- Database Settings
    try {
        $db = Get-DhcpServerDatabase -ComputerName $dhcpServer -ErrorAction Stop
        $AllDatabase.Add([PSCustomObject]@{
            Server          = $dhcpServer
            FileName        = $db.FileName
            BackupPath      = $db.BackupPath
            BackupInterval  = $db.BackupInterval
            CleanupInterval = $db.CleanupInterval
            RestoreFromBackup = $db.RestoreFromBackup
            LoggingEnabled  = $db.LoggingEnabled
        })
        Write-OK "  Database settings collected"
    } catch { Write-Warn "  Database settings failed on $dhcpServer : $_" }

    #-- Server-Level Options
    try {
        $srvOpts = Get-DhcpServerv4OptionValue -ComputerName $dhcpServer -ErrorAction Stop
        foreach ($opt in $srvOpts) {
            $AllServerOptions.Add([PSCustomObject]@{
                Server      = $dhcpServer
                OptionId    = $opt.OptionId
                Name        = $opt.Name
                Type        = $opt.Type
                Value       = ($opt.Value -join "; ")
                VendorClass = $opt.VendorClass
                UserClass   = $opt.UserClass
            })
        }
        Write-OK "  Server options collected ($($srvOpts.Count))"
    } catch { Write-Warn "  Server options failed on $dhcpServer : $_" }

    #-- Failover Relationships
    try {
        $failovers = Get-DhcpServerv4Failover -ComputerName $dhcpServer -ErrorAction Stop
        foreach ($fo in $failovers) {
            $AllFailover.Add([PSCustomObject]@{
                Server              = $dhcpServer
                Name                = $fo.Name
                Mode                = $fo.Mode
                PartnerServer       = $fo.PartnerServer
                State               = $fo.State
                AutoStateTransition = $fo.AutoStateTransition
                EnableAuth          = $fo.EnableAuth
                MaxClientLeadTime   = $fo.MaxClientLeadTime
                StateSwitchInterval = $fo.StateSwitchInterval
                LoadBalancePercent  = $fo.LoadBalancePercent
                ServerRole          = $fo.ServerRole
                PrimaryServerIP     = $fo.PrimaryServerIP
                SecondaryServerIP   = $fo.SecondaryServerIP
                ScopeId             = ($fo.ScopeId -join "; ")
            })
        }
        Write-OK "  Failover relationships collected ($($failovers.Count))"
    } catch { Write-Warn "  Failover failed on $dhcpServer : $_" }

    #-- Server-Level Policies
    try {
        $srvPolicies = Get-DhcpServerv4Policy -ComputerName $dhcpServer -ErrorAction Stop
        foreach ($pol in $srvPolicies) {
            $AllPolicies.Add([PSCustomObject]@{
                Server          = $dhcpServer
                ScopeId         = "Server-Level"
                Name            = $pol.Name
                Enabled         = $pol.Enabled
                Description     = $pol.Description
                ProcessingOrder = $pol.ProcessingOrder
                Condition       = $pol.Condition
                VendorClass     = ($pol.VendorClass -join "; ")
                UserClass       = ($pol.UserClass -join "; ")
                MacAddresses    = ($pol.MacAddress -join "; ")
                ClientId        = ($pol.ClientId -join "; ")
                Fqdn            = ($pol.Fqdn -join "; ")
                RelayAgent      = ($pol.RelayAgent -join "; ")
                CircuitId       = ($pol.CircuitId -join "; ")
                RemoteId        = ($pol.RemoteId -join "; ")
                SubscriberId    = ($pol.SubscriberId -join "; ")
            })
        }
        Write-OK "  Server policies collected ($($srvPolicies.Count))"
    } catch { Write-Warn "  Server policies failed on $dhcpServer : $_" }

    #-- Superscopes
    try {
        $superscopes = Get-DhcpServerv4Superscope -ComputerName $dhcpServer -ErrorAction Stop
        foreach ($ss in $superscopes) {
            $AllSuperscopes.Add([PSCustomObject]@{
                Server          = $dhcpServer
                SuperscopeName  = $ss.SuperscopeName
                ScopeId         = ($ss.ScopeId -join "; ")
                State           = $ss.State
            })
        }
        Write-OK "  Superscopes collected ($($superscopes.Count))"
    } catch { Write-Warn "  Superscopes failed on $dhcpServer : $_" }

    #-- Multicast Scopes
    try {
        $mscopes = Get-DhcpServerv4MulticastScope -ComputerName $dhcpServer -ErrorAction Stop
        foreach ($ms in $mscopes) {
            $AllMulticastScopes.Add([PSCustomObject]@{
                Server         = $dhcpServer
                Name           = $ms.Name
                StartRange     = $ms.StartRange
                EndRange       = $ms.EndRange
                State          = $ms.State
                Ttl            = $ms.Ttl
                LeaseDuration  = $ms.LeaseDuration
                ExpiryTime     = $ms.ExpiryTime
                Description    = $ms.Description
            })
        }
        Write-OK "  Multicast scopes collected ($($mscopes.Count))"
    } catch { Write-Warn "  Multicast scopes failed on $dhcpServer : $_" }

    #-- IPv4 Scopes
    try {
        $v4scopes = Get-DhcpServerv4Scope -ComputerName $dhcpServer -ErrorAction Stop
        Write-OK "  IPv4 Scopes: $($v4scopes.Count)"

        foreach ($scope in $v4scopes) {
            $AllIPv4Scopes.Add([PSCustomObject]@{
                Server              = $dhcpServer
                ScopeId             = $scope.ScopeId
                Name                = $scope.Name
                SubnetMask          = $scope.SubnetMask
                StartRange          = $scope.StartRange
                EndRange            = $scope.EndRange
                State               = $scope.State
                LeaseDuration       = $scope.LeaseDuration
                Description         = $scope.Description
                SuperscopeName      = $scope.SuperscopeName
                MaxBootpClients     = $scope.MaxBootpClients
                ActivatePolicies    = $scope.ActivatePolicies
                Delay               = $scope.Delay
                Type                = $scope.Type
                NapEnable           = $scope.NapEnable
                NapProfile          = $scope.NapProfile
            })

            #-- Scope Statistics / Health
            try {
                $scopeStat = Get-DhcpServerv4ScopeStatistics -ComputerName $dhcpServer -ScopeId $scope.ScopeId -ErrorAction Stop
                $pctInUse  = if ($scopeStat.PercentageInUse) { [math]::Round($scopeStat.PercentageInUse, 2) } else { 0 }

                # Health rating
                $health = if ($scope.State -ne "Active") { "Inactive" }
                          elseif ($pctInUse -ge 90)      { "Critical" }
                          elseif ($pctInUse -ge 75)      { "Warning" }
                          else                            { "Healthy" }

                $AllScopeStats.Add([PSCustomObject]@{
                    Server          = $dhcpServer
                    ScopeId         = $scope.ScopeId
                    ScopeName       = $scope.Name
                    State           = $scope.State
                    TotalAddresses  = $scopeStat.TotalAddresses
                    InUse           = $scopeStat.InUse
                    Available       = $scopeStat.Available
                    Reserved        = $scopeStat.Reserved
                    Pending         = $scopeStat.Pending
                    PercentInUse    = $pctInUse
                    Health          = $health
                })

                $AllScopeHealth.Add([PSCustomObject]@{
                    Server       = $dhcpServer
                    ScopeId      = $scope.ScopeId
                    ScopeName    = $scope.Name
                    SubnetMask   = $scope.SubnetMask
                    StartRange   = $scope.StartRange
                    EndRange     = $scope.EndRange
                    State        = $scope.State
                    TotalIPs     = $scopeStat.TotalAddresses
                    InUse        = $scopeStat.InUse
                    Available    = $scopeStat.Available
                    Reserved     = $scopeStat.Reserved
                    PercentInUse = $pctInUse
                    Health       = $health
                    LeaseDuration = $scope.LeaseDuration
                    SuperscopeName = $scope.SuperscopeName
                })
            } catch { Write-Warn "  Scope stats failed for $($scope.ScopeId) : $_" }

            #-- Exclusions
            try {
                $excl = Get-DhcpServerv4ExclusionRange -ComputerName $dhcpServer -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($ex in $excl) {
                    $AllExclusions.Add([PSCustomObject]@{
                        Server     = $dhcpServer
                        ScopeId    = $scope.ScopeId
                        ScopeName  = $scope.Name
                        StartRange = $ex.StartRange
                        EndRange   = $ex.EndRange
                    })
                }
            } catch {}

            #-- Reservations
            try {
                $reservations = Get-DhcpServerv4Reservation -ComputerName $dhcpServer -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($res in $reservations) {
                    $AllReservations.Add([PSCustomObject]@{
                        Server      = $dhcpServer
                        ScopeId     = $scope.ScopeId
                        ScopeName   = $scope.Name
                        IPAddress   = $res.IPAddress
                        ClientId    = $res.ClientId
                        Name        = $res.Name
                        Description = $res.Description
                        Type        = $res.Type
                    })
                }
            } catch {}

            #-- Scope-Level Options
            try {
                $scopeOpts = Get-DhcpServerv4OptionValue -ComputerName $dhcpServer -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($opt in $scopeOpts) {
                    $AllScopeOptions.Add([PSCustomObject]@{
                        Server      = $dhcpServer
                        ScopeId     = $scope.ScopeId
                        ScopeName   = $scope.Name
                        OptionId    = $opt.OptionId
                        Name        = $opt.Name
                        Type        = $opt.Type
                        Value       = ($opt.Value -join "; ")
                        VendorClass = $opt.VendorClass
                        UserClass   = $opt.UserClass
                    })
                }
            } catch {}

            #-- Scope-Level Policies
            try {
                $scopePolicies = Get-DhcpServerv4Policy -ComputerName $dhcpServer -ScopeId $scope.ScopeId -ErrorAction Stop
                foreach ($pol in $scopePolicies) {
                    $AllPolicies.Add([PSCustomObject]@{
                        Server          = $dhcpServer
                        ScopeId         = $scope.ScopeId
                        Name            = $pol.Name
                        Enabled         = $pol.Enabled
                        Description     = $pol.Description
                        ProcessingOrder = $pol.ProcessingOrder
                        Condition       = $pol.Condition
                        VendorClass     = ($pol.VendorClass -join "; ")
                        UserClass       = ($pol.UserClass -join "; ")
                        MacAddresses    = ($pol.MacAddress -join "; ")
                        ClientId        = ($pol.ClientId -join "; ")
                        Fqdn            = ($pol.Fqdn -join "; ")
                        RelayAgent      = ($pol.RelayAgent -join "; ")
                        CircuitId       = ($pol.CircuitId -join "; ")
                        RemoteId        = ($pol.RemoteId -join "; ")
                        SubscriberId    = ($pol.SubscriberId -join "; ")
                    })
                }
            } catch {}

            #-- Active Leases (capped at 5000 per scope to avoid massive output)
            try {
                $leases = Get-DhcpServerv4Lease -ComputerName $dhcpServer -ScopeId $scope.ScopeId -ErrorAction Stop |
                          Select-Object -First 5000
                foreach ($lease in $leases) {
                    $AllLeases.Add([PSCustomObject]@{
                        Server          = $dhcpServer
                        ScopeId         = $scope.ScopeId
                        ScopeName       = $scope.Name
                        IPAddress       = $lease.IPAddress
                        ClientId        = $lease.ClientId
                        HostName        = $lease.HostName
                        AddressState    = $lease.AddressState
                        LeaseExpiryTime = if ($lease.LeaseExpiryTime) { $lease.LeaseExpiryTime.ToString() } else { "Reservation" }
                        ClientType      = $lease.ClientType
                        Description     = $lease.Description
                        DnsRegistration = $lease.DnsRegistration
                        DnsRR           = $lease.DnsRR
                    })
                }
                Write-Info "    Scope $($scope.ScopeId) - $($leases.Count) leases"
            } catch { Write-Warn "  Leases failed for $($scope.ScopeId) : $_" }
        }
    } catch { Write-Fail "  IPv4 Scopes failed on $dhcpServer : $_" }

    #-- IPv6 Scopes
    try {
        $v6scopes = Get-DhcpServerv6Scope -ComputerName $dhcpServer -ErrorAction Stop
        foreach ($scope in $v6scopes) {
            $AllIPv6Scopes.Add([PSCustomObject]@{
                Server          = $dhcpServer
                Prefix          = $scope.Prefix
                Name            = $scope.Name
                State           = $scope.State
                Preference      = $scope.Preference
                ValidLifeTime   = $scope.ValidLifeTime
                T1              = $scope.T1
                T2              = $scope.T2
                Description     = $scope.Description
            })
        }
        Write-OK "  IPv6 Scopes collected ($($v6scopes.Count))"
    } catch { Write-Warn "  IPv6 Scopes failed on $dhcpServer : $_" }
}

#-- Derived summaries
$ScopeSummary = $AllScopeHealth | Sort-Object Server, ScopeId

$CriticalScopes  = @($AllScopeHealth | Where-Object { $_.Health -eq "Critical" })
$WarningScopes   = @($AllScopeHealth | Where-Object { $_.Health -eq "Warning" })
$HealthyScopes   = @($AllScopeHealth | Where-Object { $_.Health -eq "Healthy" })
$InactiveScopes  = @($AllScopeHealth | Where-Object { $_.Health -eq "Inactive" })

Write-Section "Exporting CSV Files"
Safe-Export-CSV -Data $AllServerSettings.ToArray()  -Path "$OutputFolder\03_DHCP_Server_Settings.csv"   -Label "DHCP Server Settings"
Safe-Export-CSV -Data $AllServerStats.ToArray()     -Path "$OutputFolder\04_DHCP_Server_Statistics.csv" -Label "DHCP Server Statistics"
Safe-Export-CSV -Data $AllAuditLog.ToArray()        -Path "$OutputFolder\05_DHCP_Audit_Log.csv"         -Label "Audit Log Settings"
Safe-Export-CSV -Data $AllDatabase.ToArray()        -Path "$OutputFolder\06_DHCP_Database.csv"          -Label "Database Settings"
Safe-Export-CSV -Data $AllIPv4Scopes.ToArray()      -Path "$OutputFolder\07_DHCP_IPv4_Scopes.csv"       -Label "IPv4 Scopes"
Safe-Export-CSV -Data $AllIPv6Scopes.ToArray()      -Path "$OutputFolder\08_DHCP_IPv6_Scopes.csv"       -Label "IPv6 Scopes"
Safe-Export-CSV -Data $ScopeSummary                 -Path "$OutputFolder\09_DHCP_Scope_Health.csv"       -Label "Scope Health Summary"
Safe-Export-CSV -Data $AllScopeStats.ToArray()      -Path "$OutputFolder\10_DHCP_Scope_Statistics.csv"  -Label "Scope Statistics"
Safe-Export-CSV -Data $AllSuperscopes.ToArray()     -Path "$OutputFolder\11_DHCP_Superscopes.csv"        -Label "Superscopes"
Safe-Export-CSV -Data $AllMulticastScopes.ToArray() -Path "$OutputFolder\12_DHCP_Multicast_Scopes.csv"  -Label "Multicast Scopes"
Safe-Export-CSV -Data $AllFailover.ToArray()        -Path "$OutputFolder\13_DHCP_Failover.csv"          -Label "Failover Relationships"
Safe-Export-CSV -Data $AllPolicies.ToArray()        -Path "$OutputFolder\14_DHCP_Policies.csv"          -Label "DHCP Policies"
Safe-Export-CSV -Data $AllReservations.ToArray()    -Path "$OutputFolder\15_DHCP_Reservations.csv"      -Label "Reservations"
Safe-Export-CSV -Data $AllExclusions.ToArray()      -Path "$OutputFolder\16_DHCP_Exclusions.csv"        -Label "Exclusion Ranges"
Safe-Export-CSV -Data $AllServerOptions.ToArray()   -Path "$OutputFolder\17_DHCP_Server_Options.csv"    -Label "Server-Level Options"
Safe-Export-CSV -Data $AllScopeOptions.ToArray()    -Path "$OutputFolder\18_DHCP_Scope_Options.csv"     -Label "Scope-Level Options"
Safe-Export-CSV -Data $AllLeases.ToArray()          -Path "$OutputFolder\19_DHCP_Active_Leases.csv"     -Label "Active Leases"
#endregion

#region --- XLSX EXPORT ---
if (-not $SkipXLSX) {
    Write-Section "Building Excel Workbook (.xlsx)"
    $XLSXPath = "$OutputFolder\AD_DHCP_Report_$DateTimeStamp.xlsx"
    $xlP = @{ Path = $XLSXPath; AutoSize = $true; FreezeTopRow = $true; BoldTopRow = $true; TableStyle = "Medium9" }

    function Add-Sheet {
        param($Data, [string]$Sheet)
        if ($Data -and @($Data).Count -gt 0) {
            @($Data) | Export-Excel @xlP -WorksheetName $Sheet -Append
            Write-OK "Sheet '$Sheet' ($(@($Data).Count) rows)"
        } else { Write-Warn "Sheet '$Sheet' - no data" }
    }

    Add-Sheet -Data @($ADSummaryData)              -Sheet "AD_Summary"
    Add-Sheet -Data $AuthServerData                -Sheet "DHCP_Auth_Servers"
    Add-Sheet -Data $AllServerSettings.ToArray()   -Sheet "DHCP_Server_Settings"
    Add-Sheet -Data $AllServerStats.ToArray()      -Sheet "DHCP_Server_Statistics"
    Add-Sheet -Data $AllAuditLog.ToArray()         -Sheet "DHCP_Audit_Log"
    Add-Sheet -Data $AllDatabase.ToArray()         -Sheet "DHCP_Database"
    Add-Sheet -Data $AllIPv4Scopes.ToArray()       -Sheet "DHCP_IPv4_Scopes"
    Add-Sheet -Data $AllIPv6Scopes.ToArray()       -Sheet "DHCP_IPv6_Scopes"
    Add-Sheet -Data $ScopeSummary                  -Sheet "Scope_Health_Summary"
    Add-Sheet -Data $AllScopeStats.ToArray()       -Sheet "Scope_Statistics"
    Add-Sheet -Data $AllSuperscopes.ToArray()      -Sheet "Superscopes"
    Add-Sheet -Data $AllMulticastScopes.ToArray()  -Sheet "Multicast_Scopes"
    Add-Sheet -Data $AllFailover.ToArray()         -Sheet "DHCP_Failover"
    Add-Sheet -Data $AllPolicies.ToArray()         -Sheet "DHCP_Policies"
    Add-Sheet -Data $AllReservations.ToArray()     -Sheet "DHCP_Reservations"
    Add-Sheet -Data $AllExclusions.ToArray()       -Sheet "DHCP_Exclusions"
    Add-Sheet -Data $AllServerOptions.ToArray()    -Sheet "Server_Options"
    Add-Sheet -Data $AllScopeOptions.ToArray()     -Sheet "Scope_Options"
    Add-Sheet -Data $AllLeases.ToArray()           -Sheet "Active_Leases"
    Write-OK "XLSX saved: $XLSXPath"
}
#endregion

#region --- HTML REPORT ---
Write-Section "Building Enterprise HTML Report"

# Turn off strict mode for HTML generation
Set-StrictMode -Off

# Derived stats
$totalServers    = $AuthDHCPServers.Count
$totalV4Scopes   = ($AllIPv4Scopes | Select-Object ScopeId -Unique).Count
$totalV6Scopes   = ($AllIPv6Scopes | Select-Object Prefix  -Unique).Count
$totalReserv     = $AllReservations.Count
$totalLeases     = $AllLeases.Count
$totalExcl       = $AllExclusions.Count
$totalPolicies   = $AllPolicies.Count
$totalFailover   = $AllFailover.Count
$totalSuperscope = $AllSuperscopes.Count
$cntCritical     = $CriticalScopes.Count
$cntWarning      = $WarningScopes.Count
$cntHealthy      = $HealthyScopes.Count
$cntInactive     = $InactiveScopes.Count

# Ensure table variables have safe defaults
$tADSummary     = Safe-HtmlTableBuild -Data @($ADSummaryData)              -TableId "t_adsum"    -Caption "AD Summary"
$tAuthSrv       = Safe-HtmlTableBuild -Data $AuthServerData                -TableId "t_authsrv"  -Caption "Authorized Servers"
$tSrvSettings   = Safe-HtmlTableBuild -Data $AllServerSettings.ToArray()   -TableId "t_srvsett"  -Caption "Server Settings"
$tSrvStats      = Safe-HtmlTableBuild -Data $AllServerStats.ToArray()      -TableId "t_srvstats" -Caption "Server Statistics"
$tAuditLog      = Safe-HtmlTableBuild -Data $AllAuditLog.ToArray()         -TableId "t_audit"    -Caption "Audit Log"
$tDatabase      = Safe-HtmlTableBuild -Data $AllDatabase.ToArray()         -TableId "t_db"       -Caption "Database Settings"
$tV4Scopes      = Safe-HtmlTableBuild -Data $AllIPv4Scopes.ToArray()       -TableId "t_v4scope"  -Caption "IPv4 Scopes"
$tV6Scopes      = Safe-HtmlTableBuild -Data $AllIPv6Scopes.ToArray()       -TableId "t_v6scope"  -Caption "IPv6 Scopes"
$tScopeHealth   = Safe-HtmlTableBuild -Data $ScopeSummary                  -TableId "t_health"   -Caption "Scope Health"
$tScopeStats    = Safe-HtmlTableBuild -Data $AllScopeStats.ToArray()       -TableId "t_scstat"   -Caption "Scope Statistics"
$tSuperscopes   = Safe-HtmlTableBuild -Data $AllSuperscopes.ToArray()      -TableId "t_super"    -Caption "Superscopes"
$tMulticast     = Safe-HtmlTableBuild -Data $AllMulticastScopes.ToArray()  -TableId "t_mcast"    -Caption "Multicast Scopes"
$tFailover      = Safe-HtmlTableBuild -Data $AllFailover.ToArray()         -TableId "t_fo"       -Caption "Failover"
$tPolicies      = Safe-HtmlTableBuild -Data $AllPolicies.ToArray()         -TableId "t_pol"      -Caption "Policies"
$tReservations  = Safe-HtmlTableBuild -Data $AllReservations.ToArray()     -TableId "t_res"      -Caption "Reservations"
$tExclusions    = Safe-HtmlTableBuild -Data $AllExclusions.ToArray()       -TableId "t_excl"     -Caption "Exclusions"
$tSrvOptions    = Safe-HtmlTableBuild -Data $AllServerOptions.ToArray()    -TableId "t_srvopt"   -Caption "Server Options"
$tScopeOptions  = Safe-HtmlTableBuild -Data $AllScopeOptions.ToArray()     -TableId "t_scopt"    -Caption "Scope Options"
$tLeases        = Safe-HtmlTableBuild -Data $AllLeases.ToArray()           -TableId "t_leases"   -Caption "Active Leases"

if (-not $tADSummary)    { $tADSummary    = "<p class='no-data'>No data.</p>" }
if (-not $tAuthSrv)      { $tAuthSrv      = "<p class='no-data'>No data.</p>" }
if (-not $tSrvSettings)  { $tSrvSettings  = "<p class='no-data'>No data.</p>" }
if (-not $tSrvStats)     { $tSrvStats     = "<p class='no-data'>No data.</p>" }
if (-not $tAuditLog)     { $tAuditLog     = "<p class='no-data'>No data.</p>" }
if (-not $tDatabase)     { $tDatabase     = "<p class='no-data'>No data.</p>" }
if (-not $tV4Scopes)     { $tV4Scopes     = "<p class='no-data'>No data.</p>" }
if (-not $tV6Scopes)     { $tV6Scopes     = "<p class='no-data'>No data.</p>" }
if (-not $tScopeHealth)  { $tScopeHealth  = "<p class='no-data'>No data.</p>" }
if (-not $tScopeStats)   { $tScopeStats   = "<p class='no-data'>No data.</p>" }
if (-not $tSuperscopes)  { $tSuperscopes  = "<p class='no-data'>No data.</p>" }
if (-not $tMulticast)    { $tMulticast    = "<p class='no-data'>No data.</p>" }
if (-not $tFailover)     { $tFailover     = "<p class='no-data'>No data.</p>" }
if (-not $tPolicies)     { $tPolicies     = "<p class='no-data'>No data.</p>" }
if (-not $tReservations) { $tReservations = "<p class='no-data'>No data.</p>" }
if (-not $tExclusions)   { $tExclusions   = "<p class='no-data'>No data.</p>" }
if (-not $tSrvOptions)   { $tSrvOptions   = "<p class='no-data'>No data.</p>" }
if (-not $tScopeOptions) { $tScopeOptions = "<p class='no-data'>No data.</p>" }
if (-not $tLeases)       { $tLeases       = "<p class='no-data'>No data.</p>" }

# Health chart data
$chartLabels = "'Critical','Warning','Healthy','Inactive'"
$chartCounts = "$cntCritical,$cntWarning,$cntHealthy,$cntInactive"
$chartColors = "'#f85149','#d29922','#3fb950','#7d8590'"

# Utilization bar chart (top scopes by usage)
$topScopes = $AllScopeStats | Sort-Object PercentInUse -Descending | Select-Object -First 12
$scopeChartLabels = ($topScopes | ForEach-Object { "'$($_.ScopeId)'" }) -join ","
$scopeChartCounts = ($topScopes | ForEach-Object { $_.PercentInUse }) -join ","

$htmlDomainName = if ($DomainInfo) { $DomainInfo.DNSRoot } else { "N/A" }
$htmlForestName = if ($ForestInfo) { $ForestInfo.Name    } else { "N/A" }
$htmlRunByUser  = $env:USERNAME

$HTMLPath = "$OutputFolder\AD_DHCP_Report_$DateTimeStamp.html"

$HTML = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1.0"/>
<title>AD DHCP Report - Stephen McKee - IGT Everi</title>
<style>
:root {
  --bg:#0d1117; --surface:#161b22; --surface2:#1c2128; --border:#30363d;
  --accent:#2563eb; --text:#e6edf3; --muted:#7d8590;
  --green:#3fb950; --yellow:#d29922; --red:#f85149;
  --purple:#a371f7; --orange:#db6d28; --cyan:#39d3f2; --teal:#56d364;
}
*{box-sizing:border-box;margin:0;padding:0;}
html{scroll-behavior:smooth;}
body{background:var(--bg);color:var(--text);font-family:'Segoe UI',system-ui,sans-serif;font-size:13px;}

.top-header{
  background:linear-gradient(135deg,#0a0e14 0%,#161b22 60%,#0a0e14 100%);
  border-bottom:2px solid var(--accent);padding:22px 32px 18px;
  position:sticky;top:0;z-index:200;
  display:flex;justify-content:space-between;align-items:flex-end;
}
.title-block h1{font-size:21px;font-weight:700;color:#fff;letter-spacing:.5px;}
.title-block h1 em{color:var(--cyan);font-style:normal;}
.title-block .sub{font-size:11px;color:var(--muted);margin-top:3px;}
.meta-block{text-align:right;font-size:11px;color:var(--muted);line-height:1.8;}
.meta-block strong{color:var(--cyan);}

.sidenav{
  position:fixed;top:0;left:0;width:220px;height:100vh;
  background:var(--surface);border-right:1px solid var(--border);
  overflow-y:auto;padding-top:85px;z-index:100;
}
.nav-group{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:1.2px;padding:14px 16px 5px;}
.sidenav a{
  display:block;padding:6px 16px;color:var(--muted);text-decoration:none;
  font-size:12px;border-left:3px solid transparent;transition:all .15s;
}
.sidenav a:hover,.sidenav a.active{color:var(--text);background:var(--surface2);border-left-color:var(--accent);}

.main{margin-left:220px;padding:22px 28px 40px;}

.global-bar{
  background:var(--surface);border:1px solid var(--border);border-radius:8px;
  padding:12px 18px;margin-bottom:20px;display:flex;align-items:center;gap:12px;flex-wrap:wrap;
}
.global-bar input{
  background:var(--surface2);border:1px solid var(--border);color:var(--text);
  padding:7px 14px;border-radius:6px;font-size:13px;width:340px;outline:none;
  transition:border-color .2s;
}
.global-bar input:focus{border-color:var(--accent);}
.global-bar label{color:var(--muted);font-size:12px;}
.match-count{font-size:12px;color:var(--cyan);}

.cards-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(148px,1fr));gap:12px;margin-bottom:24px;}
.card{
  background:var(--surface);border:1px solid var(--border);border-radius:10px;
  padding:15px;position:relative;overflow:hidden;
}
.card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.card.c-blue::before{background:var(--accent);}
.card.c-green::before{background:var(--green);}
.card.c-yellow::before{background:var(--yellow);}
.card.c-red::before{background:var(--red);}
.card.c-purple::before{background:var(--purple);}
.card.c-cyan::before{background:var(--cyan);}
.card.c-orange::before{background:var(--orange);}
.card.c-teal::before{background:var(--teal);}
.card.c-muted::before{background:var(--muted);}
.card-val{font-size:30px;font-weight:700;color:#fff;line-height:1.1;}
.card-lbl{font-size:11px;color:var(--muted);margin-top:4px;}
.card.c-red .card-val{color:var(--red);}
.card.c-yellow .card-val{color:var(--yellow);}

.charts-row{display:grid;grid-template-columns:1fr 2fr;gap:16px;margin-bottom:22px;}
.chart-container{
  background:var(--surface);border:1px solid var(--border);border-radius:10px;
  padding:18px;
}
.chart-container h3{font-size:13px;color:var(--muted);margin-bottom:14px;font-weight:500;}

/* Donut chart */
.donut-wrap{display:flex;align-items:center;gap:20px;flex-wrap:wrap;}
.donut-svg{flex-shrink:0;}
.donut-legend{display:flex;flex-direction:column;gap:8px;}
.legend-item{display:flex;align-items:center;gap:8px;font-size:12px;}
.legend-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;}

/* Bar chart */
.chart-inner{display:flex;align-items:flex-end;gap:5px;height:130px;padding-bottom:20px;position:relative;}
.bar-wrap{display:flex;flex-direction:column;align-items:center;flex:1;min-width:30px;height:100%;justify-content:flex-end;}
.bar{width:100%;border-radius:3px 3px 0 0;min-height:4px;transition:height .3s;}
.bar-lbl{font-size:8px;color:var(--muted);margin-top:4px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;width:100%;}
.bar-count{font-size:9px;color:var(--text);margin-bottom:2px;font-weight:600;}

.section{margin-bottom:18px;border:1px solid var(--border);border-radius:10px;overflow:hidden;}
.section-header{
  display:flex;align-items:center;gap:10px;padding:12px 16px;
  background:var(--surface2);cursor:pointer;user-select:none;transition:background .15s;
}
.section-header:hover{background:#1e242c;}
.section-header h2{font-size:15px;font-weight:600;color:#fff;flex:1;}
.badge{background:var(--accent);color:#fff;font-size:10px;font-weight:700;padding:2px 8px;border-radius:20px;}
.badge.red{background:var(--red);}
.badge.green{background:var(--green);}
.badge.yellow{background:var(--yellow);}
.badge.muted{background:var(--muted);}
.chevron{width:18px;height:18px;flex-shrink:0;color:var(--muted);transition:transform .25s ease;}
.section.collapsed .chevron{transform:rotate(-90deg);}
.section-body{
  padding:16px;overflow:hidden;max-height:20000px;
  transition:max-height .35s ease,padding .3s ease,opacity .25s ease;opacity:1;
}
.section.collapsed .section-body{max-height:0;padding-top:0;padding-bottom:0;opacity:0;pointer-events:none;}
.collapse-controls{display:flex;gap:8px;align-items:center;}
.ctrl-btn{
  background:var(--surface2);border:1px solid var(--border);color:var(--muted);
  padding:5px 12px;border-radius:5px;font-size:11px;cursor:pointer;transition:all .15s;white-space:nowrap;
}
.ctrl-btn:hover{background:var(--surface);color:var(--text);border-color:var(--accent);}

.table-wrapper{position:relative;}
.search-box{
  background:var(--surface2);border:1px solid var(--border);color:var(--text);
  padding:6px 12px;border-radius:5px;font-size:12px;margin-bottom:8px;min-width:260px;outline:none;
}
.search-box:focus{border-color:var(--accent);}
.table-scroll{overflow-x:auto;border-radius:8px;border:1px solid var(--border);}
.data-table{width:100%;border-collapse:collapse;white-space:nowrap;}
.data-table thead{background:var(--surface2);position:sticky;top:0;z-index:10;}
.data-table th{
  padding:8px 12px;text-align:left;font-size:11px;font-weight:600;
  color:var(--muted);text-transform:uppercase;letter-spacing:.5px;
  border-bottom:1px solid var(--border);cursor:pointer;user-select:none;
}
.data-table th:hover{color:var(--text);}
.sort-icon{opacity:.35;font-size:10px;}
.data-table td{
  padding:7px 12px;border-bottom:1px solid #21262d;font-size:12px;
  max-width:300px;overflow:hidden;text-overflow:ellipsis;
}
.data-table tbody tr:hover{background:var(--surface2);}
.data-table tbody tr:last-child td{border-bottom:none;}
.data-table tr.hidden{display:none;}

td.s-active    {color:var(--green);font-weight:600;}
td.s-inactive  {color:var(--muted);}
td.util-crit   {color:var(--red);font-weight:700;}
td.util-warn   {color:var(--yellow);font-weight:600;}
td.util-ok     {color:var(--green);}
td.bool-yes    {color:var(--green);font-weight:600;}
td.bool-no     {color:var(--muted);}

.no-data{color:var(--muted);font-style:italic;padding:12px 0;}

.health-pills{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap;}
.pill{padding:6px 16px;border-radius:20px;font-size:12px;font-weight:600;display:flex;align-items:center;gap:6px;}
.pill-crit  {background:rgba(248,81,73,.15);border:1px solid var(--red);color:var(--red);}
.pill-warn  {background:rgba(210,153,34,.15);border:1px solid var(--yellow);color:var(--yellow);}
.pill-ok    {background:rgba(63,185,80,.15);border:1px solid var(--green);color:var(--green);}
.pill-off   {background:rgba(125,133,144,.15);border:1px solid var(--muted);color:var(--muted);}

.footer{
  margin-left:220px;padding:14px 28px;border-top:1px solid var(--border);
  color:var(--muted);font-size:11px;display:flex;justify-content:space-between;flex-wrap:wrap;gap:8px;
}

::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:var(--bg);}
::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px;}

@media print{
  .sidenav,.global-bar,.search-box,.top-header{display:none!important;}
  .main,.footer{margin-left:0!important;}
}
</style>
</head>
<body>

<header class="top-header">
  <div class="title-block">
    <h1><em>AD DHCP</em> Report</h1>
    <div class="sub">Created by $ScriptAuthor</div>
  </div>
  <div class="meta-block">
    <div>Generated: <strong>$DateDisplay</strong></div>
    <div>Run By: <strong>$htmlRunByUser</strong></div>
    <div>Domain: <strong>$htmlDomainName</strong></div>
    <div>Forest: <strong>$htmlForestName</strong></div>
  </div>
</header>

<nav class="sidenav">
  <div class="nav-group">Overview</div>
  <a href="#sec-summary">Summary &amp; Health</a>
  <a href="#sec-ad">AD Domain Info</a>
  <a href="#sec-auth">Authorized Servers</a>
  <div class="nav-group">Server Config</div>
  <a href="#sec-srv-settings">Server Settings</a>
  <a href="#sec-srv-stats">Server Statistics</a>
  <a href="#sec-audit">Audit Log</a>
  <a href="#sec-db">Database Settings</a>
  <div class="nav-group">Scopes</div>
  <a href="#sec-health">Scope Health</a>
  <a href="#sec-scope-stats">Scope Statistics</a>
  <a href="#sec-v4scopes">IPv4 Scopes</a>
  <a href="#sec-v6scopes">IPv6 Scopes</a>
  <a href="#sec-superscopes">Superscopes</a>
  <a href="#sec-multicast">Multicast Scopes</a>
  <div class="nav-group">High Availability</div>
  <a href="#sec-failover">Failover</a>
  <div class="nav-group">Policy &amp; Options</div>
  <a href="#sec-policies">DHCP Policies</a>
  <a href="#sec-srv-options">Server Options</a>
  <a href="#sec-scope-options">Scope Options</a>
  <div class="nav-group">Leases &amp; Reservations</div>
  <a href="#sec-reservations">Reservations</a>
  <a href="#sec-exclusions">Exclusion Ranges</a>
  <a href="#sec-leases">Active Leases</a>
</nav>

<main class="main">

  <div class="global-bar">
    <label>&#128269; Global Search:</label>
    <input type="text" id="globalSearch" placeholder="Search all tables..." oninput="globalFilter()"/>
    <span class="match-count" id="matchCount"></span>
    <div class="collapse-controls" style="margin-left:auto;">
      <button class="ctrl-btn" onclick="expandAll()">&#9660; Expand All</button>
      <button class="ctrl-btn" onclick="collapseAll()">&#9658; Collapse All</button>
    </div>
  </div>

  <!-- SUMMARY -->
  <div class="section" id="sec-summary">
    <div class="section-header" onclick="toggleSection('sec-summary')">
      <h2>&#128202; DHCP Environment Summary &amp; Health</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">
      <div class="cards-grid">
        <div class="card c-blue"><div class="card-val">$totalServers</div><div class="card-lbl">DHCP Servers</div></div>
        <div class="card c-cyan"><div class="card-val">$totalV4Scopes</div><div class="card-lbl">IPv4 Scopes</div></div>
        <div class="card c-purple"><div class="card-val">$totalV6Scopes</div><div class="card-lbl">IPv6 Scopes</div></div>
        <div class="card c-red"><div class="card-val">$cntCritical</div><div class="card-lbl">Critical (&ge;90%)</div></div>
        <div class="card c-yellow"><div class="card-val">$cntWarning</div><div class="card-lbl">Warning (&ge;75%)</div></div>
        <div class="card c-green"><div class="card-val">$cntHealthy</div><div class="card-lbl">Healthy Scopes</div></div>
        <div class="card c-muted"><div class="card-val">$cntInactive</div><div class="card-lbl">Inactive Scopes</div></div>
        <div class="card c-teal"><div class="card-val">$totalLeases</div><div class="card-lbl">Active Leases</div></div>
        <div class="card c-orange"><div class="card-val">$totalReserv</div><div class="card-lbl">Reservations</div></div>
        <div class="card c-blue"><div class="card-val">$totalExcl</div><div class="card-lbl">Exclusion Ranges</div></div>
        <div class="card c-green"><div class="card-val">$totalFailover</div><div class="card-lbl">Failover Relations</div></div>
        <div class="card c-purple"><div class="card-val">$totalPolicies</div><div class="card-lbl">DHCP Policies</div></div>
      </div>
      <div class="charts-row">
        <div class="chart-container">
          <h3>Scope Health Distribution</h3>
          <div class="donut-wrap">
            <svg class="donut-svg" width="110" height="110" viewBox="0 0 110 110" id="donutChart"></svg>
            <div class="donut-legend" id="donutLegend"></div>
          </div>
        </div>
        <div class="chart-container">
          <h3>Top Scopes by % Utilization</h3>
          <div class="chart-inner" id="scopeBarChart"></div>
        </div>
      </div>
    </div>
  </div>

  <!-- AD SUMMARY -->
  <div class="section" id="sec-ad">
    <div class="section-header" onclick="toggleSection('sec-ad')">
      <h2>Active Directory Domain Summary</h2><span class="badge">AD</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tADSummary</div>
  </div>

  <!-- AUTHORIZED SERVERS -->
  <div class="section" id="sec-auth">
    <div class="section-header" onclick="toggleSection('sec-auth')">
      <h2>AD-Authorized DHCP Servers</h2><span class="badge">$totalServers</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tAuthSrv</div>
  </div>

  <!-- SERVER SETTINGS -->
  <div class="section" id="sec-srv-settings">
    <div class="section-header" onclick="toggleSection('sec-srv-settings')">
      <h2>DHCP Server Settings</h2><span class="badge">Per Server</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tSrvSettings</div>
  </div>

  <!-- SERVER STATISTICS -->
  <div class="section" id="sec-srv-stats">
    <div class="section-header" onclick="toggleSection('sec-srv-stats')">
      <h2>DHCP Server Statistics</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tSrvStats</div>
  </div>

  <!-- AUDIT LOG -->
  <div class="section" id="sec-audit">
    <div class="section-header" onclick="toggleSection('sec-audit')">
      <h2>Audit Log Settings</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tAuditLog</div>
  </div>

  <!-- DATABASE -->
  <div class="section" id="sec-db">
    <div class="section-header" onclick="toggleSection('sec-db')">
      <h2>DHCP Database Settings</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tDatabase</div>
  </div>

  <!-- SCOPE HEALTH -->
  <div class="section" id="sec-health">
    <div class="section-header" onclick="toggleSection('sec-health')">
      <h2>Scope Health &amp; Utilization</h2>
      <span class="badge red">$cntCritical Critical</span>
      <span class="badge yellow">$cntWarning Warning</span>
      <span class="badge green">$cntHealthy Healthy</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">
      <div class="health-pills">
        <div class="pill pill-crit">&#9632; Critical (&ge;90%): $cntCritical</div>
        <div class="pill pill-warn">&#9632; Warning (&ge;75%): $cntWarning</div>
        <div class="pill pill-ok">&#9632; Healthy: $cntHealthy</div>
        <div class="pill pill-off">&#9632; Inactive: $cntInactive</div>
      </div>
      $tScopeHealth
    </div>
  </div>

  <!-- SCOPE STATISTICS -->
  <div class="section" id="sec-scope-stats">
    <div class="section-header" onclick="toggleSection('sec-scope-stats')">
      <h2>Scope Statistics (Addresses)</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tScopeStats</div>
  </div>

  <!-- IPV4 SCOPES -->
  <div class="section" id="sec-v4scopes">
    <div class="section-header" onclick="toggleSection('sec-v4scopes')">
      <h2>IPv4 Scopes (Full Detail)</h2><span class="badge">$totalV4Scopes Scopes</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tV4Scopes</div>
  </div>

  <!-- IPV6 SCOPES -->
  <div class="section" id="sec-v6scopes">
    <div class="section-header" onclick="toggleSection('sec-v6scopes')">
      <h2>IPv6 Scopes</h2><span class="badge">$totalV6Scopes Scopes</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tV6Scopes</div>
  </div>

  <!-- SUPERSCOPES -->
  <div class="section" id="sec-superscopes">
    <div class="section-header" onclick="toggleSection('sec-superscopes')">
      <h2>Superscopes</h2><span class="badge">$totalSuperscope</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tSuperscopes</div>
  </div>

  <!-- MULTICAST -->
  <div class="section" id="sec-multicast">
    <div class="section-header" onclick="toggleSection('sec-multicast')">
      <h2>Multicast Scopes</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tMulticast</div>
  </div>

  <!-- FAILOVER -->
  <div class="section" id="sec-failover">
    <div class="section-header" onclick="toggleSection('sec-failover')">
      <h2>DHCP Failover Relationships</h2><span class="badge">$totalFailover</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tFailover</div>
  </div>

  <!-- POLICIES -->
  <div class="section" id="sec-policies">
    <div class="section-header" onclick="toggleSection('sec-policies')">
      <h2>DHCP Policies</h2><span class="badge">$totalPolicies</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tPolicies</div>
  </div>

  <!-- SERVER OPTIONS -->
  <div class="section" id="sec-srv-options">
    <div class="section-header" onclick="toggleSection('sec-srv-options')">
      <h2>Server-Level DHCP Options</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tSrvOptions</div>
  </div>

  <!-- SCOPE OPTIONS -->
  <div class="section" id="sec-scope-options">
    <div class="section-header" onclick="toggleSection('sec-scope-options')">
      <h2>Scope-Level DHCP Options</h2>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tScopeOptions</div>
  </div>

  <!-- RESERVATIONS -->
  <div class="section" id="sec-reservations">
    <div class="section-header" onclick="toggleSection('sec-reservations')">
      <h2>DHCP Reservations</h2><span class="badge">$totalReserv</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tReservations</div>
  </div>

  <!-- EXCLUSIONS -->
  <div class="section" id="sec-exclusions">
    <div class="section-header" onclick="toggleSection('sec-exclusions')">
      <h2>Exclusion Ranges</h2><span class="badge">$totalExcl</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tExclusions</div>
  </div>

  <!-- ACTIVE LEASES -->
  <div class="section" id="sec-leases">
    <div class="section-header" onclick="toggleSection('sec-leases')">
      <h2>Active Leases</h2><span class="badge">$totalLeases</span>
      <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><polyline points="6 9 12 15 18 9"/></svg>
    </div>
    <div class="section-body">$tLeases</div>
  </div>

</main>

<footer class="footer">
  <span>$ScriptTitle &mdash; $ScriptAuthor</span>
  <span>Generated: $DateDisplay &nbsp;|&nbsp; By: $htmlRunByUser</span>
</footer>

<script>
// ---- Collapsible Sections ----
function toggleSection(id) {
  var sec = document.getElementById(id);
  var wasCollapsed = sec.classList.contains('collapsed');
  sec.classList.toggle('collapsed');
  try {
    var states = JSON.parse(sessionStorage.getItem('secStates') || '{}');
    states[id] = !wasCollapsed;
    sessionStorage.setItem('secStates', JSON.stringify(states));
  } catch(e) {}
}
function collapseAll() {
  document.querySelectorAll('.section').forEach(function(s) { s.classList.add('collapsed'); });
  try { sessionStorage.removeItem('secStates'); } catch(e) {}
}
function expandAll() {
  document.querySelectorAll('.section').forEach(function(s) { s.classList.remove('collapsed'); });
  try { sessionStorage.removeItem('secStates'); } catch(e) {}
}
(function restoreStates() {
  try {
    var states = JSON.parse(sessionStorage.getItem('secStates') || '{}');
    Object.keys(states).forEach(function(id) {
      if (states[id]) { var el = document.getElementById(id); if (el) el.classList.add('collapsed'); }
    });
  } catch(e) {}
})();
document.querySelectorAll('.sidenav a').forEach(function(a) {
  a.addEventListener('click', function(e) {
    var href = a.getAttribute('href');
    if (href && href.startsWith('#')) {
      var target = document.getElementById(href.slice(1));
      if (target) target.classList.remove('collapsed');
    }
  });
});

// ---- Per-table search ----
function filterTable(input, tid) {
  var f = input.value.toUpperCase();
  var rows = document.getElementById(tid).getElementsByTagName('tr');
  for (var i = 1; i < rows.length; i++) {
    var found = false;
    var cells = rows[i].getElementsByTagName('td');
    for (var j = 0; j < cells.length; j++) {
      if (cells[j].textContent.toUpperCase().indexOf(f) > -1) { found = true; break; }
    }
    rows[i].classList.toggle('hidden', !found);
  }
}

// ---- Global search ----
function globalFilter() {
  var f = document.getElementById('globalSearch').value.toUpperCase();
  var tables = document.querySelectorAll('.data-table');
  var hits = 0;
  tables.forEach(function(t) {
    var rows = t.getElementsByTagName('tr');
    var tableHits = 0;
    for (var i = 1; i < rows.length; i++) {
      var show = !f || rows[i].textContent.toUpperCase().indexOf(f) > -1;
      rows[i].classList.toggle('hidden', !show);
      if (show && f) { hits++; tableHits++; }
    }
    if (f && tableHits > 0) { var sec = t.closest('.section'); if (sec) sec.classList.remove('collapsed'); }
    var wrap = t.closest('.table-wrapper');
    if (wrap && f) { var sb = wrap.querySelector('.search-box'); if (sb) sb.value = ''; }
  });
  var el = document.getElementById('matchCount');
  el.textContent = f ? (hits + ' match' + (hits !== 1 ? 'es' : '')) : '';
}

// ---- Sort ----
function sortTable(tid, col) {
  var t = document.getElementById(tid);
  var rows = Array.from(t.tBodies[0].rows);
  var asc = t.dataset.sc == col && t.dataset.sd == 'asc' ? false : true;
  t.dataset.sc = col; t.dataset.sd = asc ? 'asc' : 'desc';
  rows.sort(function(a, b) {
    var av = a.cells[col] ? a.cells[col].textContent.trim() : '';
    var bv = b.cells[col] ? b.cells[col].textContent.trim() : '';
    var an = parseFloat(av), bn = parseFloat(bv);
    if (!isNaN(an) && !isNaN(bn)) return asc ? an - bn : bn - an;
    return asc ? av.localeCompare(bv) : bv.localeCompare(av);
  });
  rows.forEach(function(r) { t.tBodies[0].appendChild(r); });
}

// ---- Nav scroll highlight ----
window.addEventListener('scroll', function() {
  var secs = document.querySelectorAll('.section');
  var links = document.querySelectorAll('.sidenav a');
  var cur = '';
  secs.forEach(function(s) { if (s.getBoundingClientRect().top <= 130) cur = s.id; });
  links.forEach(function(a) { a.classList.toggle('active', a.getAttribute('href') === '#' + cur); });
});

// ---- Donut chart ----
(function() {
  var vals   = [$chartCounts];
  var colors = [$chartColors];
  var labels = ['Critical','Warning','Healthy','Inactive'];
  var total  = vals.reduce(function(a,b){ return a+b; }, 0);
  if (total === 0) return;
  var svg    = document.getElementById('donutChart');
  var legend = document.getElementById('donutLegend');
  var cx = 55, cy = 55, r = 38, innerR = 22;
  var startAngle = -Math.PI / 2;
  vals.forEach(function(val, i) {
    if (val === 0) return;
    var slice = (val / total) * 2 * Math.PI;
    var endAngle = startAngle + slice;
    var x1 = cx + r * Math.cos(startAngle), y1 = cy + r * Math.sin(startAngle);
    var x2 = cx + r * Math.cos(endAngle),   y2 = cy + r * Math.sin(endAngle);
    var xi1= cx + innerR * Math.cos(endAngle),   yi1= cy + innerR * Math.sin(endAngle);
    var xi2= cx + innerR * Math.cos(startAngle), yi2= cy + innerR * Math.sin(startAngle);
    var large = slice > Math.PI ? 1 : 0;
    var path = document.createElementNS('http://www.w3.org/2000/svg','path');
    path.setAttribute('d', 'M'+x1+','+y1+' A'+r+','+r+' 0 '+large+',1 '+x2+','+y2+' L'+xi1+','+yi1+' A'+innerR+','+innerR+' 0 '+large+',0 '+xi2+','+yi2+' Z');
    path.setAttribute('fill', colors[i]);
    path.setAttribute('opacity','0.9');
    svg.appendChild(path);
    startAngle = endAngle;
    var item = document.createElement('div'); item.className = 'legend-item';
    var dot  = document.createElement('div'); dot.className  = 'legend-dot'; dot.style.background = colors[i];
    var txt  = document.createElement('span'); txt.textContent = labels[i] + ': ' + val;
    item.appendChild(dot); item.appendChild(txt); legend.appendChild(item);
  });
  var pct = total > 0 ? Math.round((vals[2]/total)*100) : 0;
  var txt = document.createElementNS('http://www.w3.org/2000/svg','text');
  txt.setAttribute('x', cx); txt.setAttribute('y', cy+4);
  txt.setAttribute('text-anchor','middle'); txt.setAttribute('font-size','13');
  txt.setAttribute('font-weight','700'); txt.setAttribute('fill','#e6edf3');
  txt.textContent = pct + '%';
  svg.appendChild(txt);
  var sub = document.createElementNS('http://www.w3.org/2000/svg','text');
  sub.setAttribute('x', cx); sub.setAttribute('y', cy+16);
  sub.setAttribute('text-anchor','middle'); sub.setAttribute('font-size','8');
  sub.setAttribute('fill','#7d8590'); sub.textContent = 'healthy';
  svg.appendChild(sub);
})();

// ---- Scope utilization bar chart ----
(function() {
  var labels = [$scopeChartLabels];
  var counts = [$scopeChartCounts];
  if (!labels.length || !counts.length) return;
  var max = Math.max.apply(null, counts);
  if (max === 0) return;
  var container = document.getElementById('scopeBarChart');
  labels.forEach(function(lbl, i) {
    var pct = (counts[i] / 100) * 100;
    var color = counts[i] >= 90 ? '#f85149' : counts[i] >= 75 ? '#d29922' : '#3fb950';
    var wrap = document.createElement('div'); wrap.className = 'bar-wrap';
    var cntEl = document.createElement('div'); cntEl.className = 'bar-count'; cntEl.textContent = counts[i] + '%';
    var bar = document.createElement('div'); bar.className = 'bar';
    bar.style.height = pct + '%';
    bar.style.background = color;
    bar.title = lbl + ': ' + counts[i] + '%';
    var lblEl = document.createElement('div'); lblEl.className = 'bar-lbl'; lblEl.textContent = lbl;
    wrap.appendChild(cntEl); wrap.appendChild(bar); wrap.appendChild(lblEl);
    container.appendChild(wrap);
  });
})();
</script>
</body>
</html>
"@

try {
    [System.IO.File]::WriteAllText($HTMLPath, $HTML, [System.Text.Encoding]::UTF8)
    if (Test-Path $HTMLPath) {
        $htmlSize = (Get-Item $HTMLPath).Length
        Write-OK "HTML report saved: $HTMLPath ($htmlSize bytes)"
        Write-Info "Opening HTML report in default browser..."
        Start-Process $HTMLPath
    } else {
        Write-Fail "HTML file not found after write attempt: $HTMLPath"
    }
} catch {
    Write-Fail "HTML WriteAllText FAILED: $_"
    Write-Fail "Attempting fallback write via Set-Content..."
    try {
        $HTML | Set-Content -Path $HTMLPath -Encoding UTF8 -Force
        Write-OK "HTML fallback write succeeded: $HTMLPath"
        Write-Info "Opening HTML report in default browser..."
        Start-Process $HTMLPath
    } catch {
        Write-Fail "HTML fallback also failed: $_"
    }
}
#endregion

#region --- FINAL SUMMARY ---
Write-Host ""
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  REPORT COMPLETE" -ForegroundColor Green
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  Output Folder : $OutputFolder" -ForegroundColor Yellow
Write-Host ""
Write-Host "  Files Generated:" -ForegroundColor White
Get-ChildItem -Path $OutputFolder | ForEach-Object {
    Write-Host ("    {0,-60} {1,8}" -f $_.Name, ("{0:N0} KB" -f ($_.Length/1KB))) -ForegroundColor Gray
}
Write-Host ""
Write-Host "  DHCP Summary:" -ForegroundColor White
Write-Host "    DHCP Servers           : $totalServers"
Write-Host "    IPv4 Scopes            : $totalV4Scopes"
Write-Host "    IPv6 Scopes            : $totalV6Scopes"
Write-Host "    Critical Scopes (>=90%): $cntCritical"
Write-Host "    Warning Scopes  (>=75%): $cntWarning"
Write-Host "    Healthy Scopes         : $cntHealthy"
Write-Host "    Inactive Scopes        : $cntInactive"
Write-Host "    Active Leases          : $totalLeases"
Write-Host "    Reservations           : $totalReserv"
Write-Host "    Exclusion Ranges       : $totalExcl"
Write-Host "    Failover Relationships : $totalFailover"
Write-Host "    DHCP Policies          : $totalPolicies"
Write-Host ""
Write-Host "  Opening output folder..." -ForegroundColor Gray
Start-Process explorer.exe $OutputFolder
Write-Host "================================================================" -ForegroundColor Cyan
#endregion
