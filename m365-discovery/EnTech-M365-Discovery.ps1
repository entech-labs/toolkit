<#
.SYNOPSIS
    Read-only Microsoft 365 tenant discovery, for sizing a tenant-to-tenant migration.

.DESCRIPTION
    Collects the facts needed to estimate migration effort - which workloads are in use,
    how much data sits in each, and which structures (shared mailboxes, private channels,
    legal holds, SSO integrations) drive the hard parts of the work.

    WHAT THIS SCRIPT DOES NOT DO - and how to verify each claim yourself:

      * It makes no change to Microsoft 365. Every call into your tenant is a read. Search
        this file for "Invoke-MgGraphRequest": there are four, and all four specify
        -Method GET. Every Exchange Online cmdlet used is a Get-* (Get-EXOMailbox,
        Get-TransportRule, Get-InboundConnector, Get-OutboundConnector,
        Get-OrganizationConfig). There is no Set-, Update-, or Disable- cmdlet in the file.

      * The only New-Item and Remove-Item calls act on your own disk, not on the tenant -
        they create the output folder and replace a previous zip of the same name. Search
        for them and confirm.

      * It reads no message content, no file content and no document bodies. It reads
        counts, sizes, names and configuration only.

      * It transmits nothing. Output is written to a folder on the machine you run it on.
        You review it, and you decide whether to send it.

      * It stores no credentials. Sign-in is interactive, through Microsoft's own prompt,
        and the session is disconnected at the end.

    The reviewer-friendly version of all of the above is in the companion document,
    "What the discovery script does" - share that with whoever has to approve running this.

.PARAMETER OutputPath
    Folder for the results. Default: a timestamped folder on your Desktop.

.PARAMETER SkipExchange
    Skip the Exchange Online section. Use if the Exchange module cannot be installed, or
    if the account running this does not hold an Exchange role. The rest still runs.

.PARAMETER SkipPdf
    Skip PDF generation and leave the HTML report only.

.PARAMETER ReportPeriod
    Reporting window for usage figures: D7, D30, D90 or D180. Default D30.
    D30 is the sweet spot - long enough to catch monthly-cycle users, short enough that
    recently departed staff do not inflate the counts.

.EXAMPLE
    .\EnTech-M365-Discovery.ps1

    Runs everything, writes to the Desktop, opens the report when finished.

.EXAMPLE
    .\EnTech-M365-Discovery.ps1 -OutputPath 'C:\Temp\Discovery' -ReportPeriod D90

    Writes to a specific folder using a 90-day usage window.

.NOTES
    Prepared by EnTech Engineering, P.C.
    Requires PowerShell 5.1 or PowerShell 7+, and a Microsoft 365 account that can read
    directory and reporting data. See the companion document for the exact permissions
    and why each one is needed.
#>

[CmdletBinding()]
param(
    [string] $OutputPath,
    [switch] $SkipExchange,
    [switch] $SkipPdf,
    [ValidateSet('D7','D30','D90','D180')]
    [string] $ReportPeriod = 'D30',

    # Builds the report from representative sample data without touching any tenant and
    # without signing in. Used to verify the report and PDF pipeline end to end before the
    # script is sent to anyone. It is not a simulation of your environment.
    [switch] $SelfTest
)

#-------------------------------------------------------------------------------------------
# REGION 0a - Embedded EnTech logo
#   Base64 PNG used in the report header. It has to be assigned before the report is built,
#   so it lives here rather than at the end of the file. Nothing below depends on reading it:
#   collapse this region and carry on.
#-------------------------------------------------------------------------------------------
#region Logo
$script:LogoBase64 =
    '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' +
    '70YAAAAASUVORK5CYII='
#endregion

#-------------------------------------------------------------------------------------------
# REGION 0 - Setup
#-------------------------------------------------------------------------------------------
#region Setup

$ErrorActionPreference = 'Stop'          # fail loudly inside try blocks rather than limp on
$script:StartedAt      = Get-Date
$script:Findings       = [ordered]@{}    # every collector drops its results in here
$script:Issues         = New-Object System.Collections.Generic.List[string]
$script:Brand          = '#1d4e5f'

if (-not $OutputPath) {
    $stamp      = Get-Date -Format 'yyyy-MM-dd_HHmm'
    $OutputPath = Join-Path ([Environment]::GetFolderPath('Desktop')) "M365-Discovery_$stamp"
}
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }
$CsvPath = Join-Path $OutputPath 'data'
if (-not (Test-Path $CsvPath)) { New-Item -ItemType Directory -Path $CsvPath -Force | Out-Null }

function Write-Step   { param($m) Write-Host "`n>> $m" -ForegroundColor Cyan }
function Write-Ok     { param($m) Write-Host "   $m"   -ForegroundColor Green }
function Write-Note   { param($m) Write-Host "   $m"   -ForegroundColor Gray  }

# A collector that fails should cost us that one section, not the whole run. Anything that
# goes wrong is recorded and surfaced in the report, because a silently missing section is
# worse than a visible gap - the estimate would be built on a hole nobody knew about.
function Add-Issue {
    param([string] $Section, [string] $Message)
    $script:Issues.Add("$Section - $Message")
    Write-Host "   ! $Section - $Message" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "  EnTech Engineering - Microsoft 365 Discovery" -ForegroundColor White
Write-Host "  Read-only. Nothing is changed and nothing is sent anywhere." -ForegroundColor Gray
Write-Host "  Output: $OutputPath" -ForegroundColor Gray

#endregion

#-------------------------------------------------------------------------------------------
# REGION 0b - Self-test
#   -SelfTest fills the findings with obviously-fictional sample data and skips straight to
#   the report. It exists so the report and PDF pipeline can be verified end to end without
#   touching any tenant. Everything between here and REGION 5 is skipped when it is set.
#-------------------------------------------------------------------------------------------
if ($SelfTest) {

    Write-Host "`n  SELF-TEST - sample data, no tenant is contacted" -ForegroundColor Magenta

    $script:Findings.Tenant = [ordered]@{
        DisplayName = 'Fabrikam Example Agency (SAMPLE DATA)'; TenantId = '00000000-0000-0000-0000-000000000000'
        Country = 'US'; DefaultDomain = 'example.onmicrosoft.com'
        DomainsTotal = 3; DomainsVerified = 3; DomainsCustom = 2
        DirSyncEnabled = $true; DirSyncLastRun = (Get-Date).AddMinutes(-22)
    }
    $script:Findings.ReportsAnonymised = $false
    $script:Findings.Licensing = @{
        SkuCount = 2; TotalAssigned = 54
        Rows = @(
            [pscustomobject]@{ SkuPartNumber='SPB'; Purchased=55; Assigned=51; Available=4; Warning=0 }
            [pscustomobject]@{ SkuPartNumber='EXCHANGESTANDARD'; Purchased=5; Assigned=3; Available=2; Warning=0 }
        )
    }
    $script:Findings.Accounts   = [ordered]@{ Total=54; Enabled=47; Disabled=7; Guests=6; Licensed=24; DirSynced=31 }
    $script:Findings.Mailboxes  = [ordered]@{
        Count=38; TotalGB=412.6; AverageGB=10.86; TotalItems=1284310; Over10GB=9; Over50GB=2
        Largest = @(
            [pscustomobject]@{ Mailbox='info@example.org';    SizeGB=78.4; Items=201443; LastWrite='2026-10-01' }
            [pscustomobject]@{ Mailbox='a.planner@example.org'; SizeGB=61.2; Items=154002; LastWrite='2026-10-02' }
            [pscustomobject]@{ Mailbox='records@example.org';  SizeGB=33.9; Items=88120;  LastWrite='2026-09-30' }
        )
    }
    $script:Findings.OneDrive   = [ordered]@{
        Sites=24; TotalGB=286.1; TotalFiles=312004; Empty=5
        Largest = @(
            [pscustomobject]@{ Owner='gis.analyst@example.org'; SizeGB=96.7; Files=41203 }
            [pscustomobject]@{ Owner='director@example.org';    SizeGB=38.2; Files=22810 }
        )
    }
    $script:Findings.SharePoint = [ordered]@{
        Sites=14; TotalGB=503.8; TotalFiles=486221; Over100GB=1
        Largest = @(
            [pscustomobject]@{ Site='https://example.sharepoint.com/sites/Planning'; SizeGB=188.4; Files=201884; Template='STS' }
            [pscustomobject]@{ Site='https://example.sharepoint.com/sites/Projects'; SizeGB=92.1;  Files=104221; Template='STS' }
        )
    }
    $script:Findings.SharePointShape = [ordered]@{
        SitesInspected=14; DocumentLibs=31; CustomLists=7; SitesWithLists=3; FileStorageOnly=$false
        Detail = @(
            [pscustomobject]@{ Site='https://example.sharepoint.com/sites/Planning'; Libraries=6; CustomLists=4; ListNames='Project Tracker; TIP Submissions; Meeting Actions; Vendor Log' }
            [pscustomobject]@{ Site='https://example.sharepoint.com/sites/Intranet'; Libraries=3; CustomLists=3; ListNames='Staff Directory; Policies; Announcements' }
        )
    }
    $script:Findings.Groups = [ordered]@{
        Total=29; M365Groups=12; Distribution=11; Security=6; Teams=9; PrivateTeams=4; PrivateChannels=3
    }
    $script:Findings.Applications = [ordered]@{
        Total=118; ThirdParty=17; WithSso=6; ConditionalAccessPolicies=4; CaPoliciesEnabled=3
    }
    $script:Findings.Exchange = [ordered]@{
        UserMailboxes=21; SharedMailboxes=14; RoomMailboxes=2; EquipmentMailbox=1
        ArchivesEnabled=6; LitigationHold=2; TransportRules=9; Connectors=2; PublicFolders='None'
    }
    $script:Findings.MailboxSplit = @(
        [pscustomobject]@{ Type='UserMailbox';   Count=21; TotalGB=238.4 }
        [pscustomobject]@{ Type='SharedMailbox'; Count=14; TotalGB=168.9 }
        [pscustomobject]@{ Type='RoomMailbox';   Count=2;  TotalGB=3.1 }
    )
    $script:Findings.NonUserMailboxes = @(
        [pscustomobject]@{ Mailbox='info@example.org';    Type='SharedMailbox'; SizeGB=78.4 }
        [pscustomobject]@{ Mailbox='records@example.org'; Type='SharedMailbox'; SizeGB=33.9 }
        [pscustomobject]@{ Mailbox='scanner@example.org'; Type='SharedMailbox'; SizeGB=11.2 }
    )
    Add-Issue 'Self-test' 'This report was produced from sample data, not from a live tenant'

} else {

#-------------------------------------------------------------------------------------------
# REGION 1 - Module preflight
#-------------------------------------------------------------------------------------------
#region Modules

# Only two modules are needed. We deliberately avoid the retired MSOnline and AzureAD
# modules: Microsoft stopped supporting them in 2026 and they no longer return complete
# data. Microsoft.Graph.Authentication gives us a signed-in Graph session, and every Graph
# call below goes through Invoke-MgGraphRequest against the REST API directly - that keeps
# the dependency surface to one module instead of the forty-odd Microsoft.Graph.* submodules,
# and it does not break when the SDK reshapes its cmdlets between versions.

Write-Step 'Checking prerequisites'

function Confirm-Module {
    param([string] $Name, [string] $Why)

    if (Get-Module -ListAvailable -Name $Name) {
        Write-Ok "$Name is present"
        return $true
    }

    Write-Host "   $Name is not installed." -ForegroundColor Yellow
    Write-Host "   Needed for: $Why" -ForegroundColor Gray
    $answer = Read-Host "   Install it now for the current user only? (Y/N)"

    if ($answer -notmatch '^[Yy]') {
        Add-Issue $Name 'Declined installation - the matching section will be skipped'
        return $false
    }

    try {
        Install-Module $Name -Scope CurrentUser -Force -AllowClobber
        Write-Ok "$Name installed"
        return $true
    } catch {
        Add-Issue $Name "Installation failed: $($_.Exception.Message)"
        return $false
    }
}

$hasGraph = Confirm-Module -Name 'Microsoft.Graph.Authentication' `
                           -Why  'tenant, users, licensing, usage reports, Teams, SharePoint'

$hasExo = $false
if (-not $SkipExchange) {
    $hasExo = Confirm-Module -Name 'ExchangeOnlineManagement' `
                             -Why  'shared mailboxes, public folders, mail flow rules, legal holds'
}

if (-not $hasGraph) {
    Write-Host "`n  Cannot continue without Microsoft.Graph.Authentication." -ForegroundColor Red
    exit 1
}

#endregion

#-------------------------------------------------------------------------------------------
# REGION 2 - Sign in
#-------------------------------------------------------------------------------------------
#region Connect

# Every scope below is a READ scope. Microsoft will show this exact list on the consent
# prompt. If your organisation requires admin consent, an administrator approves the list
# once; the script never stores a token or a credential of any kind.
$GraphScopes = @(
    'Organization.Read.All'        # tenant name, verified domains, service plans
    'User.Read.All'                # account counts, types, licence assignment
    'Group.Read.All'               # Microsoft 365 groups, distribution lists, Teams backing groups
    'Directory.Read.All'           # directory-wide counts, domain configuration
    'Sites.Read.All'               # SharePoint site inventory
    'Reports.Read.All'             # storage and usage figures - the sizing numbers
    'ReportSettings.Read.All'      # detects whether reports are anonymised (see below)
    'Application.Read.All'         # app registrations and SSO integrations to rebuild
    'Policy.Read.All'              # conditional access and authentication policies
    'TeamSettings.Read.All'        # Teams configuration
)

Write-Step 'Signing in to Microsoft Graph'
Write-Note 'A Microsoft sign-in window will open. Read-only access is requested.'

try {
    Connect-MgGraph -Scopes $GraphScopes -NoWelcome
    $ctx = Get-MgContext
    Write-Ok "Signed in as $($ctx.Account)"
} catch {
    Write-Host "   Sign-in failed: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

#endregion

#-------------------------------------------------------------------------------------------
# REGION 3 - Graph helpers
#-------------------------------------------------------------------------------------------
#region Helpers

<#
    Graph pages its collections. Anything that returns a list has to be followed through
    @odata.nextLink or you silently get the first 100 rows and a wrong total - which is
    exactly the kind of error that produces a confident, incorrect estimate.
#>
function Get-GraphAll {
    param([Parameter(Mandatory)][string] $Uri)

    $all = New-Object System.Collections.Generic.List[object]
    $next = $Uri

    while ($next) {
        $page = Invoke-MgGraphRequest -Method GET -Uri $next -OutputType PSObject
        if ($page.value) { $page.value | ForEach-Object { $all.Add($_) } }
        $next = $page.'@odata.nextLink'
    }
    return $all
}

<#
    The Microsoft 365 usage reports are the authoritative source for storage sizes, and they
    are the fast one: one call returns every mailbox or site rather than a per-object loop.
    They return CSV rather than JSON, so they need handling of their own.
#>
function Get-GraphReportCsv {
    param([Parameter(Mandatory)][string] $Report)

    $uri  = "https://graph.microsoft.com/v1.0/reports/$Report(period='$ReportPeriod')"
    $resp = Invoke-MgGraphRequest -Method GET -Uri $uri -OutputType HttpResponseMessage
    $text = $resp.Content.ReadAsStringAsync().Result

    # Reports come back with a byte-order mark that would otherwise corrupt the first
    # column name and make every lookup against it fail.
    $text = $text -replace '^\xEF\xBB\xBF', '' -replace '^\uFEFF', ''

    if ([string]::IsNullOrWhiteSpace($text)) { return @() }
    return ($text | ConvertFrom-Csv)
}

function Save-Csv {
    param([string] $Name, $Data)
    if (-not $Data -or @($Data).Count -eq 0) { return }
    $file = Join-Path $CsvPath "$Name.csv"
    $Data | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Write-Note "saved data\$Name.csv ($(@($Data).Count) rows)"
}

# Sizes arrive as bytes. Humans estimate in gigabytes.
function ConvertTo-GB {
    param($Bytes)
    if (-not $Bytes) { return 0 }
    return [math]::Round(([double]$Bytes) / 1GB, 2)
}

#endregion

#-------------------------------------------------------------------------------------------
# REGION 4 - Collection
#-------------------------------------------------------------------------------------------

#region 4.1 Tenant and domains

Write-Step 'Tenant and domains'
try {
    $org = (Invoke-MgGraphRequest -Method GET -OutputType PSObject `
            -Uri 'https://graph.microsoft.com/v1.0/organization').value | Select-Object -First 1

    $domains = Get-GraphAll 'https://graph.microsoft.com/v1.0/domains'

    # Verified custom domains are the single biggest scheduling constraint in a tenant move:
    # a domain can only live in one tenant at a time, so every one of these has to be
    # released and re-verified inside the cutover window.
    $script:Findings.Tenant = [ordered]@{
        DisplayName     = $org.displayName
        TenantId        = $org.id
        Country         = $org.countryLetterCode
        DefaultDomain   = ($domains | Where-Object { $_.isDefault }).id
        DomainsTotal    = @($domains).Count
        DomainsVerified = @($domains | Where-Object { $_.isVerified }).Count
        DomainsCustom   = @($domains | Where-Object { $_.id -notlike '*.onmicrosoft.com' }).Count
        DirSyncEnabled  = [bool]$org.onPremisesSyncEnabled
        DirSyncLastRun  = $org.onPremisesLastSyncDateTime
    }
    Save-Csv 'domains' ($domains | Select-Object id, isDefault, isVerified, authenticationType, supportedServices)
    Write-Ok "$($org.displayName) - $(@($domains).Count) domains"

    if ($org.onPremisesSyncEnabled) {
        Write-Note 'Directory synchronisation is ON - identities originate on-premises.'
    }
} catch {
    Add-Issue 'Tenant' $_.Exception.Message
}

#endregion

#region 4.2 Report anonymisation check

<#
    Microsoft 365 can be configured to conceal user, group and site names in all reports.
    If that setting is on, every name below arrives as a meaningless hash. The data is still
    usable for totals but useless for "which mailboxes are the big ones", so we check first
    and say so plainly rather than handing over a report full of hashes.
#>
Write-Step 'Checking report privacy setting'
try {
    $rs = Invoke-MgGraphRequest -Method GET -OutputType PSObject `
          -Uri 'https://graph.microsoft.com/v1.0/admin/reportSettings'

    $concealed = [bool]$rs.displayConcealedNames
    $script:Findings.ReportsAnonymised = $concealed

    if ($concealed) {
        Write-Host '   ! Reports are anonymised in this tenant.' -ForegroundColor Yellow
        Write-Note 'Totals will be correct; per-user and per-site names will be hashes.'
        Write-Note 'A Global Administrator can turn this off temporarily in the M365 admin'
        Write-Note 'centre under Settings > Org settings > Reports, then re-run this script.'
    } else {
        Write-Ok 'Reports show real names'
    }
} catch {
    $script:Findings.ReportsAnonymised = $null
    Add-Issue 'ReportSettings' "Could not read the setting: $($_.Exception.Message)"
}

#endregion

#region 4.3 Licensing

Write-Step 'Licensing'
try {
    $skus = Get-GraphAll 'https://graph.microsoft.com/v1.0/subscribedSkus'

    $licRows = $skus | ForEach-Object {
        [pscustomobject]@{
            SkuPartNumber = $_.skuPartNumber
            Purchased     = $_.prepaidUnits.enabled
            Assigned      = $_.consumedUnits
            Available     = [int]$_.prepaidUnits.enabled - [int]$_.consumedUnits
            Warning       = $_.prepaidUnits.warning
        }
    } | Sort-Object Assigned -Descending

    $script:Findings.Licensing = @{
        SkuCount      = @($skus).Count
        TotalAssigned = ($licRows | Measure-Object Assigned -Sum).Sum
        Rows          = $licRows
    }
    Save-Csv 'licensing' $licRows
    Write-Ok "$(@($skus).Count) licence types, $(($licRows | Measure-Object Assigned -Sum).Sum) seats assigned"
} catch {
    Add-Issue 'Licensing' $_.Exception.Message
}

#endregion

#region 4.4 Accounts

Write-Step 'Accounts'
try {
    # $select keeps the payload small; without it Graph returns every property on every user,
    # which on a slow link is the difference between seconds and minutes.
    $users = Get-GraphAll ('https://graph.microsoft.com/v1.0/users?$top=999&$select=' +
             'id,displayName,userPrincipalName,userType,accountEnabled,assignedLicenses,' +
             'onPremisesSyncEnabled,createdDateTime,signInActivity')

    $script:Findings.Accounts = [ordered]@{
        Total        = @($users).Count
        Enabled      = @($users | Where-Object { $_.accountEnabled }).Count
        Disabled     = @($users | Where-Object { -not $_.accountEnabled }).Count
        Guests       = @($users | Where-Object { $_.userType -eq 'Guest' }).Count
        Licensed     = @($users | Where-Object { @($_.assignedLicenses).Count -gt 0 }).Count
        DirSynced    = @($users | Where-Object { $_.onPremisesSyncEnabled }).Count
    }

    Save-Csv 'users' ($users | Select-Object displayName, userPrincipalName, userType,
                      accountEnabled, onPremisesSyncEnabled, createdDateTime)
    Write-Ok "$(@($users).Count) accounts - $($script:Findings.Accounts.Licensed) licensed, $($script:Findings.Accounts.Guests) guests"
} catch {
    Add-Issue 'Accounts' $_.Exception.Message
}

#endregion

#region 4.5 Exchange Online - sizing

Write-Step 'Mailboxes (sizing)'
try {
    $mbx = Get-GraphReportCsv 'getMailboxUsageDetail'

    if (@($mbx).Count -gt 0) {
        $sizeCol = ($mbx[0].PSObject.Properties.Name | Where-Object { $_ -match 'Storage Used' }) |
                   Select-Object -First 1
        $itemCol = ($mbx[0].PSObject.Properties.Name | Where-Object { $_ -match 'Item Count' }) |
                   Select-Object -First 1

        $totalBytes = ($mbx | Measure-Object -Property $sizeCol -Sum).Sum
        $totalItems = ($mbx | Measure-Object -Property $itemCol -Sum).Sum

        # The largest mailboxes drive the migration schedule far more than the average does:
        # a single 90 GB mailbox can take longer than two hundred small ones.
        $top = $mbx | Sort-Object { [double]$_.$sizeCol } -Descending | Select-Object -First 15 |
               ForEach-Object {
                   [pscustomobject]@{
                       Mailbox   = $_.'User Principal Name'
                       SizeGB    = ConvertTo-GB $_.$sizeCol
                       Items     = $_.$itemCol
                       LastWrite = $_.'Last Activity Date'
                   }
               }

        $script:Findings.Mailboxes = [ordered]@{
            Count       = @($mbx).Count
            TotalGB     = ConvertTo-GB $totalBytes
            AverageGB   = if (@($mbx).Count) { [math]::Round((ConvertTo-GB $totalBytes) / @($mbx).Count, 2) } else { 0 }
            TotalItems  = $totalItems
            Over10GB    = @($mbx | Where-Object { (ConvertTo-GB $_.$sizeCol) -gt 10 }).Count
            Over50GB    = @($mbx | Where-Object { (ConvertTo-GB $_.$sizeCol) -gt 50 }).Count
            Largest     = $top
        }
        # Held for the shared-versus-user split in the Exchange section, which needs the
        # recipient type that this report does not carry.
        $script:MailboxRaw     = $mbx
        $script:MailboxSizeCol = $sizeCol

        Save-Csv 'mailbox-usage' $mbx
        Write-Ok "$(@($mbx).Count) mailboxes, $(ConvertTo-GB $totalBytes) GB total"
    }
} catch {
    Add-Issue 'Mailbox usage' $_.Exception.Message
}

#endregion

#region 4.6 OneDrive

Write-Step 'OneDrive'
try {
    $od = Get-GraphReportCsv 'getOneDriveUsageAccountDetail'

    if (@($od).Count -gt 0) {
        $sizeCol = ($od[0].PSObject.Properties.Name | Where-Object { $_ -match 'Storage Used' }) | Select-Object -First 1
        $fileCol = ($od[0].PSObject.Properties.Name | Where-Object { $_ -match '^File Count' }) | Select-Object -First 1

        $totalBytes = ($od | Measure-Object -Property $sizeCol -Sum).Sum

        $top = $od | Sort-Object { [double]$_.$sizeCol } -Descending | Select-Object -First 15 |
               ForEach-Object {
                   [pscustomobject]@{
                       Owner  = $_.'Owner Principal Name'
                       SizeGB = ConvertTo-GB $_.$sizeCol
                       Files  = $_.$fileCol
                   }
               }

        $script:Findings.OneDrive = [ordered]@{
            Sites      = @($od).Count
            TotalGB    = ConvertTo-GB $totalBytes
            TotalFiles = ($od | Measure-Object -Property $fileCol -Sum).Sum
            Empty      = @($od | Where-Object { [double]$_.$sizeCol -lt 1MB }).Count
            Largest    = $top
        }
        Save-Csv 'onedrive-usage' $od
        Write-Ok "$(@($od).Count) OneDrive accounts, $(ConvertTo-GB $totalBytes) GB"
    }
} catch {
    Add-Issue 'OneDrive' $_.Exception.Message
}

#endregion

#region 4.7 SharePoint

Write-Step 'SharePoint'
try {
    $sp = Get-GraphReportCsv 'getSharePointSiteUsageDetail'

    if (@($sp).Count -gt 0) {
        $sizeCol = ($sp[0].PSObject.Properties.Name | Where-Object { $_ -match 'Storage Used' }) | Select-Object -First 1
        $fileCol = ($sp[0].PSObject.Properties.Name | Where-Object { $_ -match '^File Count' }) | Select-Object -First 1

        $totalBytes = ($sp | Measure-Object -Property $sizeCol -Sum).Sum

        $top = $sp | Sort-Object { [double]$_.$sizeCol } -Descending | Select-Object -First 15 |
               ForEach-Object {
                   [pscustomobject]@{
                       Site     = $_.'Site URL'
                       SizeGB   = ConvertTo-GB $_.$sizeCol
                       Files    = $_.$fileCol
                       Template = $_.'Root Web Template'
                   }
               }

        $script:Findings.SharePoint = [ordered]@{
            Sites      = @($sp).Count
            TotalGB    = ConvertTo-GB $totalBytes
            TotalFiles = ($sp | Measure-Object -Property $fileCol -Sum).Sum
            Over100GB  = @($sp | Where-Object { (ConvertTo-GB $_.$sizeCol) -gt 100 }).Count
            Largest    = $top
        }
        Save-Csv 'sharepoint-usage' $sp
        Write-Ok "$(@($sp).Count) sites, $(ConvertTo-GB $totalBytes) GB"
    }
} catch {
    Add-Issue 'SharePoint' $_.Exception.Message
}

#endregion

#region 4.7b SharePoint composition

<#
    Size alone does not tell us how hard SharePoint is to move. A site that is only document
    libraries is a copy job. A site carrying custom lists, pages or forms is a small
    application, and applications are rebuilt by hand in the destination rather than migrated.
    Counting lists by template is the cheapest reliable way to tell the two apart.
#>
Write-Step 'SharePoint composition'
try {
    $sites = Get-GraphAll 'https://graph.microsoft.com/v1.0/sites?search=*'

    $libraries   = 0
    $customLists = 0
    $siteRows    = New-Object System.Collections.Generic.List[object]
    $listErrors  = 0

    foreach ($site in $sites) {
        try {
            $lists = Get-GraphAll "https://graph.microsoft.com/v1.0/sites/$($site.id)/lists"

            # Hidden system lists exist on every site and would drown the real signal.
            $visible = $lists | Where-Object { -not $_.list.hidden }
            $docLibs = @($visible | Where-Object { $_.list.template -eq 'documentLibrary' })
            $custom  = @($visible | Where-Object { $_.list.template -eq 'genericList' })

            $libraries   += $docLibs.Count
            $customLists += $custom.Count

            if ($custom.Count -gt 0) {
                $siteRows.Add([pscustomobject]@{
                    Site        = $site.webUrl
                    Libraries   = $docLibs.Count
                    CustomLists = $custom.Count
                    ListNames   = (($custom | Select-Object -First 6).displayName -join '; ')
                })
            }
        } catch { $listErrors++ }
    }

    $script:Findings.SharePointShape = [ordered]@{
        SitesInspected  = @($sites).Count
        DocumentLibs    = $libraries
        CustomLists     = $customLists
        SitesWithLists  = $siteRows.Count
        Detail          = $siteRows
        FileStorageOnly = ($customLists -eq 0)
    }
    Save-Csv 'sharepoint-custom-lists' $siteRows
    if ($listErrors) { Add-Issue 'SharePoint lists' "$listErrors site(s) could not be inspected" }

    if ($customLists -eq 0) {
        Write-Ok "$libraries document libraries, no custom lists - file storage only"
    } else {
        Write-Ok "$libraries libraries and $customLists custom list(s) across $($siteRows.Count) site(s)"
        Write-Note 'Custom lists mean SharePoint is more than file storage here.'
    }
} catch {
    Add-Issue 'SharePoint composition' $_.Exception.Message
}

#endregion

#region 4.8 Teams and groups

Write-Step 'Teams and groups'
try {
    $groups = Get-GraphAll ('https://graph.microsoft.com/v1.0/groups?$top=999&$select=' +
              'id,displayName,groupTypes,mailEnabled,securityEnabled,resourceProvisioningOptions,visibility,createdDateTime')

    # A group is Teams-enabled when resourceProvisioningOptions contains "Team". This is the
    # reliable test - a team always has a backing Microsoft 365 group, and reading it this
    # way needs no Teams-specific permission.
    $teams = $groups | Where-Object { $_.resourceProvisioningOptions -contains 'Team' }

    $script:Findings.Groups = [ordered]@{
        Total        = @($groups).Count
        M365Groups   = @($groups | Where-Object { $_.groupTypes -contains 'Unified' }).Count
        Distribution = @($groups | Where-Object { $_.mailEnabled -and -not $_.securityEnabled }).Count
        Security     = @($groups | Where-Object { $_.securityEnabled -and -not $_.mailEnabled }).Count
        Teams        = @($teams).Count
        PrivateTeams = @($teams | Where-Object { $_.visibility -eq 'Private' }).Count
    }
    Save-Csv 'groups' ($groups | Select-Object displayName, visibility, mailEnabled,
                       securityEnabled, createdDateTime, @{n='IsTeam';e={$_.resourceProvisioningOptions -contains 'Team'}})

    # Private channels get their own hidden SharePoint site each, and most migration tooling
    # handles them badly or not at all. Counting them early prevents a nasty surprise later.
    $privateChannels = 0
    $channelErrors   = 0
    foreach ($t in $teams) {
        try {
            $ch = Get-GraphAll "https://graph.microsoft.com/v1.0/teams/$($t.id)/channels"
            $privateChannels += @($ch | Where-Object { $_.membershipType -eq 'private' }).Count
        } catch { $channelErrors++ }
    }
    $script:Findings.Groups.PrivateChannels = $privateChannels
    if ($channelErrors) { Add-Issue 'Teams channels' "$channelErrors team(s) could not be enumerated" }

    Write-Ok "$(@($teams).Count) teams, $privateChannels private channels, $(@($groups).Count) groups total"
} catch {
    Add-Issue 'Teams and groups' $_.Exception.Message
}

#endregion

#region 4.9 Identity integrations

<#
    Every application below has to be re-registered and re-consented in the destination
    tenant, and each one needs its vendor contacted. In practice this is the most commonly
    under-estimated part of a tenant move - the mailboxes are predictable, the twenty
    integrations nobody wrote down are not.
#>
Write-Step 'Applications and identity'
try {
    $apps = Get-GraphAll ('https://graph.microsoft.com/v1.0/servicePrincipals?$top=999&$select=' +
            'id,displayName,appId,servicePrincipalType,tags,accountEnabled,preferredSingleSignOnMode')

    $thirdParty = $apps | Where-Object {
        $_.servicePrincipalType -eq 'Application' -and
        $_.tags -contains 'WindowsAzureActiveDirectoryIntegratedApp'
    }
    $ssoApps = $apps | Where-Object { $_.preferredSingleSignOnMode }

    $script:Findings.Applications = [ordered]@{
        Total      = @($apps).Count
        ThirdParty = @($thirdParty).Count
        WithSso    = @($ssoApps).Count
    }
    Save-Csv 'applications' ($thirdParty | Select-Object displayName, appId, accountEnabled, preferredSingleSignOnMode)
    Write-Ok "$(@($thirdParty).Count) integrated applications, $(@($ssoApps).Count) with SSO configured"

    try {
        $ca = Get-GraphAll 'https://graph.microsoft.com/v1.0/identity/conditionalAccess/policies'
        $script:Findings.Applications.ConditionalAccessPolicies = @($ca).Count
        $script:Findings.Applications.CaPoliciesEnabled = @($ca | Where-Object { $_.state -eq 'enabled' }).Count
        Save-Csv 'conditional-access' ($ca | Select-Object displayName, state, createdDateTime)
        Write-Ok "$(@($ca).Count) conditional access policies"
    } catch {
        Add-Issue 'Conditional access' 'Not readable - may require Entra ID P1 or additional rights'
    }
} catch {
    Add-Issue 'Applications' $_.Exception.Message
}

#endregion

#region 4.10 Exchange Online - structure

<#
    Graph reports give sizes. These Exchange cmdlets give the structures that make a
    migration complicated: shared and resource mailboxes that need re-permissioning, public
    folders that most tooling cannot move, transport rules and connectors that have to be
    rebuilt by hand, and legal holds that may legally prevent moving a mailbox at all.
#>
if ($hasExo) {
    Write-Step 'Exchange Online structure'
    try {
        Connect-ExchangeOnline -ShowBanner:$false
        Write-Ok 'Connected to Exchange Online'

        $allMbx = Get-EXOMailbox -ResultSize Unlimited -Properties LitigationHoldEnabled, ArchiveStatus

        $script:Findings.Exchange = [ordered]@{
            UserMailboxes     = @($allMbx | Where-Object { $_.RecipientTypeDetails -eq 'UserMailbox' }).Count
            SharedMailboxes   = @($allMbx | Where-Object { $_.RecipientTypeDetails -eq 'SharedMailbox' }).Count
            RoomMailboxes     = @($allMbx | Where-Object { $_.RecipientTypeDetails -eq 'RoomMailbox' }).Count
            EquipmentMailbox  = @($allMbx | Where-Object { $_.RecipientTypeDetails -eq 'EquipmentMailbox' }).Count
            ArchivesEnabled   = @($allMbx | Where-Object { $_.ArchiveStatus -eq 'Active' }).Count
            LitigationHold    = @($allMbx | Where-Object { $_.LitigationHoldEnabled }).Count
        }

        Save-Csv 'mailbox-types' ($allMbx | Select-Object DisplayName, PrimarySmtpAddress,
                                  RecipientTypeDetails, ArchiveStatus, LitigationHoldEnabled)

        <#
            The usage report carries sizes but not recipient type; Exchange carries type but
            not size. Joining them is the only way to answer the question that actually
            matters here - how much of the data belongs to real people, and how much sits in
            shared mailboxes and system accounts that nobody thinks about until cutover.
        #>
        if ($script:MailboxRaw) {
            $sizeByUpn = @{}
            foreach ($row in $script:MailboxRaw) {
                $upn = $row.'User Principal Name'
                if ($upn) { $sizeByUpn[$upn.ToLower()] = [double]$row.($script:MailboxSizeCol) }
            }

            $byType = @{}
            $sharedRows = New-Object System.Collections.Generic.List[object]

            foreach ($m in $allMbx) {
                $key   = ($m.PrimarySmtpAddress).ToString().ToLower()
                $bytes = if ($sizeByUpn.ContainsKey($key)) { $sizeByUpn[$key] } else { 0 }
                $type  = $m.RecipientTypeDetails

                if (-not $byType.ContainsKey($type)) {
                    $byType[$type] = [pscustomobject]@{ Type = $type; Count = 0; Bytes = 0 }
                }
                $byType[$type].Count++
                $byType[$type].Bytes += $bytes

                if ($type -ne 'UserMailbox' -and $bytes -gt 0) {
                    $sharedRows.Add([pscustomobject]@{
                        Mailbox = $m.PrimarySmtpAddress
                        Type    = $type
                        SizeGB  = ConvertTo-GB $bytes
                    })
                }
            }

            $script:Findings.MailboxSplit = $byType.Values |
                ForEach-Object {
                    [pscustomobject]@{
                        Type    = $_.Type
                        Count   = $_.Count
                        TotalGB = ConvertTo-GB $_.Bytes
                    }
                } | Sort-Object TotalGB -Descending

            $script:Findings.NonUserMailboxes = $sharedRows | Sort-Object SizeGB -Descending |
                                                Select-Object -First 20
            Save-Csv 'mailbox-size-by-type' $script:Findings.MailboxSplit
            Save-Csv 'non-user-mailboxes'   ($sharedRows | Sort-Object SizeGB -Descending)

            $nonUserGb = ($sharedRows | Measure-Object SizeGB -Sum).Sum
            Write-Ok "$($sharedRows.Count) non-user mailboxes holding $nonUserGb GB"
        }

        try {
            $rules = Get-TransportRule
            $script:Findings.Exchange.TransportRules = @($rules).Count
            Save-Csv 'transport-rules' ($rules | Select-Object Name, State, Priority, Description)
        } catch { Add-Issue 'Transport rules' $_.Exception.Message }

        try {
            $conn = Get-InboundConnector
            $conn2 = Get-OutboundConnector
            $script:Findings.Exchange.Connectors = @($conn).Count + @($conn2).Count
            Save-Csv 'mail-connectors' (@($conn) + @($conn2) | Select-Object Name, Enabled, ConnectorType)
        } catch { Add-Issue 'Mail connectors' $_.Exception.Message }

        try {
            # Public folders are worth singling out: if they exist, they are almost always
            # the longest pole in the migration.
            $pf = Get-OrganizationConfig | Select-Object -ExpandProperty PublicFoldersEnabled
            $script:Findings.Exchange.PublicFolders = $pf
            if ($pf -eq 'Remote' -or $pf -eq 'Local') {
                Write-Note "Public folders are enabled ($pf) - these need separate planning."
            }
        } catch { Add-Issue 'Public folders' $_.Exception.Message }

        Write-Ok ("$($script:Findings.Exchange.UserMailboxes) user, " +
                  "$($script:Findings.Exchange.SharedMailboxes) shared, " +
                  "$($script:Findings.Exchange.RoomMailboxes) room mailboxes")

        Disconnect-ExchangeOnline -Confirm:$false | Out-Null
    } catch {
        Add-Issue 'Exchange Online' $_.Exception.Message
    }
} else {
    Write-Note 'Exchange section skipped.'
}

#endregion

}   # end of the collection phase skipped by -SelfTest

#-------------------------------------------------------------------------------------------
# REGION 5 - Report
#-------------------------------------------------------------------------------------------
#region Report

Write-Step 'Building the report'

function HtmlEncode { param($s) if ($null -eq $s) { return '' } [System.Net.WebUtility]::HtmlEncode([string]$s) }

function New-StatCard {
    param($Label, $Value, $Sub)
    @"
<div class="card"><span class="k">$(HtmlEncode $Label)</span>
<span class="v">$(HtmlEncode $Value)</span>
<span class="s">$(HtmlEncode $Sub)</span></div>
"@
}

function New-Table {
    param($Rows, [string] $Caption)
    if (-not $Rows -or @($Rows).Count -eq 0) { return '' }
    $cols = $Rows[0].PSObject.Properties.Name
    $head = ($cols | ForEach-Object { "<th>$(HtmlEncode $_)</th>" }) -join ''
    $body = foreach ($r in $Rows) {
        $cells = ($cols | ForEach-Object { "<td>$(HtmlEncode $r.$_)</td>" }) -join ''
        "<tr>$cells</tr>"
    }
    "<h3>$(HtmlEncode $Caption)</h3><div class='tw'><table><thead><tr>$head</tr></thead><tbody>$($body -join '')</tbody></table></div>"
}

$f = $script:Findings
$generated = Get-Date -Format 'd MMMM yyyy, HH:mm'

# --- headline numbers -------------------------------------------------------------------
$totalDataGb = [math]::Round(
    ([double]($f.Mailboxes.TotalGB))  +
    ([double]($f.OneDrive.TotalGB))   +
    ([double]($f.SharePoint.TotalGB)), 2)

$cards  = New-StatCard 'Licensed users'  $f.Accounts.Licensed      "$($f.Accounts.Total) accounts total"
$cards += New-StatCard 'Mailboxes'       $f.Mailboxes.Count        "$($f.Mailboxes.TotalGB) GB"
$cards += New-StatCard 'OneDrive'        $f.OneDrive.Sites         "$($f.OneDrive.TotalGB) GB"
$cards += New-StatCard 'SharePoint sites' $f.SharePoint.Sites      "$($f.SharePoint.TotalGB) GB"
$cards += New-StatCard 'Teams'           $f.Groups.Teams           "$($f.Groups.PrivateChannels) private channels"
$cards += New-StatCard 'Total data'      "$totalDataGb GB"         'mail + files'

# --- complexity flags -------------------------------------------------------------------
# These are the specific conditions that move an estimate materially. Each one is stated
# with the reason, so the reader can see why it matters rather than taking it on trust.
$flags = New-Object System.Collections.Generic.List[string]

if ($f.Tenant.DirSyncEnabled) {
    $flags.Add('Directory synchronisation is enabled, so identities originate from an on-premises directory. That directory has to be repointed or retired as part of the move, and it is a project in its own right.')
}
if ([int]$f.Tenant.DomainsCustom -gt 0) {
    $flags.Add("There are $($f.Tenant.DomainsCustom) custom domains. A domain can only exist in one tenant at a time, so each must be removed from the source and verified in the destination inside the cutover window - this is what makes the cutover time-boxed rather than gradual.")
}
if ([int]$f.Exchange.LitigationHold -gt 0) {
    $flags.Add("$($f.Exchange.LitigationHold) mailbox(es) are on litigation hold. Held content may not be movable without legal sign-off, and in some cases the source tenant must be retained until the hold lifts.")
}
if ([int]$f.Groups.PrivateChannels -gt 0) {
    $flags.Add("$($f.Groups.PrivateChannels) private Teams channel(s). Each has its own hidden SharePoint site, and most migration tooling handles these poorly or not at all - they often need manual recreation.")
}
if ([int]$f.Exchange.SharedMailboxes -gt 0) {
    $flags.Add("$($f.Exchange.SharedMailboxes) shared mailbox(es). Delegate and Send-As permissions do not migrate automatically and have to be re-applied and tested per mailbox.")
}
if ([int]$f.Mailboxes.Over50GB -gt 0) {
    $flags.Add("$($f.Mailboxes.Over50GB) mailbox(es) exceed 50 GB. Large mailboxes dominate the migration schedule and usually need to be seeded well ahead of the cutover.")
}
if ([int]$f.Applications.WithSso -gt 0) {
    $flags.Add("$($f.Applications.WithSso) application(s) use single sign-on against this tenant. Each has to be re-registered in the destination and re-tested, usually with the vendor involved.")
}
if ($f.Exchange.PublicFolders -and $f.Exchange.PublicFolders -ne 'None') {
    $flags.Add('Public folders are in use. These are consistently the slowest workload to move and warrant their own plan.')
}
if ($f.ReportsAnonymised -eq $true) {
    $flags.Add('Usage reports are anonymised in this tenant, so per-user and per-site names appear as hashes. Totals are accurate; the detail tables are not attributable until the setting is turned off and the script is re-run.')
}

$flagHtml = if ($flags.Count) {
    "<ul>" + (($flags | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join '') + "</ul>"
} else {
    "<p>No unusual complexity factors were detected.</p>"
}

$issueHtml = if ($script:Issues.Count) {
    "<div class='note warn'><b>Sections that could not be collected</b><ul>" +
    (($script:Issues | ForEach-Object { "<li>$(HtmlEncode $_)</li>" }) -join '') +
    "</ul><p>These gaps are listed so the estimate is not built on data that silently went missing.</p></div>"
} else { '' }

# --- logo (embedded at the bottom of this file) ------------------------------------------
$logoTag = if ($script:LogoBase64) {
    "<img src='data:image/png;base64,$script:LogoBase64' alt='EnTech Engineering' class='logo'>"
} else {
    "<span class='wordmark'>EnTech</span>"
}

$html = @"
<!doctype html><html lang="en"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Microsoft 365 Discovery - $(HtmlEncode $f.Tenant.DisplayName)</title>
<style>
 *{box-sizing:border-box}
 body{margin:0;background:#fff;color:#15191b;font:14px/1.6 "Segoe UI",system-ui,sans-serif}
 .sheet{max-width:950px;margin:0 auto;padding:36px 30px 60px}
 .hd{display:flex;align-items:flex-end;justify-content:space-between;gap:20px;
     border-bottom:2.5px solid $script:Brand;padding-bottom:12px;margin-bottom:22px}
 .logo{height:34px} .wordmark{font-size:23px;font-weight:700;color:$script:Brand}
 .hd .rt{text-align:right;font-size:10px;letter-spacing:.14em;text-transform:uppercase;color:#798286;font-weight:600;line-height:1.6}
 h1{font-size:25px;margin:0 0 4px;letter-spacing:-.02em}
 h2{font-size:16px;margin:30px 0 4px;padding-bottom:6px;border-bottom:1px solid #dde3e4}
 h3{font-size:13px;margin:18px 0 4px;color:#4a5255}
 p{margin:9px 0;max-width:76ch}
 .sub{color:#4a5255;font-size:15px;margin:0 0 4px}
 .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(135px,1fr));gap:1px;
        background:#dde3e4;border:1px solid #dde3e4;margin:20px 0}
 .card{background:#fff;padding:11px 13px}
 .card .k{display:block;font-size:9px;letter-spacing:.13em;text-transform:uppercase;color:#798286;margin-bottom:3px}
 .card .v{display:block;font-size:21px;font-weight:700;line-height:1.15;font-variant-numeric:tabular-nums}
 .card .s{display:block;font-size:11px;color:#798286;margin-top:1px}
 .tw{overflow-x:auto;border:1px solid #dde3e4;margin:8px 0 14px}
 table{width:100%;border-collapse:collapse;font-size:12.5px}
 th,td{text-align:left;padding:6px 10px;border-bottom:1px solid #eef1f2}
 th{background:#f3f6f6;font-size:9.5px;letter-spacing:.1em;text-transform:uppercase;color:#798286;font-weight:600}
 tr:last-child td{border-bottom:0}
 .note{border-left:3px solid $script:Brand;background:#f3f6f6;padding:12px 16px;margin:16px 0}
 .note.warn{border-left-color:#8a5a00}
 .note b{display:block;margin-bottom:5px}
 ul{margin:8px 0;padding-left:20px;max-width:76ch} li{margin:5px 0}
 .foot{margin-top:34px;padding-top:12px;border-top:1px solid #dde3e4;font-size:11px;color:#798286}
 @media print{ .sheet{max-width:none;padding:14mm} h2{page-break-after:avoid}
   table,.note{page-break-inside:avoid} *{-webkit-print-color-adjust:exact;print-color-adjust:exact} }
 /* margin:0 leaves the browser nowhere to draw its own header and footer, which would
    otherwise stamp the local file path across every page of the client's report. */
 @page{size:Letter;margin:0}
</style></head><body><div class="sheet">

<div class="hd">
  $logoTag
  <div class="rt">Microsoft 365 Discovery<br>Prepared for migration sizing</div>
</div>

<h1>Microsoft 365 tenant discovery</h1>
<p class="sub">$(HtmlEncode $f.Tenant.DisplayName) &middot; $(HtmlEncode $f.Tenant.DefaultDomain)</p>

<div class="note">
  <b>How this was produced</b>
  <p style="margin:0">$(if ($SelfTest) { '<b>SAMPLE DATA - self-test run. No tenant was contacted and no figure below describes a real environment.</b> ' })Collected by a read-only script run by your own administrator inside your
  own tenant on $generated, over a $ReportPeriod reporting window. No content of any message or
  document was read, nothing was changed, and nothing left the machine it ran on.</p>
</div>

<div class="cards">$cards</div>

<h2>Tenant</h2>
<div class="tw"><table><tbody>
<tr><th style="width:34%">Tenant name</th><td>$(HtmlEncode $f.Tenant.DisplayName)</td></tr>
<tr><th>Tenant ID</th><td>$(HtmlEncode $f.Tenant.TenantId)</td></tr>
<tr><th>Default domain</th><td>$(HtmlEncode $f.Tenant.DefaultDomain)</td></tr>
<tr><th>Custom domains</th><td>$(HtmlEncode $f.Tenant.DomainsCustom) of $(HtmlEncode $f.Tenant.DomainsTotal) total</td></tr>
<tr><th>Directory sync from on-premises</th><td>$(if($f.Tenant.DirSyncEnabled){'Yes'}else{'No'})</td></tr>
</tbody></table></div>

<h2>What drives the estimate</h2>
<p>Volume sets the duration of a migration; the items below set its difficulty. Each is
something that adds work beyond simply copying data.</p>
$flagHtml

<h2>Detail</h2>
$(New-Table $f.Licensing.Rows      'Licences')
$(New-Table $f.Mailboxes.Largest   'Largest mailboxes')
$(New-Table $f.OneDrive.Largest    'Largest OneDrive accounts')
$(New-Table $f.SharePoint.Largest  'Largest SharePoint sites')

$(if ($f.MailboxSplit) { "<h2>Where the mail data sits</h2>
<p>Split by mailbox type. The non-user rows - shared mailboxes, rooms and system accounts -
are the ones most often missed when a migration is scoped from a headcount.</p>" +
(New-Table $f.MailboxSplit 'Data by mailbox type') +
(New-Table $f.NonUserMailboxes 'Largest non-user mailboxes') })

$(if ($f.SharePointShape) { "<h2>What SharePoint is being used for</h2>" + $(
  if ($f.SharePointShape.FileStorageOnly) {
    "<p><b>File storage only.</b> Across $($f.SharePointShape.SitesInspected) site(s) we found
    $($f.SharePointShape.DocumentLibs) document libraries and no custom lists. Content of this
    shape is copied rather than rebuilt, which is the straightforward case.</p>"
  } else {
    "<p><b>More than file storage.</b> Alongside $($f.SharePointShape.DocumentLibs) document
    libraries there are <b>$($f.SharePointShape.CustomLists) custom list(s)</b> across
    $($f.SharePointShape.SitesWithLists) site(s). Lists, pages and forms do not migrate between
    tenants - they are rebuilt by hand, and whoever built them originally is the person who
    knows what they do. This is usually the largest single unknown in a SharePoint move.</p>" +
    (New-Table $f.SharePointShape.Detail 'Sites carrying custom lists')
  }) })

<h2>Workload summary</h2>
<div class="tw"><table><tbody>
<tr><th style="width:34%">User mailboxes</th><td>$(HtmlEncode $f.Exchange.UserMailboxes)</td></tr>
<tr><th>Shared mailboxes</th><td>$(HtmlEncode $f.Exchange.SharedMailboxes)</td></tr>
<tr><th>Room / equipment mailboxes</th><td>$(HtmlEncode $f.Exchange.RoomMailboxes) / $(HtmlEncode $f.Exchange.EquipmentMailbox)</td></tr>
<tr><th>Archive mailboxes enabled</th><td>$(HtmlEncode $f.Exchange.ArchivesEnabled)</td></tr>
<tr><th>Mailboxes on litigation hold</th><td>$(HtmlEncode $f.Exchange.LitigationHold)</td></tr>
<tr><th>Mail transport rules</th><td>$(HtmlEncode $f.Exchange.TransportRules)</td></tr>
<tr><th>Mail connectors</th><td>$(HtmlEncode $f.Exchange.Connectors)</td></tr>
<tr><th>Microsoft 365 groups</th><td>$(HtmlEncode $f.Groups.M365Groups)</td></tr>
<tr><th>Distribution lists</th><td>$(HtmlEncode $f.Groups.Distribution)</td></tr>
<tr><th>Teams</th><td>$(HtmlEncode $f.Groups.Teams) ($(HtmlEncode $f.Groups.PrivateChannels) private channels)</td></tr>
<tr><th>Integrated applications</th><td>$(HtmlEncode $f.Applications.ThirdParty) ($(HtmlEncode $f.Applications.WithSso) with SSO)</td></tr>
<tr><th>Conditional access policies</th><td>$(HtmlEncode $f.Applications.ConditionalAccessPolicies)</td></tr>
</tbody></table></div>

$issueHtml

<div class="foot">
  Generated $generated by the EnTech Microsoft 365 discovery script, version 1.0, over a
  $ReportPeriod window. Supporting CSV extracts accompany this report in the <code>data</code>
  folder. Figures reflect the tenant as configured on the date of collection.<br>
  EnTech Engineering, P.C. &middot; 17 State St, 36th Fl, New York, NY 10004
</div>

</div></body></html>
"@

$htmlFile = Join-Path $OutputPath 'M365-Discovery-Report.html'
$html | Out-File -FilePath $htmlFile -Encoding UTF8
Write-Ok "report written: $(Split-Path $htmlFile -Leaf)"

#endregion

#-------------------------------------------------------------------------------------------
# REGION 6 - PDF
#-------------------------------------------------------------------------------------------
#region Pdf

<#
    PowerShell has no native PDF writer. Rather than take a dependency on a third-party
    module - which a government tenant may not permit installing - we use the headless mode
    of a browser that is already present on every supported Windows build. If neither is
    found we leave the HTML, which opens and prints perfectly well from any browser.
#>
if (-not $SkipPdf) {
    Write-Step 'Creating PDF'
    $browsers = @(
        "$env:ProgramFiles\Microsoft\Edge\Application\msedge.exe"
        "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe"
        "$env:ProgramFiles\Google\Chrome\Application\chrome.exe"
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe"
    ) | Where-Object { Test-Path $_ }

    if ($browsers) {
        $pdfFile = Join-Path $OutputPath 'M365-Discovery-Report.pdf'
        try {
            & $browsers[0] --headless --disable-gpu --no-sandbox `
                           --print-to-pdf-no-header "--print-to-pdf=$pdfFile" `
                           "file:///$($htmlFile -replace '\\','/')" 2>$null
            Start-Sleep -Seconds 3
            if (Test-Path $pdfFile) { Write-Ok 'PDF created' }
            else { Add-Issue 'PDF' 'The browser ran but produced no file - use the HTML instead' }
        } catch {
            Add-Issue 'PDF' $_.Exception.Message
        }
    } else {
        Write-Note 'No Edge or Chrome found - HTML report only. Open it and print to PDF.'
    }
}

#endregion

#-------------------------------------------------------------------------------------------
# REGION 7 - Package and finish
#-------------------------------------------------------------------------------------------
#region Finish

Write-Step 'Packaging'
try {
    $zip = Join-Path ([IO.Path]::GetDirectoryName($OutputPath)) `
                     ("$([IO.Path]::GetFileName($OutputPath)).zip")
    if (Test-Path $zip) { Remove-Item $zip -Force }
    Compress-Archive -Path "$OutputPath\*" -DestinationPath $zip -ErrorAction Stop
    if (Test-Path $zip) { Write-Ok "zipped: $zip" }
    else { Add-Issue 'Packaging' 'Archive was not created - send the folder contents instead' }
} catch {
    Add-Issue 'Packaging' $_.Exception.Message
}

try { Disconnect-MgGraph | Out-Null } catch { }

$elapsed = [math]::Round(((Get-Date) - $script:StartedAt).TotalMinutes, 1)

Write-Host ""
Write-Host "  Done in $elapsed minute(s)." -ForegroundColor Green
Write-Host "  Folder: $OutputPath" -ForegroundColor White
Write-Host ""
Write-Host "  Please open the report and review it before sending it to anyone." -ForegroundColor Yellow
Write-Host "  It contains account names, site names and group names from your tenant." -ForegroundColor Gray
Write-Host ""

if (Test-Path $htmlFile) { Invoke-Item $htmlFile }

#endregion
