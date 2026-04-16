#Requires -Modules ExchangeOnlineManagement
<#!
.SYNOPSIS
  Exports Microsoft Defender for Office 365 policy and rule configuration to JSON and HTML.

.DESCRIPTION
  This script is intentionally separate from the deployment script. It reads the current
  Defender for Office 365 / Exchange Online protection configuration and produces:
    1. A JSON export for structured comparison and archival
    2. An HTML report for readable validation and customer-friendly output

.PARAMETER OutputFolder
  Folder where the JSON and HTML files will be written.

.PARAMETER Prefix
  Prefix used for the output file names.

.PARAMETER ConnectIfNeeded
  If set, the script will try to connect to Exchange Online if it does not detect an
  existing session.

.EXAMPLE
  .\Export-DefenderForOffice365Report_Styled.ps1

.EXAMPLE
  .\Export-DefenderForOffice365Report_Styled.ps1 -OutputFolder .\output -Prefix Contoso-MDO -ConnectIfNeeded
#>

[CmdletBinding()]
param(
    [string]$OutputFolder = (Join-Path -Path (Get-Location) -ChildPath 'output'),
    [string]$Prefix = 'DefenderForOffice365',
    [switch]$ConnectIfNeeded
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Write-Info {
    param([string]$Message)
    Write-Host "[INFO] $Message" -ForegroundColor Cyan
}

function Write-WarnMsg {
    param([string]$Message)
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Test-ExchangeOnlineConnection {
    try {
        $null = Get-OrganizationConfig -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

function Ensure-ExchangeOnlineConnection {
    if (Test-ExchangeOnlineConnection) {
        Write-Info 'Exchange Online session detected.'
        return
    }

    if (-not $ConnectIfNeeded) {
        throw 'No active Exchange Online session detected. Connect first with Connect-ExchangeOnline, or rerun with -ConnectIfNeeded.'
    }

    Write-Info 'No active Exchange Online session detected. Connecting to Exchange Online...'
    Connect-ExchangeOnline -ShowBanner:$false

    if (-not (Test-ExchangeOnlineConnection)) {
        throw 'Unable to verify Exchange Online connection after Connect-ExchangeOnline.'
    }
}

function Convert-ObjectForJson {
    param(
        [Parameter(Mandatory)]$InputObject,
        [int]$Depth = 8
    )

    if ($null -eq $InputObject) { return $null }
    return ($InputObject | ConvertTo-Json -Depth $Depth -Compress | ConvertFrom-Json)
}

function Get-SafeData {
    param(
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)][scriptblock]$ScriptBlock
    )

    try {
        Write-Info "Collecting $Name..."
        $result = & $ScriptBlock
        if ($null -eq $result) { return @() }
        return @($result)
    }
    catch {
        Write-WarnMsg "Failed to collect $Name. $($_.Exception.Message)"
        return @(
            [pscustomobject]@{
                Collection = $Name
                Error      = $_.Exception.Message
            }
        )
    }
}

function Get-DisplayValue {
    param($Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [System.Array]) {
        if ($Value.Count -eq 0) { return '' }
        return (($Value | ForEach-Object { [string]$_ }) -join ', ')
    }
    if ($Value -is [datetime]) {
        return $Value.ToString('yyyy-MM-dd HH:mm:ss')
    }
    return [string]$Value
}

function Select-UsefulProperties {
    param(
        [Parameter(Mandatory)]$Objects,
        [string[]]$PreferredProperties
    )

    $list = @($Objects)
    if (-not $list -or $list.Count -eq 0) { return @() }

    $available = @($list[0].PSObject.Properties.Name)
    $props = foreach ($p in $PreferredProperties) {
        if ($available -contains $p) { $p }
    }

    if (-not $props -or $props.Count -eq 0) {
        return $list
    }

    return $list | Select-Object -Property $props
}

function Get-StatusClass {
    param([string]$Value)

    switch -Regex ($Value) {
        '^(True|Enabled|On)$' { return 'status-good' }
        '^(False|Disabled|Off)$' { return 'status-muted' }
        '^(Quarantine|Block|Reject)$' { return 'status-warn' }
        default { return '' }
    }
}

function Convert-DataToHtmlTable {
    param(
        [Parameter(Mandatory)]$Data,
        [string[]]$PreferredProperties
    )

    $rows = @(Select-UsefulProperties -Objects $Data -PreferredProperties $PreferredProperties)
    if (-not $rows -or $rows.Count -eq 0) {
        return "<p class='muted'>No data returned.</p>"
    }

    $columns = @($rows[0].PSObject.Properties.Name)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("<div class='table-wrap'><table>")
    [void]$sb.AppendLine('<thead><tr>')
    foreach ($col in $columns) {
        [void]$sb.AppendLine("<th>$([System.Net.WebUtility]::HtmlEncode($col))</th>")
    }
    [void]$sb.AppendLine('</tr></thead>')
    [void]$sb.AppendLine('<tbody>')

    foreach ($row in $rows) {
        [void]$sb.AppendLine('<tr>')
        foreach ($col in $columns) {
            $value = Get-DisplayValue -Value $row.$col
            $encoded = [System.Net.WebUtility]::HtmlEncode($value)
            $class = Get-StatusClass -Value $value
            if ([string]::IsNullOrWhiteSpace($class)) {
                [void]$sb.AppendLine("<td>$encoded</td>")
            }
            else {
                [void]$sb.AppendLine("<td><span class='badge $class'>$encoded</span></td>")
            }
        }
        [void]$sb.AppendLine('</tr>')
    }

    [void]$sb.AppendLine('</tbody></table></div>')
    return $sb.ToString()
}

function New-ReportSection {
    param(
        [Parameter(Mandatory)][string]$Title,
        [string]$Description,
        [Parameter(Mandatory)]$Data,
        [string[]]$PreferredProperties
    )

    $table = Convert-DataToHtmlTable -Data $Data -PreferredProperties $PreferredProperties
    $descHtml = if ($Description) { "<p class='section-desc'>$([System.Net.WebUtility]::HtmlEncode($Description))</p>" } else { '' }
    return @"
<section class='section-card'>
  <div class='section-header'>
    <h2>$([System.Net.WebUtility]::HtmlEncode($Title))</h2>
    $descHtml
  </div>
  $table
</section>
"@
}

function New-SummaryTile {
    param(
        [Parameter(Mandatory)][string]$Label,
        [Parameter(Mandatory)][int]$Count
    )

    return @"
<div class='summary-tile'>
  <div class='summary-value'>$Count</div>
  <div class='summary-label'>$([System.Net.WebUtility]::HtmlEncode($Label))</div>
</div>
"@
}

Ensure-ExchangeOnlineConnection

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
}

$timestamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$jsonPath  = Join-Path -Path $OutputFolder -ChildPath "$Prefix-$timestamp.json"
$htmlPath  = Join-Path -Path $OutputFolder -ChildPath "$Prefix-$timestamp.html"

$orgConfig = $null
try {
    $orgConfig = Get-OrganizationConfig | Select-Object Name, DisplayName
}
catch {
    $orgConfig = [pscustomobject]@{
        Name        = 'Unknown'
        DisplayName = 'Unknown'
    }
}

$collections = [ordered]@{
    TenantInfo = [pscustomobject]@{
        ExportedAt      = (Get-Date).ToString('s')
        TenantName      = $orgConfig.Name
        TenantDisplay   = $orgConfig.DisplayName
        ReportGenerator = $env:USERNAME
        Hostname        = $env:COMPUTERNAME
    }

    AntiPhishPolicies           = Get-SafeData -Name 'Anti-Phish Policies' -ScriptBlock { Get-AntiPhishPolicy }
    AntiPhishRules              = Get-SafeData -Name 'Anti-Phish Rules' -ScriptBlock { Get-AntiPhishRule }

    AntiSpamInboundPolicies     = Get-SafeData -Name 'Inbound Anti-Spam Policies' -ScriptBlock { Get-HostedContentFilterPolicy }
    AntiSpamInboundRules        = Get-SafeData -Name 'Inbound Anti-Spam Rules' -ScriptBlock { Get-HostedContentFilterRule }

    AntiSpamOutboundPolicies    = Get-SafeData -Name 'Outbound Anti-Spam Policies' -ScriptBlock { Get-HostedOutboundSpamFilterPolicy }
    AntiSpamOutboundRules       = Get-SafeData -Name 'Outbound Anti-Spam Rules' -ScriptBlock { Get-HostedOutboundSpamFilterRule }

    SafeLinksPolicies           = Get-SafeData -Name 'Safe Links Policies' -ScriptBlock { Get-SafeLinksPolicy }
    SafeLinksRules              = Get-SafeData -Name 'Safe Links Rules' -ScriptBlock { Get-SafeLinksRule }

    SafeAttachmentPolicies      = Get-SafeData -Name 'Safe Attachment Policies' -ScriptBlock { Get-SafeAttachmentPolicy }
    SafeAttachmentRules         = Get-SafeData -Name 'Safe Attachment Rules' -ScriptBlock { Get-SafeAttachmentRule }

    MalwarePolicies             = Get-SafeData -Name 'Malware Policies' -ScriptBlock { Get-MalwareFilterPolicy }
    MalwareRules                = Get-SafeData -Name 'Malware Rules' -ScriptBlock { Get-MalwareFilterRule }

    AcceptedDomains             = Get-SafeData -Name 'Accepted Domains' -ScriptBlock {
        Get-AcceptedDomain | Select-Object DomainName, Default
    }
}

$jsonExport = [ordered]@{}
foreach ($key in $collections.Keys) {
    $jsonExport[$key] = Convert-ObjectForJson -InputObject $collections[$key] -Depth 8
}

$jsonExport | ConvertTo-Json -Depth 8 | Out-File -LiteralPath $jsonPath -Encoding utf8
Write-Info "JSON export written to: $jsonPath"

$summaryTiles = @()
$summaryTiles += New-SummaryTile -Label 'Anti-Phish Policies' -Count @($collections.AntiPhishPolicies).Count
$summaryTiles += New-SummaryTile -Label 'Anti-Phish Rules' -Count @($collections.AntiPhishRules).Count
$summaryTiles += New-SummaryTile -Label 'Inbound Spam Policies' -Count @($collections.AntiSpamInboundPolicies).Count
$summaryTiles += New-SummaryTile -Label 'Inbound Spam Rules' -Count @($collections.AntiSpamInboundRules).Count
$summaryTiles += New-SummaryTile -Label 'Outbound Spam Policies' -Count @($collections.AntiSpamOutboundPolicies).Count
$summaryTiles += New-SummaryTile -Label 'Outbound Spam Rules' -Count @($collections.AntiSpamOutboundRules).Count
$summaryTiles += New-SummaryTile -Label 'Safe Links Policies' -Count @($collections.SafeLinksPolicies).Count
$summaryTiles += New-SummaryTile -Label 'Safe Links Rules' -Count @($collections.SafeLinksRules).Count
$summaryTiles += New-SummaryTile -Label 'Safe Attachments Policies' -Count @($collections.SafeAttachmentPolicies).Count
$summaryTiles += New-SummaryTile -Label 'Safe Attachments Rules' -Count @($collections.SafeAttachmentRules).Count
$summaryTiles += New-SummaryTile -Label 'Malware Policies' -Count @($collections.MalwarePolicies).Count
$summaryTiles += New-SummaryTile -Label 'Malware Rules' -Count @($collections.MalwareRules).Count

$style = @"
<style>
body {
    margin: 0;
    background: #f2f2f2;
    color: #1f1f1f;
    font-family: Segoe UI, Arial, sans-serif;
}
.container {
    max-width: 1500px;
    margin: 0 auto;
    padding: 24px;
}
.hero {
    background: #ffffff;
    border: 1px solid #dddddd;
    border-left: 8px solid #311640;
    border-radius: 12px;
    padding: 24px;
    margin-bottom: 24px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.hero h1 {
    margin: 0 0 8px 0;
    color: #311640;
    font-size: 28px;
}
.hero p {
    margin: 6px 0;
}
.meta-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
    gap: 12px;
    margin-top: 16px;
}
.meta-card,
.summary-tile,
.section-card {
    background: #ffffff;
    border: 1px solid #dddddd;
    border-left: 6px solid #311640;
    border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05);
}
.meta-card {
    padding: 14px 16px;
}
.meta-label {
    font-size: 12px;
    text-transform: uppercase;
    letter-spacing: .04em;
    color: #666666;
    margin-bottom: 4px;
}
.meta-value {
    font-size: 15px;
    font-weight: 600;
    word-break: break-word;
}
.summary-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 14px;
    margin-bottom: 24px;
}
.summary-tile {
    padding: 16px;
}
.summary-value {
    font-size: 30px;
    font-weight: 700;
    color: #311640;
    line-height: 1;
    margin-bottom: 8px;
}
.summary-label {
    font-size: 13px;
    color: #444444;
}
.section-card {
    padding: 20px;
    margin-bottom: 22px;
}
.section-header {
    margin-bottom: 14px;
}
.section-header h2 {
    margin: 0 0 6px 0;
    color: #311640;
    font-size: 20px;
}
.section-desc {
    margin: 0;
    color: #555555;
    font-size: 13px;
}
.table-wrap {
    overflow-x: auto;
}
table {
    width: 100%;
    border-collapse: collapse;
    background: #ffffff;
}
th, td {
    border: 1px solid #d9d9d9;
    padding: 9px 10px;
    text-align: left;
    vertical-align: top;
    font-size: 13px;
}
th {
    background: #f2f2f2;
    color: #311640;
    position: sticky;
    top: 0;
}
tr:nth-child(even) td {
    background: #fafafa;
}
.badge {
    display: inline-block;
    padding: 2px 8px;
    border-radius: 999px;
    font-size: 12px;
    font-weight: 600;
    border: 1px solid transparent;
    white-space: nowrap;
}
.status-good {
    background: #eef7ee;
    color: #256029;
    border-color: #cde7cf;
}
.status-muted {
    background: #f0f0f0;
    color: #5a5a5a;
    border-color: #d9d9d9;
}
.status-warn {
    background: #fff4e5;
    color: #8a4b08;
    border-color: #f3d19c;
}
.muted {
    color: #666666;
}
.footer {
    margin-top: 28px;
    color: #666666;
    font-size: 12px;
    text-align: center;
}
</style>
"@

$htmlSections = @()
$htmlSections += @"
<div class='hero'>
  <h1>Defender for Office 365 Export Report</h1>
  <p>Read-only export of current Defender for Office 365 and Exchange Online protection settings.</p>
  <div class='meta-grid'>
    <div class='meta-card'>
      <div class='meta-label'>Generated</div>
      <div class='meta-value'>$(Get-Date)</div>
    </div>
    <div class='meta-card'>
      <div class='meta-label'>Tenant Display Name</div>
      <div class='meta-value'>$([System.Net.WebUtility]::HtmlEncode([string]$orgConfig.DisplayName))</div>
    </div>
    <div class='meta-card'>
      <div class='meta-label'>Tenant Name</div>
      <div class='meta-value'>$([System.Net.WebUtility]::HtmlEncode([string]$orgConfig.Name))</div>
    </div>
    <div class='meta-card'>
      <div class='meta-label'>JSON Companion File</div>
      <div class='meta-value'>$([System.Net.WebUtility]::HtmlEncode((Split-Path -Leaf $jsonPath)))</div>
    </div>
  </div>
</div>
"@

$htmlSections += "<div class='summary-grid'>$($summaryTiles -join "`n")</div>"
$htmlSections += New-ReportSection -Title 'Tenant Information' -Description 'Basic report metadata and execution context.' -Data @([pscustomobject]$collections.TenantInfo) -PreferredProperties @('ExportedAt','TenantName','TenantDisplay','ReportGenerator','Hostname')
$htmlSections += New-ReportSection -Title 'Accepted Domains' -Description 'Accepted domains currently configured in the tenant.' -Data $collections.AcceptedDomains -PreferredProperties @('DomainName','Default')
$htmlSections += New-ReportSection -Title 'Anti-Phish Policies' -Description 'Policy settings for impersonation, spoof protection, mailbox intelligence, and user safety tips.' -Data $collections.AntiPhishPolicies -PreferredProperties @('Name','Enabled','PhishThresholdLevel','EnableMailboxIntelligence','EnableMailboxIntelligenceProtection','EnableTargetedUserProtection','EnableTargetedDomainsProtection','EnableOrganizationDomainsProtection','EnableSpoofIntelligence','TargetedUserProtectionAction','TargetedUserQuarantineTag','TargetedDomainProtectionAction','TargetedDomainQuarantineTag','MailboxIntelligenceProtectionAction','MailboxIntelligenceQuarantineTag','AuthenticationFailAction','SpoofQuarantineTag','HonorDmarcPolicy')
$htmlSections += New-ReportSection -Title 'Anti-Phish Rules' -Description 'Rule assignments and enforcement state for anti-phishing policies.' -Data $collections.AntiPhishRules -PreferredProperties @('Name','State','Enabled','Priority','AntiPhishPolicy','RecipientDomainIs','SentToMemberOf')
$htmlSections += New-ReportSection -Title 'Inbound Anti-Spam Policies' -Description 'Inbound spam scoring, ZAP, quarantine actions, and bulk mail settings.' -Data $collections.AntiSpamInboundPolicies -PreferredProperties @('Name','BulkThreshold','SpamAction','SpamQuarantineTag','HighConfidenceSpamAction','HighConfidenceSpamQuarantineTag','PhishSpamAction','PhishQuarantineTag','HighConfidencePhishAction','HighConfidencePhishQuarantineTag','BulkSpamAction','BulkQuarantineTag','InlineSafetyTipsEnabled','EnableEndUserSpamNotifications','ZapEnabled','MarkAsSpamBulkMail','IncreaseScoreWithImageLinks','IncreaseScoreWithNumericIps','IncreaseScoreWithRedirectToOtherPort','IncreaseScoreWithBizOrInfoUrls','MarkAsSpamEmptyMessages','MarkAsSpamEmbedTagsInHtml','MarkAsSpamJavaScriptInHtml','MarkAsSpamFormTagsInHtml','MarkAsSpamFramesInHtml','MarkAsSpamWebBugsInHtml','MarkAsSpamObjectTagsInHtml','MarkAsSpamSensitiveWordList','MarkAsSpamSpfRecordHardFail','MarkAsSpamFromAddressAuthFail','MarkAsSpamNdrBackscatter')
$htmlSections += New-ReportSection -Title 'Inbound Anti-Spam Rules' -Description 'Inbound anti-spam rule assignments and current state.' -Data $collections.AntiSpamInboundRules -PreferredProperties @('Name','State','Enabled','Priority','HostedContentFilterPolicy','RecipientDomainIs','SentToMemberOf')
$htmlSections += New-ReportSection -Title 'Outbound Anti-Spam Policies' -Description 'Outbound spam limits, automatic forwarding, and user restriction behavior.' -Data $collections.AntiSpamOutboundPolicies -PreferredProperties @('Name','RecipientLimitExternalPerHour','RecipientLimitInternalPerHour','RecipientLimitPerDay','ActionWhenThresholdReached','AutoForwardingMode','BccSuspiciousOutboundMail','NotifyOutboundSpam','NotifyOutboundSpamRecipients')
$htmlSections += New-ReportSection -Title 'Outbound Anti-Spam Rules' -Description 'Outbound anti-spam rule assignments and current state.' -Data $collections.AntiSpamOutboundRules -PreferredProperties @('Name','State','Enabled','Priority','HostedOutboundSpamFilterPolicy','SenderDomainIs','FromMemberOf')
$htmlSections += New-ReportSection -Title 'Safe Links Policies' -Description 'Safe Links email, Teams, URL scanning, and click tracking settings.' -Data $collections.SafeLinksPolicies -PreferredProperties @('Name','EnableSafeLinksForEmail','EnableSafeLinksForTeams','EnableForInternalSenders','ScanUrls','DeliverMessageAfterScan','DisableUrlRewrite','TrackClicks','AllowClickThrough','EnableOrganizationBranding')
$htmlSections += New-ReportSection -Title 'Safe Links Rules' -Description 'Rule assignments and enforcement state for Safe Links.' -Data $collections.SafeLinksRules -PreferredProperties @('Name','State','Enabled','Priority','SafeLinksPolicy','RecipientDomainIs','SentToMemberOf')
$htmlSections += New-ReportSection -Title 'Safe Attachments Policies' -Description 'Detonation action, quarantine tagging, and redirect behavior for Safe Attachments.' -Data $collections.SafeAttachmentPolicies -PreferredProperties @('Name','Enable','Enabled','Action','QuarantineTag','Redirect','RedirectAddress')
$htmlSections += New-ReportSection -Title 'Safe Attachments Rules' -Description 'Rule assignments and enforcement state for Safe Attachments.' -Data $collections.SafeAttachmentRules -PreferredProperties @('Name','State','Enabled','Priority','SafeAttachmentPolicy','RecipientDomainIs','SentToMemberOf')
$htmlSections += New-ReportSection -Title 'Malware Policies' -Description 'Malware filtering, file type filtering, quarantine, and admin notification settings.' -Data $collections.MalwarePolicies -PreferredProperties @('Name','EnableFileFilter','FileTypes','Action','ZapEnabled','QuarantineTag','EnableInternalSenderAdminNotifications','InternalSenderAdminAddress','EnableExternalSenderAdminNotifications','ExternalSenderAdminAddress')
$htmlSections += New-ReportSection -Title 'Malware Rules' -Description 'Rule assignments and enforcement state for anti-malware policies.' -Data $collections.MalwareRules -PreferredProperties @('Name','State','Enabled','Priority','MalwareFilterPolicy','RecipientDomainIs','SentToMemberOf')

$htmlBody = @"
<html>
<head>
<meta charset='utf-8' />
<meta name='viewport' content='width=device-width, initial-scale=1' />
<title>Defender for Office 365 Export Report</title>
$style
</head>
<body>
<div class='container'>
$($htmlSections -join "`n")
<div class='footer'>Generated by Export-DefenderForOffice365Report_Styled.ps1</div>
</div>
</body>
</html>
"@

$htmlBody | Out-File -LiteralPath $htmlPath -Encoding utf8
Write-Info "HTML report written to: $htmlPath"
Write-Host ''
Write-Host 'Export complete.' -ForegroundColor Green
Write-Host "JSON : $jsonPath" -ForegroundColor Green
Write-Host "HTML : $htmlPath" -ForegroundColor Green
