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
  .\Export-DefenderForOffice365Report.ps1

.EXAMPLE
  .\Export-DefenderForOffice365Report.ps1 -OutputFolder .\output -Prefix Contoso-MDO -ConnectIfNeeded
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
        [int]$Depth = 6
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

function New-HtmlTable {
    param(
        [Parameter(Mandatory)][string]$Title,
        [Parameter(Mandatory)]$Data,
        [string[]]$PreferredProperties
    )

    $rows = @($Data)
    if (-not $rows -or $rows.Count -eq 0) {
        return "<section><h2>$Title</h2><p class='muted'>No data returned.</p></section>"
    }

    $display = Select-UsefulProperties -Objects $rows -PreferredProperties $PreferredProperties
    $fragment = $display | ConvertTo-Html -Fragment
    return "<section><h2>$Title</h2>$fragment</section>"
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

$summary = @(
    [pscustomobject]@{ Section = 'Anti-Phish Policies'; Count = @($collections.AntiPhishPolicies).Count },
    [pscustomobject]@{ Section = 'Anti-Phish Rules'; Count = @($collections.AntiPhishRules).Count },
    [pscustomobject]@{ Section = 'Inbound Anti-Spam Policies'; Count = @($collections.AntiSpamInboundPolicies).Count },
    [pscustomobject]@{ Section = 'Inbound Anti-Spam Rules'; Count = @($collections.AntiSpamInboundRules).Count },
    [pscustomobject]@{ Section = 'Outbound Anti-Spam Policies'; Count = @($collections.AntiSpamOutboundPolicies).Count },
    [pscustomobject]@{ Section = 'Outbound Anti-Spam Rules'; Count = @($collections.AntiSpamOutboundRules).Count },
    [pscustomobject]@{ Section = 'Safe Links Policies'; Count = @($collections.SafeLinksPolicies).Count },
    [pscustomobject]@{ Section = 'Safe Links Rules'; Count = @($collections.SafeLinksRules).Count },
    [pscustomobject]@{ Section = 'Safe Attachments Policies'; Count = @($collections.SafeAttachmentPolicies).Count },
    [pscustomobject]@{ Section = 'Safe Attachments Rules'; Count = @($collections.SafeAttachmentRules).Count },
    [pscustomobject]@{ Section = 'Malware Policies'; Count = @($collections.MalwarePolicies).Count },
    [pscustomobject]@{ Section = 'Malware Rules'; Count = @($collections.MalwareRules).Count }
)

$style = @"
<style>
body {
    font-family: Segoe UI, Arial, sans-serif;
    margin: 24px;
    background: #f7f7f7;
    color: #1f1f1f;
}
h1, h2 {
    color: #311640;
}
.card {
    background: white;
    border: 1px solid #dddddd;
    border-left: 6px solid #311640;
    border-radius: 8px;
    padding: 16px;
    margin-bottom: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06);
}
table {
    border-collapse: collapse;
    width: 100%;
    margin-top: 10px;
    background: white;
}
th, td {
    border: 1px solid #d9d9d9;
    padding: 8px 10px;
    text-align: left;
    vertical-align: top;
    font-size: 13px;
}
th {
    background: #f2f2f2;
}
.muted {
    color: #666666;
}
.small {
    font-size: 12px;
}
</style>
"@

$htmlSections = @()
$htmlSections += "<div class='card'><h1>Defender for Office 365 Export Report</h1><p class='small'>Generated: $(Get-Date)</p><p class='small'>Tenant: $($orgConfig.DisplayName) ($($orgConfig.Name))</p><p class='small'>JSON companion file: $(Split-Path -Leaf $jsonPath)</p></div>"
$htmlSections += "<div class='card'><h2>Summary</h2>$($summary | ConvertTo-Html -Fragment)</div>"
$htmlSections += "<div class='card'><h2>Tenant Information</h2>$(([pscustomobject]$collections.TenantInfo | ConvertTo-Html -Fragment))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Accepted Domains' -Data $collections.AcceptedDomains -PreferredProperties @('DomainName','Default'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Anti-Phish Policies' -Data $collections.AntiPhishPolicies -PreferredProperties @('Name','Enabled','PhishThresholdLevel','EnableMailboxIntelligence','EnableMailboxIntelligenceProtection','EnableTargetedUserProtection','EnableTargetedDomainsProtection','EnableOrganizationDomainsProtection','EnableSpoofIntelligence','TargetedUserProtectionAction','TargetedUserQuarantineTag','TargetedDomainProtectionAction','TargetedDomainQuarantineTag','MailboxIntelligenceProtectionAction','MailboxIntelligenceQuarantineTag','AuthenticationFailAction','SpoofQuarantineTag','HonorDmarcPolicy'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Anti-Phish Rules' -Data $collections.AntiPhishRules -PreferredProperties @('Name','State','Enabled','Priority','AntiPhishPolicy','RecipientDomainIs','SentToMemberOf'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Inbound Anti-Spam Policies' -Data $collections.AntiSpamInboundPolicies -PreferredProperties @('Name','BulkThreshold','SpamAction','SpamQuarantineTag','HighConfidenceSpamAction','HighConfidenceSpamQuarantineTag','PhishSpamAction','PhishQuarantineTag','HighConfidencePhishAction','HighConfidencePhishQuarantineTag','BulkSpamAction','BulkQuarantineTag','InlineSafetyTipsEnabled','EnableEndUserSpamNotifications','ZapEnabled','MarkAsSpamBulkMail','IncreaseScoreWithImageLinks','IncreaseScoreWithNumericIps','IncreaseScoreWithRedirectToOtherPort','IncreaseScoreWithBizOrInfoUrls','MarkAsSpamEmptyMessages','MarkAsSpamEmbedTagsInHtml','MarkAsSpamJavaScriptInHtml','MarkAsSpamFormTagsInHtml','MarkAsSpamFramesInHtml','MarkAsSpamWebBugsInHtml','MarkAsSpamObjectTagsInHtml','MarkAsSpamSensitiveWordList','MarkAsSpamSpfRecordHardFail','MarkAsSpamFromAddressAuthFail','MarkAsSpamNdrBackscatter'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Inbound Anti-Spam Rules' -Data $collections.AntiSpamInboundRules -PreferredProperties @('Name','State','Enabled','Priority','HostedContentFilterPolicy','RecipientDomainIs','SentToMemberOf'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Outbound Anti-Spam Policies' -Data $collections.AntiSpamOutboundPolicies -PreferredProperties @('Name','RecipientLimitExternalPerHour','RecipientLimitInternalPerHour','RecipientLimitPerDay','ActionWhenThresholdReached','AutoForwardingMode','BccSuspiciousOutboundMail','NotifyOutboundSpam','NotifyOutboundSpamRecipients'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Outbound Anti-Spam Rules' -Data $collections.AntiSpamOutboundRules -PreferredProperties @('Name','State','Enabled','Priority','HostedOutboundSpamFilterPolicy','SenderDomainIs','FromMemberOf'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Safe Links Policies' -Data $collections.SafeLinksPolicies -PreferredProperties @('Name','EnableSafeLinksForEmail','EnableSafeLinksForTeams','EnableForInternalSenders','ScanUrls','DeliverMessageAfterScan','DisableUrlRewrite','TrackClicks','AllowClickThrough','EnableOrganizationBranding'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Safe Links Rules' -Data $collections.SafeLinksRules -PreferredProperties @('Name','State','Enabled','Priority','SafeLinksPolicy','RecipientDomainIs','SentToMemberOf'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Safe Attachments Policies' -Data $collections.SafeAttachmentPolicies -PreferredProperties @('Name','Enable','Enabled','Action','QuarantineTag','Redirect','RedirectAddress'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Safe Attachments Rules' -Data $collections.SafeAttachmentRules -PreferredProperties @('Name','State','Enabled','Priority','SafeAttachmentPolicy','RecipientDomainIs','SentToMemberOf'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Malware Policies' -Data $collections.MalwarePolicies -PreferredProperties @('Name','EnableFileFilter','FileTypes','Action','ZapEnabled','QuarantineTag','EnableInternalSenderAdminNotifications','InternalSenderAdminAddress','EnableExternalSenderAdminNotifications','ExternalSenderAdminAddress'))</div>"

$htmlSections += "<div class='card'>$(New-HtmlTable -Title 'Malware Rules' -Data $collections.MalwareRules -PreferredProperties @('Name','State','Enabled','Priority','MalwareFilterPolicy','RecipientDomainIs','SentToMemberOf'))</div>"

$htmlBody = @"
<html>
<head>
<meta charset='utf-8' />
<title>Defender for Office 365 Export Report</title>
$style
</head>
<body>
$($htmlSections -join "`n")
</body>
</html>
"@

$htmlBody | Out-File -LiteralPath $htmlPath -Encoding utf8
Write-Info "HTML report written to: $htmlPath"
Write-Host ''
Write-Host 'Export complete.' -ForegroundColor Green
Write-Host "JSON : $jsonPath" -ForegroundColor Green
Write-Host "HTML : $htmlPath" -ForegroundColor Green
