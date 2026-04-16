Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Script:RequiredModuleName = 'ExchangeOnlineManagement'

function Ensure-Module {
  param([Parameter(Mandatory)][string]$Name)

  $available = Get-Module -ListAvailable -Name $Name
  if (-not $available) {
    Write-Host "Required module '$Name' is not installed." -ForegroundColor Red
    Write-Host "Install it with:" -ForegroundColor Yellow
    Write-Host "Install-Module $Name -Scope CurrentUser -Force -AllowClobber" -ForegroundColor Cyan
    throw "Missing required module: $Name"
  }

  try {
    Import-Module $Name -ErrorAction Stop | Out-Null
  }
  catch {
    throw "Failed to import module '$Name'. $($_.Exception.Message)"
  }
}

function Test-ExchangeOnlineConnection {
  try {
    $null = Get-ConnectionInformation -ErrorAction Stop
    return $true
  }
  catch {
    return $false
  }
}

function Get-AllAcceptedDomains {
  try {
    (Get-AcceptedDomain -ErrorAction Stop |
      ForEach-Object { $_.DomainName.ToString() } |
      Where-Object { $_ } | Sort-Object -Unique)
  } catch { @() }
}

function Supports-Param($CommandName, $ParamName) {
  $cmd = Get-Command $CommandName -ErrorAction SilentlyContinue
  return [bool]($cmd -and $cmd.Parameters.ContainsKey($ParamName))
}

# Enable helper: uses -Enabled if present, else -State Enabled/Disabled; else just sets base params.
function Set-RuleEnabled {
  param(
    [Parameter(Mandatory)][string]$CmdletName,  # e.g. 'Set-SafeLinksRule'
    [Parameter(Mandatory)][hashtable]$BaseParams,
    [Parameter()][bool]$Enabled = $true
  )
  if (Supports-Param $CmdletName 'Enabled') {
    & $CmdletName @BaseParams -Enabled:$Enabled
  }
  elseif (Supports-Param $CmdletName 'State') {
    & $CmdletName @BaseParams -State ($(if ($Enabled) {'Enabled'} else {'Disabled'}))
  }
  else {
    & $CmdletName @BaseParams
  }
}

# ---- Baseline Ensurers (version-aware) ----
$Names = [ordered]@{
  SafeLinksPolicy           = 'Microsoft-Zero-Trust-SafeLinks-Rule'
  SafeLinksRule             = 'Microsoft-Zero-Trust-SafeLinks-Rule'
  SafeAttachmentsPolicy     = 'Microsoft-Zero-Trust-SafeAttachments'
  SafeAttachmentsRule       = 'Microsoft-Zero-Trust-SafeAttachments-Rule'
  AntiPhishPolicy           = 'Microsoft-Zero-Trust-AntiPhish'
  AntiPhishRule             = 'Microsoft-Zero-Trust-AntiPhish-Rule'
  AntiSpamInboundPolicy     = 'Microsoft-Zero-Trust-AntiSpam-Inbound'
  AntiSpamInboundRule       = 'Microsoft-Zero-Trust-AntiSpam-Inbound-Rule'
  AntiSpamOutboundPolicy    = 'Microsoft-Zero-Trust-AntiSpam-Outbound'
  AntiSpamOutboundRule      = 'Microsoft-Zero-Trust-AntiSpam-Outbound-Rule'
  AntiMalwarePolicy         = 'Microsoft-Zero-Trust-AntiMalware'
  AntiMalwareRule           = 'Microsoft-Zero-Trust-AntiMalware-Rule'
}

$AdminNotify = 'postmaster@yourdomain.com'  # change for malware notifications

function Ensure-SafeLinksPolicy {
  param([string]$Name)
  $settings = [ordered]@{
    EnableSafeLinksForEmail   = $true
    EnableSafeLinksForTeams   = $true
    EnableForInternalSenders  = $true
    ScanUrls                  = $true
    DeliverMessageAfterScan   = $true
    DisableUrlRewrite         = $false
    TrackClicks               = $true
    AllowClickThrough         = $false
    EnableOrganizationBranding = $false
  }

  $exists = Get-SafeLinksPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  if (-not $exists) {
    $p = @{ Name = $Name }
    foreach ($kv in $settings.GetEnumerator()) { if (Supports-Param 'New-SafeLinksPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    New-SafeLinksPolicy @p
  } else {
    Write-Host "Anti-Phish policy '$Name' already exists. Updating settings..." -ForegroundColor Yellow
    $p = @{ Identity = $Name }
    foreach ($kv in $settings.GetEnumerator()) { if (Supports-Param 'Set-SafeLinksPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    Set-SafeLinksPolicy @p
  }
}

function Ensure-SafeLinksRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-SafeLinksRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-Host "Safe Links rule '$RuleName' does not exist. Creating it disabled..." -ForegroundColor Cyan
    New-SafeLinksRule -Name $RuleName -SafeLinksPolicy $PolicyName -RecipientDomainIs $RecipientDomains
  } else {
    Write-Host "Safe Links rule '$RuleName' already exists. Updating settings and keeping it disabled..." -ForegroundColor Yellow
  }
  $bp = @{ Identity = $RuleName; SafeLinksPolicy = $PolicyName; RecipientDomainIs = $RecipientDomains }
  Set-RuleEnabled -CmdletName 'Set-SafeLinksRule' -BaseParams $bp -Enabled:$false
}

# ----- SAFE ATTACHMENTS (hardened) -----
function Ensure-SafeAttachmentsPolicy {
  param([string]$Name)

  function Add-EnableParam([hashtable]$h, [bool]$on=$true, [string]$newCmd, [string]$setCmd) {
    if ($newCmd -and (Supports-Param $newCmd 'Enable'))      { $h['Enable']  = $on }
    elseif ($newCmd -and (Supports-Param $newCmd 'Enabled')) { $h['Enabled'] = $on }
    elseif ($setCmd -and (Supports-Param $setCmd 'Enable'))  { $h['Enable']  = $on }
    elseif ($setCmd -and (Supports-Param $setCmd 'Enabled')) { $h['Enabled'] = $on }
    return $h
  }

  $settings = [ordered]@{
    Action         = 'Block'
    QuarantineTag  = 'AdminOnlyAccessPolicy'
    Redirect       = $false
  }

  $existing = Get-SafeAttachmentPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name

  if (-not $existing) {
    Write-Host "Creating Safe Attachments policy '$Name'" -ForegroundColor Cyan
    $p = @{ Name = $Name }
    $p = Add-EnableParam $p $true 'New-SafeAttachmentPolicy' $null
    foreach ($kv in $settings.GetEnumerator()) { if (Supports-Param 'New-SafeAttachmentPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    New-SafeAttachmentPolicy @p
  }
  else {
    Write-Host "Updating Safe Attachments policy '$Name'" -ForegroundColor DarkGray
    $sp = @{ Identity = $Name }
    $sp = Add-EnableParam $sp $true $null 'Set-SafeAttachmentPolicy'
    foreach ($kv in $settings.GetEnumerator()) { if (Supports-Param 'Set-SafeAttachmentPolicy' $kv.Key) { $sp[$kv.Key] = $kv.Value } }
    Set-SafeAttachmentPolicy @sp
  }
}

function Ensure-SafeAttachmentsRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-SafeAttachmentRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-Host "Safe Attachments rule '$RuleName' does not exist. Creating it disabled..." -ForegroundColor Cyan
    New-SafeAttachmentRule -Name $RuleName -SafeAttachmentPolicy $PolicyName -RecipientDomainIs $RecipientDomains
  } else {
    Write-Host "Safe Attachments rule '$RuleName' already exists. Updating settings and keeping it disabled..." -ForegroundColor Yellow
  }
  $bp = @{ Identity = $RuleName; SafeAttachmentPolicy = $PolicyName; RecipientDomainIs = $RecipientDomains }
  Set-RuleEnabled -CmdletName 'Set-SafeAttachmentRule' -BaseParams $bp -Enabled:$false
}

function Ensure-AntiPhishPolicy {
  param([string]$Name)
  $policy = Get-AntiPhishPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name

  $vals = [ordered]@{
    EnableMailboxIntelligence            = $true
    EnableMailboxIntelligenceProtection  = $true
    MailboxIntelligenceProtectionAction  = 'Quarantine'
    MailboxIntelligenceQuarantineTag     = 'AdminOnlyAccessPolicy'
    EnableOrganizationDomainsProtection  = $true
    EnableSpoofIntelligence              = $true
    EnableTargetedUserProtection         = $true
    TargetedUserProtectionAction         = 'Quarantine'
    TargetedUserQuarantineTag            = 'AdminOnlyAccessPolicy'
    EnableTargetedDomainsProtection      = $true
    TargetedDomainProtectionAction       = 'Quarantine'
    TargetedDomainQuarantineTag          = 'AdminOnlyAccessPolicy'
    EnableFirstContactSafetyTips         = $true
    EnableSimilarUsersSafetyTips         = $true
    EnableSimilarDomainsSafetyTips       = $true
    EnableUnusualCharactersSafetyTips    = $true
    EnableUnauthenticatedSender          = $true
    EnableViaTag                         = $true
    HonorDmarcPolicy                     = $true
    AuthenticationFailAction             = 'Quarantine'
    SpoofQuarantineTag                   = 'AdminOnlyAccessPolicy'
    PhishThresholdLevel                  = 3
  }

  if (-not $policy) {
    Write-Host "Anti-Phish policy '$Name' does not exist. Creating it..." -ForegroundColor Cyan
    $p = @{ Name = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'New-AntiPhishPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    New-AntiPhishPolicy @p
  } else {
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-AntiPhishPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    Set-AntiPhishPolicy @p
  }
}

function Ensure-AntiPhishRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-AntiPhishRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-Host "Anti-Phish rule '$RuleName' does not exist. Creating it disabled..." -ForegroundColor Cyan
    New-AntiPhishRule -Name $RuleName -AntiPhishPolicy $PolicyName -RecipientDomainIs $RecipientDomains
  } else {
    Write-Host "Anti-Phish rule '$RuleName' already exists. Updating settings and keeping it disabled..." -ForegroundColor Yellow
  }
  $bp = @{ Identity = $RuleName; AntiPhishPolicy = $PolicyName; RecipientDomainIs = $RecipientDomains }
  Set-RuleEnabled -CmdletName 'Set-AntiPhishRule' -BaseParams $bp -Enabled:$false
}

function Ensure-AntiSpamInboundPolicy {
  param([string]$Name)
  $policy = Get-HostedContentFilterPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  $vals = [ordered]@{
    BulkThreshold                     = 5
    SpamAction                        = 'Quarantine'
    SpamQuarantineTag                 = 'DefaultFullAccesswithNotificationPolicy'
    HighConfidenceSpamAction          = 'Quarantine'
    HighConfidenceSpamQuarantineTag   = 'DefaultFullAccesswithNotificationPolicy'
    BulkSpamAction                    = 'Quarantine'
    BulkQuarantineTag                 = 'DefaultFullAccesswithNotificationPolicy'
    PhishSpamAction                   = 'Quarantine'
    PhishQuarantineTag                = 'AdminOnlyAccessPolicy'
    HighConfidencePhishAction         = 'Quarantine'
    HighConfidencePhishQuarantineTag  = 'AdminOnlyAccessPolicy'
    InlineSafetyTipsEnabled           = $true
    SpamZapEnabled                    = $true
    PhishZapEnabled                   = $true
    IncreaseScoreWithImageLinks       = 'On'
    IncreaseScoreWithNumericIps       = 'On'
    IncreaseScoreWithRedirectToOtherPort = 'On'
    IncreaseScoreWithBizOrInfoUrls    = 'On'
    MarkAsSpamEmptyMessages           = 'On'
    MarkAsSpamEmbedTagsInHtml         = 'On'
    MarkAsSpamJavaScriptInHtml        = 'On'
    MarkAsSpamFormTagsInHtml          = 'On'
    MarkAsSpamFramesInHtml            = 'On'
    MarkAsSpamWebBugsInHtml           = 'On'
    MarkAsSpamObjectTagsInHtml        = 'On'
    MarkAsSpamSensitiveWordList       = 'Off'
    MarkAsSpamSpfRecordHardFail       = 'On'
    MarkAsSpamFromAddressAuthFail     = 'On'
    MarkAsSpamNdrBackscatter          = 'On'
  }
  if (-not $policy) {
    $p = @{ Name = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'New-HostedContentFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    New-HostedContentFilterPolicy @p
  } else {
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-HostedContentFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    Set-HostedContentFilterPolicy @p
  }
}

function Ensure-AntiSpamInboundRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-HostedContentFilterRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-Host "Inbound Anti-Spam rule '$RuleName' does not exist. Creating it disabled..." -ForegroundColor Cyan
    New-HostedContentFilterRule -Name $RuleName -HostedContentFilterPolicy $PolicyName -RecipientDomainIs $RecipientDomains
  } else {
    Write-Host "Inbound Anti-Spam rule '$RuleName' already exists. Updating settings and keeping it disabled..." -ForegroundColor Yellow
  }
  $bp = @{ Identity = $RuleName; HostedContentFilterPolicy = $PolicyName; RecipientDomainIs = $RecipientDomains }
  Set-RuleEnabled -CmdletName 'Set-HostedContentFilterRule' -BaseParams $bp -Enabled:$false
}

function Ensure-AntiSpamOutboundPolicy {
  param([string]$Name,[string]$NotifyAddress)
  $policy = Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  $vals = [ordered]@{
    RecipientLimitExternalPerHour = 400
    RecipientLimitInternalPerHour = 800
    RecipientLimitPerDay          = 800
    ActionWhenThresholdReached    = 'BlockUser'
    AutoForwardingMode            = 'Off'
    BccSuspiciousOutboundMail     = $false
    NotifyOutboundSpam            = $true
    NotifyOutboundSpamRecipients  = $NotifyAddress
  }
  if (-not $policy) {
    $p = @{ Name = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'New-HostedOutboundSpamFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    New-HostedOutboundSpamFilterPolicy @p
  } else {
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-HostedOutboundSpamFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    Set-HostedOutboundSpamFilterPolicy @p
  }
}

function Ensure-AntiSpamOutboundRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$SenderDomains)
  $rule = Get-HostedOutboundSpamFilterRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-Host "Outbound Anti-Spam rule '$RuleName' does not exist. Creating it disabled..." -ForegroundColor Cyan
    New-HostedOutboundSpamFilterRule -Name $RuleName -HostedOutboundSpamFilterPolicy $PolicyName -SenderDomainIs $SenderDomains
  } else {
    Write-Host "Outbound Anti-Spam rule '$RuleName' already exists. Updating settings and keeping it disabled..." -ForegroundColor Yellow
    $params = @{ Identity = $RuleName; HostedOutboundSpamFilterPolicy = $PolicyName; SenderDomainIs = $SenderDomains }
    Set-HostedOutboundSpamFilterRule @params
  }
  if (Get-Command Disable-HostedOutboundSpamFilterRule -ErrorAction SilentlyContinue) {
    Disable-HostedOutboundSpamFilterRule -Identity $RuleName -Confirm:$false
  } elseif (Supports-Param 'Set-HostedOutboundSpamFilterRule' 'State') {
    Set-HostedOutboundSpamFilterRule -Identity $RuleName -State Disabled
  } elseif (Supports-Param 'Set-HostedOutboundSpamFilterRule' 'Enabled') {
    Set-HostedOutboundSpamFilterRule -Identity $RuleName -Enabled:$false
  }
}

function Ensure-AntiMalwarePolicy {
  param([string]$Name,[string]$AdminNotify)
  $policy = Get-MalwareFilterPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  $vals = @{
    EnableInternalSenderAdminNotifications=$true; InternalSenderAdminAddress=$AdminNotify;
    Action='DeleteMessage'; EnableZeroHourAutoPurge=$true
  }
  if (-not $policy) {
    $p = @{ Name = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'New-MalwareFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    New-MalwareFilterPolicy @p
  } else {
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-MalwareFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    Set-MalwareFilterPolicy @p
  }
}

function Ensure-AntiMalwareRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-MalwareFilterRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-Host "Anti-Malware rule '$RuleName' does not exist. Creating it disabled..." -ForegroundColor Cyan
    New-MalwareFilterRule -Name $RuleName -MalwareFilterPolicy $PolicyName -RecipientDomainIs $RecipientDomains
  } else {
    Write-Host "Anti-Malware rule '$RuleName' already exists. Updating settings and keeping it disabled..." -ForegroundColor Yellow
  }
  $bp = @{ Identity = $RuleName; MalwareFilterPolicy = $PolicyName; RecipientDomainIs = $RecipientDomains }
  Set-RuleEnabled -CmdletName 'Set-MalwareFilterRule' -BaseParams $bp -Enabled:$false
}

function Export-PoliciesJson {
  param([string]$Path)
  $items = @()
  $items += Get-SafeLinksPolicy          -ErrorAction SilentlyContinue
  $items += Get-SafeLinksRule            -ErrorAction SilentlyContinue
  $items += Get-SafeAttachmentPolicy     -ErrorAction SilentlyContinue
  $items += Get-SafeAttachmentRule       -ErrorAction SilentlyContinue
  $items += Get-AntiPhishPolicy          -ErrorAction SilentlyContinue
  $items += Get-AntiPhishRule            -ErrorAction SilentlyContinue
  $items += Get-HostedContentFilterPolicy -ErrorAction SilentlyContinue
  $items += Get-HostedContentFilterRule   -ErrorAction SilentlyContinue
  $items += Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue
  $items += Get-HostedOutboundSpamFilterRule   -ErrorAction SilentlyContinue
  $items += Get-MalwareFilterPolicy      -ErrorAction SilentlyContinue
  $items += Get-MalwareFilterRule        -ErrorAction SilentlyContinue

  if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
  $i = 0
  foreach ($obj in $items) {
    $i++
    $name = ($obj.Name | ForEach-Object { $_ }) -join '_' ; if (-not $name) { $name = "item$i" }
    $file = Join-Path $Path ("{0}_{1}.json" -f $obj.GetType().Name, ($name -replace '[^\w\-]','_'))
    $obj | ConvertTo-Json -Depth 10 | Out-File -FilePath $file -Encoding UTF8
  }
}

# -------------------- GUI --------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "Defender for Office 365 Management"
$form.Size = New-Object System.Drawing.Size(620,570)
$form.StartPosition = "CenterScreen"

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$statusBox.Size = New-Object System.Drawing.Size(580,220)
$statusBox.Location = New-Object System.Drawing.Point(10,300)
$form.Controls.Add($statusBox)
function Log($msg){ $statusBox.AppendText("$msg`r`n") }

$lblConnection = New-Object System.Windows.Forms.Label
$lblConnection.Text = "Status: Not Connected"
$lblConnection.Size = New-Object System.Drawing.Size(580,20)
$lblConnection.Location = New-Object System.Drawing.Point(10,270)
$lblConnection.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblConnection)

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog

# Connect / Disconnect
$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect (Exchange Online)"
$btnConnect.Size = New-Object System.Drawing.Size(280,40)
$btnConnect.Location = New-Object System.Drawing.Point(10,10)
$btnConnect.Add_Click({
  try {
    Ensure-Module -Name ExchangeOnlineManagement
    Connect-ExchangeOnline -ShowBanner:$false
    $who = (Get-ConnectionInformation).UserPrincipalName
    $lblConnection.Text = "Status: Connected as $who"
    Log "[OK] Connected as $who"
  } catch { Log "[ERR] Connect failed: $($_.Exception.Message)" }
})
$form.Controls.Add($btnConnect)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = "Disconnect"
$btnDisconnect.Size = New-Object System.Drawing.Size(280,40)
$btnDisconnect.Location = New-Object System.Drawing.Point(310,10)
$btnDisconnect.Add_Click({
  try {
    Disconnect-ExchangeOnline -Confirm:$false
    $lblConnection.Text = "Status: Not Connected"
    Log "[OK] Disconnected."
  }
  catch {
    Log "[ERR] Disconnect failed: $($_.Exception.Message)"
  }
})  # <-- this was missing in your file
$form.Controls.Add($btnDisconnect)

# Row 2: Build Baseline & Export
$btnQuickBuild = New-Object System.Windows.Forms.Button
$btnQuickBuild.Text = "Quick Build: All Baselines"
$btnQuickBuild.Size = New-Object System.Drawing.Size(280,40)
$btnQuickBuild.Location = New-Object System.Drawing.Point(10,60)
$btnQuickBuild.Add_Click({
  try {
    $dom = Get-AllAcceptedDomains
    Ensure-SafeLinksPolicy -Name $Names.SafeLinksPolicy
    Ensure-SafeLinksRuleGlobal -RuleName $Names.SafeLinksRule -PolicyName $Names.SafeLinksPolicy -RecipientDomains $dom
    Ensure-SafeAttachmentsPolicy -Name $Names.SafeAttachmentsPolicy
    Ensure-SafeAttachmentsRuleGlobal -RuleName $Names.SafeAttachmentsRule -PolicyName $Names.SafeAttachmentsPolicy -RecipientDomains $dom
    Ensure-AntiPhishPolicy -Name $Names.AntiPhishPolicy
    Ensure-AntiPhishRuleGlobal -RuleName $Names.AntiPhishRule -PolicyName $Names.AntiPhishPolicy -RecipientDomains $dom
    Ensure-AntiSpamInboundPolicy -Name $Names.AntiSpamInboundPolicy
    Ensure-AntiSpamInboundRuleGlobal -RuleName $Names.AntiSpamInboundRule -PolicyName $Names.AntiSpamInboundPolicy -RecipientDomains $dom
    Ensure-AntiSpamOutboundPolicy -Name $Names.AntiSpamOutboundPolicy -NotifyAddress $AdminNotify
    Ensure-AntiSpamOutboundRuleGlobal -RuleName $Names.AntiSpamOutboundRule -PolicyName $Names.AntiSpamOutboundPolicy -SenderDomains $dom
    Ensure-AntiMalwarePolicy -Name $Names.AntiMalwarePolicy -AdminNotify $AdminNotify
    Ensure-AntiMalwareRuleGlobal -RuleName $Names.AntiMalwareRule -PolicyName $Names.AntiMalwarePolicy -RecipientDomains $dom
    Log "[OK] Quick Build complete."
  } catch { Log "[ERR] Quick Build error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnQuickBuild)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Current Policies (JSON)"
$btnExport.Size = New-Object System.Drawing.Size(280,40)
$btnExport.Location = New-Object System.Drawing.Point(310,60)
$btnExport.Add_Click({
  if ($folderDialog.ShowDialog() -eq "OK") {
    try {
      Export-PoliciesJson -Path $folderDialog.SelectedPath
      Log "[OK] Exported to $($folderDialog.SelectedPath)"
    } catch { Log "[ERR] Export failed: $($_.Exception.Message)" }
  }
})
$form.Controls.Add($btnExport)

# Row 3: Safe Links / Safe Attachments
$btnSL = New-Object System.Windows.Forms.Button
$btnSL.Text = "Safe Links: Create/Update Policy + Global Rule"
$btnSL.Size = New-Object System.Drawing.Size(580,40)
$btnSL.Location = New-Object System.Drawing.Point(10,110)
$btnSL.Add_Click({
  try {
    $dom = Get-AllAcceptedDomains
    Ensure-SafeLinksPolicy -Name $Names.SafeLinksPolicy
    Ensure-SafeLinksRuleGlobal -RuleName $Names.SafeLinksRule -PolicyName $Names.SafeLinksPolicy -RecipientDomains $dom
    Log "[OK] Safe Links baseline ensured."
  } catch { Log "[ERR] Safe Links error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnSL)

$btnSA = New-Object System.Windows.Forms.Button
$btnSA.Text = "Safe Attachments: Create/Update Policy + Global Rule"
$btnSA.Size = New-Object System.Drawing.Size(580,40)
$btnSA.Location = New-Object System.Drawing.Point(10,160)
$btnSA.Add_Click({
  try {
    $dom = Get-AllAcceptedDomains
    Ensure-SafeAttachmentsPolicy -Name $Names.SafeAttachmentsPolicy
    Ensure-SafeAttachmentsRuleGlobal -RuleName $Names.SafeAttachmentsRule -PolicyName $Names.SafeAttachmentsPolicy -RecipientDomains $dom
    Log "[OK] Safe Attachments baseline ensured."
  } catch { Log "[ERR] Safe Attachments error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnSA)

# Row 4: Anti-Phish / Anti-Spam / Anti-Malware
$btnAPh = New-Object System.Windows.Forms.Button
$btnAPh.Text = "Anti-Phish: Create/Update Policy + Global Rule"
$btnAPh.Size = New-Object System.Drawing.Size(180,40)
$btnAPh.Location = New-Object System.Drawing.Point(10,210)
$btnAPh.Add_Click({
  try {
    $dom = Get-AllAcceptedDomains
    Ensure-AntiPhishPolicy -Name $Names.AntiPhishPolicy
    Ensure-AntiPhishRuleGlobal -RuleName $Names.AntiPhishRule -PolicyName $Names.AntiPhishPolicy -RecipientDomains $dom
    Log "[OK] Anti-Phish baseline ensured."
  } catch { Log "[ERR] Anti-Phish error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnAPh)

$btnASp = New-Object System.Windows.Forms.Button
$btnASp.Text = "Anti-Spam: Inbound + Outbound Policies + Rules"
$btnASp.Size = New-Object System.Drawing.Size(180,40)
$btnASp.Location = New-Object System.Drawing.Point(205,210)
$btnASp.Add_Click({
  try {
    $dom = Get-AllAcceptedDomains
    Ensure-AntiSpamInboundPolicy -Name $Names.AntiSpamInboundPolicy
    Ensure-AntiSpamInboundRuleGlobal -RuleName $Names.AntiSpamInboundRule -PolicyName $Names.AntiSpamInboundPolicy -RecipientDomains $dom
    Ensure-AntiSpamOutboundPolicy -Name $Names.AntiSpamOutboundPolicy -NotifyAddress $AdminNotify
    Ensure-AntiSpamOutboundRuleGlobal -RuleName $Names.AntiSpamOutboundRule -PolicyName $Names.AntiSpamOutboundPolicy -SenderDomains $dom
    Log "[OK] Anti-Spam baseline ensured."
  } catch { Log "[ERR] Anti-Spam error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnASp)

$btnAMw = New-Object System.Windows.Forms.Button
$btnAMw.Text = "Anti-Malware: Policy + Global Rule"
$btnAMw.Size = New-Object System.Drawing.Size(180,40)
$btnAMw.Location = New-Object System.Drawing.Point(400,210)
$btnAMw.Add_Click({
  try {
    $dom = Get-AllAcceptedDomains
    Ensure-AntiMalwarePolicy -Name $Names.AntiMalwarePolicy -AdminNotify $AdminNotify
    Ensure-AntiMalwareRuleGlobal -RuleName $Names.AntiMalwareRule -PolicyName $Names.AntiMalwarePolicy -RecipientDomains $dom
    Log "[OK] Anti-Malware baseline ensured."
  } catch { Log "[ERR] Anti-Malware error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnAMw)

# Row 5: Safe Links URL management (lists)
Add-Type -AssemblyName Microsoft.VisualBasic
$btnSLUrls = New-Object System.Windows.Forms.Button
$btnSLUrls.Text = "Safe Links: Manage URL lists (Block/DoNotRewrite/Disabled)"
$btnSLUrls.Size = New-Object System.Drawing.Size(580,40)
$btnSLUrls.Location = New-Object System.Drawing.Point(10,255)
$btnSLUrls.Add_Click({
  try {
    $policyName = $Names.SafeLinksPolicy
    $mode = [System.Windows.Forms.MessageBox]::Show(
      "Choose YES=Block, NO=DoNotRewrite, Cancel=Disabled list",
      "Safe Links URL List",
      [System.Windows.Forms.MessageBoxButtons]::YesNoCancel
    )
    if ($mode -eq [System.Windows.Forms.DialogResult]::Cancel) { 
      $target = 'DisabledUrls'
    } elseif ($mode -eq [System.Windows.Forms.DialogResult]::Yes) {
      $target = 'BlockedUrls'
    } else {
      $target = 'DoNotRewriteUrls'
    }

    $urls = [Microsoft.VisualBasic.Interaction]::InputBox(
      "Enter URLs separated by commas",
      "Safe Links URLs",
      "http://example.com"
    )

    if (-not [string]::IsNullOrWhiteSpace($urls)) {
      $arr = $urls -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
      $current = (Get-SafeLinksPolicy -Identity $policyName | Select-Object -ExpandProperty $target -ErrorAction SilentlyContinue)
      $new = @()
      if ($current) { $new += $current }
      $new += $arr
      $new = $new | Sort-Object -Unique
      $p = @{ Identity = $policyName }
      $p[$target] = $new
      Set-SafeLinksPolicy @p
      Log ("[OK] {0} updated on '{1}'." -f $target, $policyName)
    }
  } catch {
    Log "[ERR] Safe Links list update error: $($_.Exception.Message)"
  }
})
$form.Controls.Add($btnSLUrls)

# Show Form
$form.Topmost = $true
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()
