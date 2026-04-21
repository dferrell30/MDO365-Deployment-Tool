Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Script:RequiredModuleName = 'ExchangeOnlineManagement'

function Ensure-Module {
  param([Parameter(Mandatory)][string]$Name)

  $installCommand = "Install-Module $Name -Scope CurrentUser -Force -AllowClobber"
  $available = Get-Module -ListAvailable -Name $Name
  if (-not $available) {
    $msg = "Required module '$Name' is not installed.`r`n`r`nRun:`r`n$installCommand"
    Write-Host $msg -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show(
      $msg,
      "Missing PowerShell Module",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Warning
    ) | Out-Null
    return [pscustomobject]@{
      Success = $false
      Message = $msg
      InstallCommand = $installCommand
    }
  }

  try {
    if (-not (Get-Module -Name $Name)) {
      Import-Module $Name -ErrorAction Stop | Out-Null
    }
    return [pscustomobject]@{
      Success = $true
      Message = "Module '$Name' is available."
      InstallCommand = $installCommand
    }
  }
  catch {
    $msg = "Failed to import module '$Name'. $($_.Exception.Message)"
    Write-Host $msg -ForegroundColor Red
    [System.Windows.Forms.MessageBox]::Show(
      $msg,
      "Module Import Error",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    return [pscustomobject]@{
      Success = $false
      Message = $msg
      InstallCommand = $installCommand
    }
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


function Get-ConnectedUserPrincipalName {
  try {
    $conn = Get-ConnectionInformation -ErrorAction Stop | Select-Object -First 1
    if ($conn -and $conn.UserPrincipalName) { return [string]$conn.UserPrincipalName }
  } catch {}
  return $null
}

function Get-TenantDisplayName {
  try {
    $defaultDomain = Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.Default -eq $true } | Select-Object -First 1
    if ($defaultDomain -and $defaultDomain.DomainName) {
      return $defaultDomain.DomainName.ToString()
    }
  } catch {}

  try {
    $initialDomain = Get-AcceptedDomain -ErrorAction Stop | Where-Object { $_.DomainName -like '*.onmicrosoft.com' } | Select-Object -First 1
    if ($initialDomain -and $initialDomain.DomainName) {
      return $initialDomain.DomainName.ToString()
    }
  } catch {}

  try {
    $firstDomain = Get-AcceptedDomain -ErrorAction Stop | Select-Object -First 1
    if ($firstDomain -and $firstDomain.DomainName) {
      return $firstDomain.DomainName.ToString()
    }
  } catch {}

  try {
    $org = Get-OrganizationConfig -ErrorAction Stop
    if ($org -and $org.Name) { return [string]$org.Name }
  } catch {}

  return 'Unknown Tenant'
}

function Update-ConnectionLabel {
  param(
    [Parameter(Mandatory)][System.Windows.Forms.Label]$Label
  )

  if (Test-ExchangeOnlineConnection) {
    $who = Get-ConnectedUserPrincipalName
    $tenant = Get-TenantDisplayName
    if ([string]::IsNullOrWhiteSpace($who)) {
      $Label.Text = "Status: Connected to $tenant"
    } else {
      $Label.Text = "Status: Connected to $tenant as $who"
    }
    $Label.ForeColor = [System.Drawing.Color]::DarkGreen
  } else {
    $Label.Text = "Status: Not Connected"
    $Label.ForeColor = [System.Drawing.Color]::Black
  }
}

function Ensure-ExchangeOnlineAuthenticated {
  param(
    [switch]$ForceReauth,
    [System.Windows.Forms.Label]$ConnectionLabel,
    [scriptblock]$Logger
  )

  $moduleCheck = Ensure-Module -Name $Script:RequiredModuleName
  if (-not $moduleCheck.Success) {
    if ($Logger) { & $Logger "[ERR] $($moduleCheck.Message)" }
    return $false
  }

  try {
    if ($ForceReauth -and (Test-ExchangeOnlineConnection)) {
      Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
      Start-Sleep -Milliseconds 300
    }

    if (-not (Test-ExchangeOnlineConnection)) {
      Connect-ExchangeOnline -ShowBanner:$false -ErrorAction Stop
    }

    if ($ConnectionLabel) {
      Update-ConnectionLabel -Label $ConnectionLabel
    }

    if ($Logger) {
      $who = Get-ConnectedUserPrincipalName
      $tenant = Get-TenantDisplayName
      if ([string]::IsNullOrWhiteSpace($who)) {
        & $Logger "[OK] Connected to tenant: $tenant"
      } else {
        & $Logger "[OK] Connected to tenant: $tenant as $who"
      }
    }
    return $true
  }
  catch {
    $msg = "Connect failed: $($_.Exception.Message)"
    if ($Logger) { & $Logger "[ERR] $msg" }
    [System.Windows.Forms.MessageBox]::Show(
      $msg,
      "Exchange Online Connection Error",
      [System.Windows.Forms.MessageBoxButtons]::OK,
      [System.Windows.Forms.MessageBoxIcon]::Error
    ) | Out-Null
    if ($ConnectionLabel) {
      Update-ConnectionLabel -Label $ConnectionLabel
    }
    return $false
  }
}


function Get-ConnectedTenantDisplay {
  try {
    $upn = (Get-ConnectionInformation -ErrorAction Stop).UserPrincipalName
    if ($upn -and ($upn -match '@')) {
      return ($upn -split '@')[-1]
    }
  }
  catch { }
  return 'Unknown Tenant'
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

function Write-UiStatus {
  param(
    [Parameter(Mandatory)][string]$Message,
    [string]$Color = 'White'
  )

  try { Write-Host $Message -ForegroundColor $Color } catch { }
  try {
    if (Get-Command Log -ErrorAction SilentlyContinue) { Log $Message }
  } catch { }
}

function Disable-RuleOnly {
  param(
    [Parameter(Mandatory)][string]$SetCmdletName,
    [Parameter(Mandatory)][string]$Identity,
    [string]$DisableCmdletName = ''
  )

  if ($DisableCmdletName -and (Get-Command $DisableCmdletName -ErrorAction SilentlyContinue)) {
    & $DisableCmdletName -Identity $Identity -Confirm:$false
    return
  }

  $base = @{ Identity = $Identity }
  Set-RuleEnabled -CmdletName $SetCmdletName -BaseParams $base -Enabled:$false
}

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

# NOTE: Some Exchange Online rules may default to Enabled on creation. Rules are explicitly set to Disabled during deployment.
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
    Write-UiStatus "Safe Links policy '$Name' already exists. Updating settings..." "Yellow"
    $p = @{ Identity = $Name }
    foreach ($kv in $settings.GetEnumerator()) { if (Supports-Param 'Set-SafeLinksPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    Set-SafeLinksPolicy @p
  }
}

function Ensure-SafeLinksRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-SafeLinksRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-UiStatus "Safe Links rule '$RuleName' does not exist. Creating it disabled..." "Cyan"
    $params = @{
      Name = $RuleName
      SafeLinksPolicy = $PolicyName
      RecipientDomainIs = $RecipientDomains
    }
    if (Supports-Param 'New-SafeLinksRule' 'Enabled') { $params['Enabled'] = $false }
    New-SafeLinksRule @params
  } else {
    Write-UiStatus "SafeLinksRule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-SafeLinksRule' -Identity $RuleName
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
    Write-UiStatus "Safe Attachments policy '$Name' does not exist. Creating it..." "Cyan"
    $p = @{ Name = $Name }
    $p = Add-EnableParam $p $true 'New-SafeAttachmentPolicy' $null
    foreach ($kv in $settings.GetEnumerator()) { if (Supports-Param 'New-SafeAttachmentPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    New-SafeAttachmentPolicy @p
  }
  else {
    Write-UiStatus "Safe Attachments policy '$Name' already exists. Updating settings..." "Yellow"
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
    Write-UiStatus "Safe Attachments rule '$RuleName' does not exist. Creating it disabled..." "Cyan"
    $params = @{
      Name = $RuleName
      SafeAttachmentPolicy = $PolicyName
      RecipientDomainIs = $RecipientDomains
    }
    if (Supports-Param 'New-SafeAttachmentRule' 'Enabled') { $params['Enabled'] = $false }
    New-SafeAttachmentRule @params
  } else {
    Write-UiStatus "SafeAttachmentsRule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-SafeAttachmentRule' -Identity $RuleName
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
    Write-UiStatus "Anti-Phish policy '$Name' does not exist. Creating it..." "Cyan"
    $p = @{ Name = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'New-AntiPhishPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    New-AntiPhishPolicy @p
  } else {
    Write-UiStatus "Anti-Phish policy '$Name' already exists. Updating settings..." "Yellow"
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-AntiPhishPolicy' $kv.Key) { $p[$kv.Key] = $kv.Value } }
    Set-AntiPhishPolicy @p
  }
}

function Ensure-AntiPhishRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-AntiPhishRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-UiStatus "Anti-Phish rule '$RuleName' does not exist. Creating it disabled..." "Cyan"
    $params = @{
      Name = $RuleName
      AntiPhishPolicy = $PolicyName
      RecipientDomainIs = $RecipientDomains
    }
    if (Supports-Param 'New-AntiPhishRule' 'Enabled') { $params['Enabled'] = $false }
    New-AntiPhishRule @params
  } else {
    Write-UiStatus "AntiPhishRule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-AntiPhishRule' -Identity $RuleName
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
    Write-UiStatus "Inbound Anti-Spam policy '$Name' already exists. Updating settings..." "Yellow"
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-HostedContentFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    Set-HostedContentFilterPolicy @p
  }
}

function Ensure-AntiSpamInboundRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-HostedContentFilterRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-UiStatus "Inbound Anti-Spam rule '$RuleName' does not exist. Creating it disabled..." "Cyan"
    $params = @{
      Name = $RuleName
      HostedContentFilterPolicy = $PolicyName
      RecipientDomainIs = $RecipientDomains
    }
    if (Supports-Param 'New-HostedContentFilterRule' 'Enabled') { $params['Enabled'] = $false }
    New-HostedContentFilterRule @params
  } else {
    Write-UiStatus "AntiSpamInboundRule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-HostedContentFilterRule' -Identity $RuleName
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
    Write-UiStatus "Outbound Anti-Spam policy '$Name' already exists. Updating settings..." "Yellow"
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-HostedOutboundSpamFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    Set-HostedOutboundSpamFilterPolicy @p
  }
}

function Ensure-AntiSpamOutboundRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$SenderDomains)
  $rule = Get-HostedOutboundSpamFilterRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-UiStatus "Outbound Anti-Spam rule '$RuleName' does not exist. Creating it disabled..." "Cyan"
    $params = @{
      Name = $RuleName
      HostedOutboundSpamFilterPolicy = $PolicyName
      SenderDomainIs = $SenderDomains
    }
    if (Supports-Param 'New-HostedOutboundSpamFilterRule' 'Enabled') { $params['Enabled'] = $false }
    New-HostedOutboundSpamFilterRule @params
  } else {
    Write-UiStatus "AntiSpamOutboundRule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-HostedOutboundSpamFilterRule' -Identity $RuleName -DisableCmdletName 'Disable-HostedOutboundSpamFilterRule'
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
    Write-UiStatus "Anti-Malware policy '$Name' already exists. Updating settings..." "Yellow"
    $p = @{ Identity = $Name }
    foreach ($kv in $vals.GetEnumerator()) { if (Supports-Param 'Set-MalwareFilterPolicy' $kv.Key) { $p[$kv.Key]=$kv.Value } }
    Set-MalwareFilterPolicy @p
  }
}

function Ensure-AntiMalwareRuleGlobal {
  param([string]$RuleName,[string]$PolicyName,[string[]]$RecipientDomains)
  $rule = Get-MalwareFilterRule -ErrorAction SilentlyContinue | Where-Object Name -eq $RuleName
  if (-not $rule) {
    Write-UiStatus "Anti-Malware rule '$RuleName' does not exist. Creating it disabled..." "Cyan"
    $params = @{
      Name = $RuleName
      MalwareFilterPolicy = $PolicyName
      RecipientDomainIs = $RecipientDomains
    }
    if (Supports-Param 'New-MalwareFilterRule' 'Enabled') { $params['Enabled'] = $false }
    New-MalwareFilterRule @params
  } else {
    Write-UiStatus "AntiMalwareRule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-MalwareFilterRule' -Identity $RuleName -DisableCmdletName 'Disable-MalwareFilterRule'
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
$form.Text = "DFO365 Deployment Tool - V1"
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
    
  } catch { Log "[ERR] Connect failed: $($_.Exception.Message)" }
})
$form.Controls.Add($btnConnect)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = "Disconnect"
$btnDisconnect.Size = New-Object System.Drawing.Size(280,40)
$btnDisconnect.Location = New-Object System.Drawing.Point(310,10)
$btnDisconnect.Add_Click({
  try {
    if (Test-ExchangeOnlineConnection) {
      Disconnect-ExchangeOnline -Confirm:$false -ErrorAction Stop
    }
    Update-ConnectionLabel -Label $lblConnection
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
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
      if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
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
$form.Add_Shown({
  $form.Activate()
  Update-ConnectionLabel -Label $lblConnection
})
[void]$form.ShowDialog()
