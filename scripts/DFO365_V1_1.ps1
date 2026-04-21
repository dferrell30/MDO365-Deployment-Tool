
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$Script:RequiredModuleName = 'ExchangeOnlineManagement'
$Script:Config = $null
$Script:LoadedConfigPath = $null
$Script:ProfileConfigMap = @{}

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
  $upn = Get-ConnectedUserPrincipalName
  if ($upn -and ($upn -match '@')) {
    return (($upn -split '@')[-1]).ToLower()
  }
  return 'Not Connected'
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
        & $Logger "[OK] Connected to $tenant"
      } else {
        & $Logger "[OK] Connected to $tenant as $who"
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
    [Parameter(Mandatory)][string]$CmdletName,
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

function ConvertTo-Hashtable {
  param([Parameter(Mandatory)]$InputObject)
  $hash = [ordered]@{}
  if ($null -eq $InputObject) { return $hash }

  if ($InputObject -is [System.Collections.IDictionary]) {
    foreach ($key in $InputObject.Keys) {
      $value = $InputObject[$key]
      if ($value -is [System.Management.Automation.PSCustomObject] -or $value -is [System.Collections.IDictionary]) {
        $hash[$key] = ConvertTo-Hashtable -InputObject $value
      } elseif ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
        $items = @()
        foreach ($item in $value) {
          if ($item -is [System.Management.Automation.PSCustomObject] -or $item -is [System.Collections.IDictionary]) {
            $items += ,(ConvertTo-Hashtable -InputObject $item)
          } else {
            $items += ,$item
          }
        }
        $hash[$key] = $items
      } else {
        $hash[$key] = $value
      }
    }
    return $hash
  }

  foreach ($prop in $InputObject.PSObject.Properties) {
    $value = $prop.Value
    if ($value -is [System.Management.Automation.PSCustomObject] -or $value -is [System.Collections.IDictionary]) {
      $hash[$prop.Name] = ConvertTo-Hashtable -InputObject $value
    } elseif ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
      $items = @()
      foreach ($item in $value) {
        if ($item -is [System.Management.Automation.PSCustomObject] -or $item -is [System.Collections.IDictionary]) {
          $items += ,(ConvertTo-Hashtable -InputObject $item)
        } else {
          $items += ,$item
        }
      }
      $hash[$prop.Name] = $items
    } else {
      $hash[$prop.Name] = $value
    }
  }
  return $hash
}

function Get-ScriptRootSafe {
  if ($PSScriptRoot) { return $PSScriptRoot }
  if ($MyInvocation.MyCommand.Path) { return (Split-Path -Parent $MyInvocation.MyCommand.Path) }
  return (Get-Location).Path
}

function Get-ConfigDirectory {
  $scriptRoot = Get-ScriptRootSafe
  $candidate = Join-Path (Split-Path -Parent $scriptRoot) 'config'
  if (Test-Path $candidate) { return $candidate }
  return (Join-Path $scriptRoot 'config')
}

function Initialize-ProfileConfigMap {
  $configDir = Get-ConfigDirectory
  $Script:ProfileConfigMap = @{
    'Default' = (Join-Path $configDir 'DFO365_Default.json')
    'P1'      = (Join-Path $configDir 'DFO365_P1.json')
    'P2'      = (Join-Path $configDir 'DFO365_P2.json')
  }
}

function Load-ConfigFile {
  param(
    [Parameter(Mandatory)][string]$Path,
    [System.Windows.Forms.Label]$ConfigLabel
  )

  if (-not (Test-Path $Path)) {
    $msg = "Config file not found: $Path"
    Write-UiStatus "[ERR] $msg" 'Red'
    return $false
  }

  try {
    $raw = Get-Content -Path $Path -Raw -Encoding UTF8
    $json = $raw | ConvertFrom-Json -ErrorAction Stop
    $Script:Config = $json
    $Script:LoadedConfigPath = $Path
    $leaf = Split-Path -Leaf $Path
    if ($ConfigLabel) {
      $ConfigLabel.Text = "Config: $leaf"
      $ConfigLabel.ForeColor = [System.Drawing.Color]::DarkBlue
    }
    Write-UiStatus "[OK] Config loaded: $leaf" 'Green'
    return $true
  }
  catch {
    Write-UiStatus "[ERR] Failed to load config: $($_.Exception.Message)" 'Red'
    return $false
  }
}

function Load-ProfileConfig {
  param(
    [Parameter(Mandatory)][string]$ProfileName,
    [System.Windows.Forms.Label]$ConfigLabel
  )

  if (-not $Script:ProfileConfigMap.ContainsKey($ProfileName)) {
    Write-UiStatus "[ERR] Unknown profile: $ProfileName" 'Red'
    return $false
  }
  return (Load-ConfigFile -Path $Script:ProfileConfigMap[$ProfileName] -ConfigLabel $ConfigLabel)
}

function Get-ConfigSection {
  param(
    [Parameter(Mandatory)][string]$SectionName,
    [hashtable]$Fallback = [ordered]@{}
  )

  if ($Script:Config -and $Script:Config.PSObject.Properties.Name -contains $SectionName) {
    return (ConvertTo-Hashtable -InputObject $Script:Config.$SectionName)
  }
  return $Fallback
}

function Get-ConfigValue {
  param(
    [Parameter(Mandatory)][string]$SectionName,
    [Parameter(Mandatory)][string]$Key,
    $DefaultValue = $null
  )

  if ($Script:Config -and $Script:Config.PSObject.Properties.Name -contains $SectionName) {
    $section = $Script:Config.$SectionName
    if ($section -and $section.PSObject.Properties.Name -contains $Key) {
      return $section.$Key
    }
  }
  return $DefaultValue
}

function Get-NamesMap {
  $fallback = [ordered]@{
    SafeLinksPolicy        = 'Microsoft-Zero-Trust-SafeLinks-Rule'
    SafeLinksRule          = 'Microsoft-Zero-Trust-SafeLinks-Rule'
    SafeAttachmentsPolicy  = 'Microsoft-Zero-Trust-SafeAttachments'
    SafeAttachmentsRule    = 'Microsoft-Zero-Trust-SafeAttachments-Rule'
    AntiPhishPolicy        = 'Microsoft-Zero-Trust-AntiPhish'
    AntiPhishRule          = 'Microsoft-Zero-Trust-AntiPhish-Rule'
    AntiSpamInboundPolicy  = 'Microsoft-Zero-Trust-AntiSpam-Inbound'
    AntiSpamInboundRule    = 'Microsoft-Zero-Trust-AntiSpam-Inbound-Rule'
    AntiSpamOutboundPolicy = 'Microsoft-Zero-Trust-AntiSpam-Outbound'
    AntiSpamOutboundRule   = 'Microsoft-Zero-Trust-AntiSpam-Outbound-Rule'
    AntiMalwarePolicy      = 'Microsoft-Zero-Trust-AntiMalware'
    AntiMalwareRule        = 'Microsoft-Zero-Trust-AntiMalware-Rule'
  }
  return (Get-ConfigSection -SectionName 'Names' -Fallback $fallback)
}

# NOTE: Some Exchange Online rules may default to Enabled on creation. Rules are explicitly set to Disabled during deployment.

function Ensure-SafeLinksPolicy {
  param([string]$Name)
  $settings = Get-ConfigSection -SectionName 'SafeLinksPolicy' -Fallback ([ordered]@{
    EnableSafeLinksForEmail    = $true
    EnableSafeLinksForTeams    = $true
    EnableForInternalSenders   = $true
    ScanUrls                   = $true
    DeliverMessageAfterScan    = $true
    DisableUrlRewrite          = $false
    TrackClicks                = $true
    AllowClickThrough          = $false
    EnableOrganizationBranding = $false
  })

  $exists = Get-SafeLinksPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  if (-not $exists) {
    Write-UiStatus "Safe Links policy '$Name' does not exist. Creating it..." "Cyan"
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
    Write-UiStatus "Safe Links rule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-SafeLinksRule' -Identity $RuleName
}

function Ensure-SafeAttachmentsPolicy {
  param([string]$Name)

  function Add-EnableParam([hashtable]$h, [bool]$on=$true, [string]$newCmd, [string]$setCmd) {
    if ($newCmd -and (Supports-Param $newCmd 'Enable'))      { $h['Enable']  = $on }
    elseif ($newCmd -and (Supports-Param $newCmd 'Enabled')) { $h['Enabled'] = $on }
    elseif ($setCmd -and (Supports-Param $setCmd 'Enable'))  { $h['Enable']  = $on }
    elseif ($setCmd -and (Supports-Param $setCmd 'Enabled')) { $h['Enabled'] = $on }
    return $h
  }

  $settings = Get-ConfigSection -SectionName 'SafeAttachmentsPolicy' -Fallback ([ordered]@{
    Action        = 'Block'
    QuarantineTag = 'AdminOnlyAccessPolicy'
    Redirect      = $false
  })

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
    Write-UiStatus "Safe Attachments rule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-SafeAttachmentRule' -Identity $RuleName
}

function Ensure-AntiPhishPolicy {
  param([string]$Name)
  $policy = Get-AntiPhishPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name

  $vals = Get-ConfigSection -SectionName 'AntiPhishPolicy' -Fallback ([ordered]@{
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
  })

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
    Write-UiStatus "Anti-Phish rule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-AntiPhishRule' -Identity $RuleName
}

function Ensure-AntiSpamInboundPolicy {
  param([string]$Name)
  $policy = Get-HostedContentFilterPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  $vals = Get-ConfigSection -SectionName 'AntiSpamInboundPolicy' -Fallback ([ordered]@{
    BulkThreshold                        = 5
    SpamAction                           = 'Quarantine'
    SpamQuarantineTag                    = 'DefaultFullAccesswithNotificationPolicy'
    HighConfidenceSpamAction             = 'Quarantine'
    HighConfidenceSpamQuarantineTag      = 'DefaultFullAccesswithNotificationPolicy'
    BulkSpamAction                       = 'Quarantine'
    BulkQuarantineTag                    = 'DefaultFullAccesswithNotificationPolicy'
    PhishSpamAction                      = 'Quarantine'
    PhishQuarantineTag                   = 'AdminOnlyAccessPolicy'
    HighConfidencePhishAction            = 'Quarantine'
    HighConfidencePhishQuarantineTag     = 'AdminOnlyAccessPolicy'
    InlineSafetyTipsEnabled              = $true
    SpamZapEnabled                       = $true
    PhishZapEnabled                      = $true
    IncreaseScoreWithImageLinks          = 'On'
    IncreaseScoreWithNumericIps          = 'On'
    IncreaseScoreWithRedirectToOtherPort = 'On'
    IncreaseScoreWithBizOrInfoUrls       = 'On'
    MarkAsSpamEmptyMessages              = 'On'
    MarkAsSpamEmbedTagsInHtml            = 'On'
    MarkAsSpamJavaScriptInHtml           = 'On'
    MarkAsSpamFormTagsInHtml             = 'On'
    MarkAsSpamFramesInHtml               = 'On'
    MarkAsSpamWebBugsInHtml              = 'On'
    MarkAsSpamObjectTagsInHtml           = 'On'
    MarkAsSpamSensitiveWordList          = 'Off'
    MarkAsSpamSpfRecordHardFail          = 'On'
    MarkAsSpamFromAddressAuthFail        = 'On'
    MarkAsSpamNdrBackscatter             = 'On'
  })
  if (-not $policy) {
    Write-UiStatus "Inbound Anti-Spam policy '$Name' does not exist. Creating it..." "Cyan"
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
    Write-UiStatus "Inbound Anti-Spam rule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-HostedContentFilterRule' -Identity $RuleName -DisableCmdletName 'Disable-HostedContentFilterRule'
}

function Ensure-AntiSpamOutboundPolicy {
  param([string]$Name,[string]$NotifyAddress)
  $policy = Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  $vals = Get-ConfigSection -SectionName 'AntiSpamOutboundPolicy' -Fallback ([ordered]@{
    RecipientLimitExternalPerHour = 400
    RecipientLimitInternalPerHour = 800
    RecipientLimitPerDay          = 800
    ActionWhenThresholdReached    = 'BlockUser'
    AutoForwardingMode            = 'Off'
    BccSuspiciousOutboundMail     = $false
    NotifyOutboundSpam            = $true
    NotifyOutboundSpamRecipients  = $NotifyAddress
  })
  if (-not $policy) {
    Write-UiStatus "Outbound Anti-Spam policy '$Name' does not exist. Creating it..." "Cyan"
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
    Write-UiStatus "Outbound Anti-Spam rule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-HostedOutboundSpamFilterRule' -Identity $RuleName -DisableCmdletName 'Disable-HostedOutboundSpamFilterRule'
}

function Ensure-AntiMalwarePolicy {
  param([string]$Name,[string]$AdminNotify)
  $policy = Get-MalwareFilterPolicy -ErrorAction SilentlyContinue | Where-Object Name -eq $Name
  $vals = Get-ConfigSection -SectionName 'AntiMalwarePolicy' -Fallback ([ordered]@{
    EnableInternalSenderAdminNotifications = $true
    InternalSenderAdminAddress             = $AdminNotify
    Action                                = 'DeleteMessage'
    EnableZeroHourAutoPurge               = $true
  })
  if (-not $policy) {
    Write-UiStatus "Anti-Malware policy '$Name' does not exist. Creating it..." "Cyan"
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
    Write-UiStatus "Anti-Malware rule '$RuleName' already exists. Keeping it disabled..." "Yellow"
  }
  Disable-RuleOnly -SetCmdletName 'Set-MalwareFilterRule' -Identity $RuleName -DisableCmdletName 'Disable-MalwareFilterRule'
}

function Export-PoliciesJson {
  param([string]$Path)
  $items = @()
  $items += Get-SafeLinksPolicy           -ErrorAction SilentlyContinue
  $items += Get-SafeLinksRule             -ErrorAction SilentlyContinue
  $items += Get-SafeAttachmentPolicy      -ErrorAction SilentlyContinue
  $items += Get-SafeAttachmentRule        -ErrorAction SilentlyContinue
  $items += Get-AntiPhishPolicy           -ErrorAction SilentlyContinue
  $items += Get-AntiPhishRule             -ErrorAction SilentlyContinue
  $items += Get-HostedContentFilterPolicy -ErrorAction SilentlyContinue
  $items += Get-HostedContentFilterRule   -ErrorAction SilentlyContinue
  $items += Get-HostedOutboundSpamFilterPolicy -ErrorAction SilentlyContinue
  $items += Get-HostedOutboundSpamFilterRule   -ErrorAction SilentlyContinue
  $items += Get-MalwareFilterPolicy       -ErrorAction SilentlyContinue
  $items += Get-MalwareFilterRule         -ErrorAction SilentlyContinue

  if (-not (Test-Path $Path)) { New-Item -ItemType Directory -Path $Path | Out-Null }
  $i = 0
  foreach ($obj in $items) {
    $i++
    $name = ($obj.Name | ForEach-Object { $_ }) -join '_' ; if (-not $name) { $name = "item$i" }
    $file = Join-Path $Path ("{0}_{1}.json" -f $obj.GetType().Name, ($name -replace '[^\w\-]','_'))
    $obj | ConvertTo-Json -Depth 10 | Out-File -FilePath $file -Encoding UTF8
  }
}

function Run-Validation {
  param([hashtable]$NamesMap)

  Write-UiStatus "Running validation..." "Cyan"

  $checks = @(
    @{ Name='Anti-Phish';          RuleCmd='Get-AntiPhishRule';             RuleName=$NamesMap.AntiPhishRule },
    @{ Name='Safe Links';          RuleCmd='Get-SafeLinksRule';             RuleName=$NamesMap.SafeLinksRule },
    @{ Name='Safe Attachments';    RuleCmd='Get-SafeAttachmentRule';        RuleName=$NamesMap.SafeAttachmentsRule },
    @{ Name='Inbound Anti-Spam';   RuleCmd='Get-HostedContentFilterRule';   RuleName=$NamesMap.AntiSpamInboundRule },
    @{ Name='Outbound Anti-Spam';  RuleCmd='Get-HostedOutboundSpamFilterRule'; RuleName=$NamesMap.AntiSpamOutboundRule },
    @{ Name='Anti-Malware';        RuleCmd='Get-MalwareFilterRule';         RuleName=$NamesMap.AntiMalwareRule }
  )

  foreach ($c in $checks) {
    try {
      $rule = & $c.RuleCmd -ErrorAction SilentlyContinue | Where-Object Name -eq $c.RuleName | Select-Object -First 1
      if (-not $rule) {
        Write-UiStatus "[FAIL] $($c.Name) rule '$($c.RuleName)' not found." "Red"
        continue
      }

      $disabled = $false
      if ($rule.PSObject.Properties.Name -contains 'Enabled') { $disabled = ($rule.Enabled -eq $false) }
      elseif ($rule.PSObject.Properties.Name -contains 'State') { $disabled = ($rule.State -eq 'Disabled') }

      if ($disabled) {
        Write-UiStatus "[PASS] $($c.Name) rule '$($c.RuleName)' is disabled." "Green"
      } else {
        Write-UiStatus "[FAIL] $($c.Name) rule '$($c.RuleName)' is enabled." "Red"
      }
    }
    catch {
      Write-UiStatus "[ERR] Validation failed for $($c.Name): $($_.Exception.Message)" "Red"
    }
  }
  Write-UiStatus "[OK] Validation complete." "Green"
}

# -------------------- GUI --------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = "DFO365 Deployment Tool - V1.1"
$form.Size = New-Object System.Drawing.Size(760,650)
$form.StartPosition = "CenterScreen"

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$statusBox.Size = New-Object System.Drawing.Size(720,240)
$statusBox.Location = New-Object System.Drawing.Point(10,390)
$form.Controls.Add($statusBox)
function Log($msg){ $statusBox.AppendText("$msg`r`n") }

$lblConnection = New-Object System.Windows.Forms.Label
$lblConnection.Text = "Status: Not Connected"
$lblConnection.Size = New-Object System.Drawing.Size(720,20)
$lblConnection.Location = New-Object System.Drawing.Point(10,325)
$lblConnection.Font = New-Object System.Drawing.Font("Segoe UI",10,[System.Drawing.FontStyle]::Bold)
$form.Controls.Add($lblConnection)

$lblConfig = New-Object System.Windows.Forms.Label
$lblConfig.Text = "Config: Not Loaded"
$lblConfig.Size = New-Object System.Drawing.Size(720,20)
$lblConfig.Location = New-Object System.Drawing.Point(10,350)
$lblConfig.Font = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Regular)
$form.Controls.Add($lblConfig)

$lblProfile = New-Object System.Windows.Forms.Label
$lblProfile.Text = "Profile"
$lblProfile.Size = New-Object System.Drawing.Size(60,20)
$lblProfile.Location = New-Object System.Drawing.Point(10,15)
$form.Controls.Add($lblProfile)

$cmbProfile = New-Object System.Windows.Forms.ComboBox
$cmbProfile.DropDownStyle = 'DropDownList'
$cmbProfile.Size = New-Object System.Drawing.Size(150,25)
$cmbProfile.Location = New-Object System.Drawing.Point(70,12)
[void]$cmbProfile.Items.AddRange(@('Default','P1','P2'))
$cmbProfile.SelectedIndex = 0
$form.Controls.Add($cmbProfile)

$folderDialog = New-Object System.Windows.Forms.FolderBrowserDialog
$openConfigDialog = New-Object System.Windows.Forms.OpenFileDialog
$openConfigDialog.Filter = "JSON Files (*.json)|*.json"

# Row 1
$btnConnect = New-Object System.Windows.Forms.Button
$btnConnect.Text = "Connect (Exchange Online)"
$btnConnect.Size = New-Object System.Drawing.Size(220,40)
$btnConnect.Location = New-Object System.Drawing.Point(230,10)
$btnConnect.Add_Click({
  [void](Ensure-ExchangeOnlineAuthenticated -ForceReauth -ConnectionLabel $lblConnection -Logger ${function:Log})
})
$form.Controls.Add($btnConnect)

$btnDisconnect = New-Object System.Windows.Forms.Button
$btnDisconnect.Text = "Disconnect"
$btnDisconnect.Size = New-Object System.Drawing.Size(220,40)
$btnDisconnect.Location = New-Object System.Drawing.Point(470,10)
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
})
$form.Controls.Add($btnDisconnect)

# Row 2
$btnLoadProfile = New-Object System.Windows.Forms.Button
$btnLoadProfile.Text = "Load Selected Profile"
$btnLoadProfile.Size = New-Object System.Drawing.Size(220,40)
$btnLoadProfile.Location = New-Object System.Drawing.Point(10,60)
$btnLoadProfile.Add_Click({
  [void](Load-ProfileConfig -ProfileName $cmbProfile.SelectedItem -ConfigLabel $lblConfig)
})
$form.Controls.Add($btnLoadProfile)

$btnLoadConfig = New-Object System.Windows.Forms.Button
$btnLoadConfig.Text = "Load Config JSON"
$btnLoadConfig.Size = New-Object System.Drawing.Size(220,40)
$btnLoadConfig.Location = New-Object System.Drawing.Point(250,60)
$btnLoadConfig.Add_Click({
  if ($openConfigDialog.ShowDialog() -eq 'OK') {
    [void](Load-ConfigFile -Path $openConfigDialog.FileName -ConfigLabel $lblConfig)
  }
})
$form.Controls.Add($btnLoadConfig)

$btnValidate = New-Object System.Windows.Forms.Button
$btnValidate.Text = "Run Validation"
$btnValidate.Size = New-Object System.Drawing.Size(220,40)
$btnValidate.Location = New-Object System.Drawing.Point(490,60)
$btnValidate.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    $Names = Get-NamesMap
    Run-Validation -NamesMap $Names
  }
  catch {
    Log "[ERR] Validation error: $($_.Exception.Message)"
  }
})
$form.Controls.Add($btnValidate)

# Row 3
$btnQuickBuild = New-Object System.Windows.Forms.Button
$btnQuickBuild.Text = "Quick Build: All Baselines"
$btnQuickBuild.Size = New-Object System.Drawing.Size(700,40)
$btnQuickBuild.Location = New-Object System.Drawing.Point(10,110)
$btnQuickBuild.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    if (-not $Script:Config) {
      [void](Load-ProfileConfig -ProfileName $cmbProfile.SelectedItem -ConfigLabel $lblConfig)
    }
    $Names = Get-NamesMap
    $AdminNotify = Get-ConfigValue -SectionName 'General' -Key 'AdminNotify' -DefaultValue 'postmaster@yourdomain.com'
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

# Row 4
$btnAPh = New-Object System.Windows.Forms.Button
$btnAPh.Text = "Anti-Phish"
$btnAPh.Size = New-Object System.Drawing.Size(220,40)
$btnAPh.Location = New-Object System.Drawing.Point(10,160)
$btnAPh.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    $Names = Get-NamesMap
    $dom = Get-AllAcceptedDomains
    Ensure-AntiPhishPolicy -Name $Names.AntiPhishPolicy
    Ensure-AntiPhishRuleGlobal -RuleName $Names.AntiPhishRule -PolicyName $Names.AntiPhishPolicy -RecipientDomains $dom
    Log "[OK] Anti-Phish baseline ensured."
  } catch { Log "[ERR] Anti-Phish error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnAPh)

$btnASp = New-Object System.Windows.Forms.Button
$btnASp.Text = "Anti-Spam"
$btnASp.Size = New-Object System.Drawing.Size(220,40)
$btnASp.Location = New-Object System.Drawing.Point(250,160)
$btnASp.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    $Names = Get-NamesMap
    $AdminNotify = Get-ConfigValue -SectionName 'General' -Key 'AdminNotify' -DefaultValue 'postmaster@yourdomain.com'
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
$btnAMw.Text = "Anti-Malware"
$btnAMw.Size = New-Object System.Drawing.Size(220,40)
$btnAMw.Location = New-Object System.Drawing.Point(490,160)
$btnAMw.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    $Names = Get-NamesMap
    $AdminNotify = Get-ConfigValue -SectionName 'General' -Key 'AdminNotify' -DefaultValue 'postmaster@yourdomain.com'
    $dom = Get-AllAcceptedDomains
    Ensure-AntiMalwarePolicy -Name $Names.AntiMalwarePolicy -AdminNotify $AdminNotify
    Ensure-AntiMalwareRuleGlobal -RuleName $Names.AntiMalwareRule -PolicyName $Names.AntiMalwarePolicy -RecipientDomains $dom
    Log "[OK] Anti-Malware baseline ensured."
  } catch { Log "[ERR] Anti-Malware error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnAMw)

# Row 5
$btnSL = New-Object System.Windows.Forms.Button
$btnSL.Text = "Safe Links"
$btnSL.Size = New-Object System.Drawing.Size(220,40)
$btnSL.Location = New-Object System.Drawing.Point(10,210)
$btnSL.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    $Names = Get-NamesMap
    $dom = Get-AllAcceptedDomains
    Ensure-SafeLinksPolicy -Name $Names.SafeLinksPolicy
    Ensure-SafeLinksRuleGlobal -RuleName $Names.SafeLinksRule -PolicyName $Names.SafeLinksPolicy -RecipientDomains $dom
    Log "[OK] Safe Links baseline ensured."
  } catch { Log "[ERR] Safe Links error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnSL)

$btnSA = New-Object System.Windows.Forms.Button
$btnSA.Text = "Safe Attachments"
$btnSA.Size = New-Object System.Drawing.Size(220,40)
$btnSA.Location = New-Object System.Drawing.Point(250,210)
$btnSA.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    $Names = Get-NamesMap
    $dom = Get-AllAcceptedDomains
    Ensure-SafeAttachmentsPolicy -Name $Names.SafeAttachmentsPolicy
    Ensure-SafeAttachmentsRuleGlobal -RuleName $Names.SafeAttachmentsRule -PolicyName $Names.SafeAttachmentsPolicy -RecipientDomains $dom
    Log "[OK] Safe Attachments baseline ensured."
  } catch { Log "[ERR] Safe Attachments error: $($_.Exception.Message)" }
})
$form.Controls.Add($btnSA)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Current Policies (JSON)"
$btnExport.Size = New-Object System.Drawing.Size(220,40)
$btnExport.Location = New-Object System.Drawing.Point(490,210)
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

# Row 6
Add-Type -AssemblyName Microsoft.VisualBasic
$btnSLUrls = New-Object System.Windows.Forms.Button
$btnSLUrls.Text = "Safe Links URL Management"
$btnSLUrls.Size = New-Object System.Drawing.Size(700,40)
$btnSLUrls.Location = New-Object System.Drawing.Point(10,260)
$btnSLUrls.Add_Click({
  try {
    if (-not (Ensure-ExchangeOnlineAuthenticated -ConnectionLabel $lblConnection -Logger ${function:Log})) { return }
    $Names = Get-NamesMap
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
  Initialize-ProfileConfigMap
  $form.Activate()
  Update-ConnectionLabel -Label $lblConnection
  [void](Load-ProfileConfig -ProfileName $cmbProfile.SelectedItem -ConfigLabel $lblConfig)
})
[void]$form.ShowDialog()
