# 🛡️ DFO365 Deployment Tool

A Windows PowerShell deployment tool for configuring a Zero Trust baseline in Microsoft Defender for Office 365, with JSON-driven configuration, policy status indicators, test mode, validation, and built-in HTML/JSON reporting.

## Interface Preview

![DFO365 Deployment Tool UI](images/dfo365-ui.png)

---

[![PowerShell Version](https://img.shields.io/badge/PowerShell-5.1%2B-blue?style=for-the-badge)](YOUR-REPO-URL/blob/main/scripts/DFO365_V1_1_FINAL_enable_services.ps1)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?style=for-the-badge)](YOUR-REPO-URL)
[![Dependency](https://img.shields.io/badge/Dependency-ExchangeOnlineManagement-blueviolet?style=for-the-badge)](YOUR-REPO-URL)
[![Configuration](https://img.shields.io/badge/Config-JSON%20Driven-informational?style=for-the-badge)](YOUR-REPO-URL/blob/main/config/DFO365_ZeroTrust_FINAL.json)
[![UI](https://img.shields.io/badge/UI-WinForms-darkgreen?style=for-the-badge)](YOUR-REPO-URL)
[![Security Model](https://img.shields.io/badge/Security-Zero%20Trust-critical?style=for-the-badge)](YOUR-REPO-URL)
[![Version](https://img.shields.io/badge/Version-V1.1.0-purple?style=for-the-badge)](YOUR-REPO-URL/releases)
[![Status](https://img.shields.io/badge/Status-Stable-success?style=for-the-badge)](YOUR-REPO-URL/releases)
[![License](https://img.shields.io/badge/License-MIT-green?style=for-the-badge)](YOUR-REPO-URL/blob/main/LICENSE)

---

# 📑 Table of Contents

📌 Current Version

🎯 Purpose

🆕 What's New in V1.1

⚠️ Important Behavior

🧱 Project Structure

⚙️ Requirements

🔐 Permissions Required

🚀 Quick Start

🖥️ UI Overview

🧪 Deployment Workflow

🧪 Test Mode

✅ Validation

📤 Export & Reporting

🧠 Design Principles

🔄 Versioning

🚧 Roadmap

💬 Notes

📌 Disclaimer

⭐ Summary

🙌 Contributions

📌 Current Version

V1.1.0 — Stable Release

This version introduces a fully interactive deployment UI, JSON-driven configuration, validation, reporting, and operational controls.

# 🎯 Purpose

The DFO365 Deployment Tool simplifies and standardizes the deployment of:

Anti-Phishing (P2)
Anti-Spam (Inbound & Outbound)
Safe Links
Safe Attachments
Anti-Malware

All aligned to a Zero Trust baseline.

🆕 What's New in V1.1
✅ Full WinForms GUI
✅ JSON-driven configuration
✅ Policy Status Indicators (real-time)
✅ Test Mode (Preview changes before deploy)
✅ Enable Services Toggle (live enforcement control)
✅ HTML Reporting directly from UI
✅ Improved deployment reliability & error handling
⚠️ Important Behavior
Policies are created or updated
Rules are deployed and can be:
Disabled (default safe mode)
Enabled via Enable Services toggle
No impact to mail flow during initial deployment
Safe to run multiple times (idempotent)

#🧱 Project Structure
repo/
├── scripts/
│   └── DFO365_V1_1.ps1
│
├── config/
│   └── DFO365_ZeroTrust.json
│
├── tests/
│   ├── smoke-test-checklist.md
│   └── validation-scenarios.md

# ⚙️ Requirements
PowerShell 5.1 or later
ExchangeOnlineManagement module

Install module:

```Powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
```

# 🔐 Permissions Required

Global Administrator
or
Security Administrator

#🚀 Quick Start
.\scripts\DFO365_V1_1.ps1

Then:

Click Connect to Exchange Online
Load config (auto-loads by default)
Deploy or validate

# 🖥️ UI Overview

The UI includes:

🔌 Connection Panel
🚀 Deploy Zero Trust Baseline
🧩 Individual Policy Deployment Buttons
🧪 Test Mode Preview
🔄 Enable Services Toggle
📊 Policy Status Indicators
📄 HTML / JSON Export
📜 Activity Log

# 🧪 Deployment Workflow
Launch tool
Connect to tenant
(Optional) Run Test Mode
Click Deploy Zero Trust Baseline
(Optional) Enable services via toggle
Validate deployment
Export report

# 🧪 Test Mode

Test Mode allows you to:

Preview changes before deployment
Identify missing policies/rules
Understand impact safely

Output includes:

“Would create”
“Would update”
“Would enable/disable”

# ✅ Validation
🔹 Built-in Validation (UI)
Policy existence
Rule existence
Rule state
🔹 Manual Testing

Location:

tests/smoke-test-checklist.md
🔹 Advanced Scenarios
tests/validation-scenarios.md

Includes:

Phishing simulation
Safe Links testing
Safe Attachments testing
Spam filtering
Malware (EICAR)

# 📤 Export & Reporting
From UI:
Export JSON
Export HTML Report
HTML Report Includes:
Tenant + account info
All policies and rules
Full configuration snapshot
🧠 Design Principles
Idempotent deployment
Safe-by-default
Visual operational feedback
Minimal tenant impact
JSON-driven flexibility
Operator-friendly UI

# 🔄 Versioning
Version	Description
V1.0.0	Baseline deployment tool
V1.1.0	UI, validation, reporting, JSON config, test mode
V2 (Planned)	Multi-profile + advanced configuration engine

# 🚧 Roadmap

V1.2
Risk scoring view
Policy drift detection
Export improvements

V2
Multiple config profiles (Zero Trust, Standard, Audit)
Full config vs tenant comparison
Enhanced UI (tabs / grouping)

# 💬 Notes
Some Exchange rules may default to enabled
Tool explicitly controls final state
Status indicators reflect real-time state

# 📌 Disclaimer

This tool is provided as-is for deployment acceleration and standardization.

Always validate in a test tenant before production use.

# ⭐ Summary

DFO365 Deployment Tool V1.1 delivers:

Secure baseline deployment
Real-time visibility
Safe testing capability
Operational control via UI
Built-in reporting

# 🙌 Contributions

Feedback, improvements, and ideas are welcome.

## 🔥 Tip

For best results:

Keep JSON in /config
Run script from /scripts
Test before enabling services in production
