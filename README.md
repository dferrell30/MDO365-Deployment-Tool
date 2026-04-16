# Defender for Office 365 Deployment Tool

![Version](https://img.shields.io/badge/version-V1-blue)
![PowerShell](https://img.shields.io/badge/powershell-5.1%2B-blue)
![Platform](https://img.shields.io/badge/platform-Microsoft%20365-lightgrey)
![Status](https://img.shields.io/badge/status-stable-success)

---

## 📌 Overview

This PowerShell-based deployment tool automates the configuration of Microsoft Defender for Office 365 using a Zero Trust-aligned security model.

It deploys core threat protection policies while ensuring all enforcement rules are **disabled by default** for safe, controlled rollout.

---

## 📖 Table of Contents

* [Overview](#-overview)
* [Quick Start](#-quick-start)
* [Features](#-features)
* [Prerequisites](#-prerequisites)
* [Usage](#-usage)
* [Export Current Configuration](#-export-current-configuration)
* [Deployment Behavior](#-deployment-behavior)
* [Project Structure](#-project-structure)
* [Validation](#-validation)
* [Roadmap](#-roadmap)
* [Change Log](#-change-log)

---

## ⚡ Quick Start

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Run deployment
.\scripts\DefenderforOffice365DeploymentV1.ps1
```

---

## 🚀 Features

* Anti-Phishing policy deployment (P2-aligned)
* Anti-Spam deployment

  * Inbound protection
  * Outbound protection
* Safe Links protection
* Safe Attachments protection
* Anti-Malware protection

### Security Behavior

* Phishing → Quarantine (**AdminOnlyAccessPolicy**)
* Malware → Block & quarantine
* Spam → Strict filtering with quarantine actions
* URL protection → Real-time Safe Links scanning

---

## 🔧 Prerequisites

* Exchange Online PowerShell Module
* Microsoft Defender for Office 365 licensing
* Required roles:

  * Security Administrator
  * Exchange Administrator (recommended)

---

## ▶️ Usage

```powershell
.\scripts\DefenderforOffice365DeploymentV1.ps1
```

---

## 📊 Export Current Configuration

This project includes a **separate reporting script** that exports the current Defender for Office 365 configuration.

### What it does

* Collects all deployed policies and rules
* Outputs:

  * **JSON** (for automation / comparison)
  * **HTML** (for readable reporting)

### Covered Areas

* Anti-Phish policies and rules
* Anti-Spam (Inbound & Outbound)
* Safe Links
* Safe Attachments
* Anti-Malware
* Accepted domains

---

### Run Export

```powershell
.\scripts\Export-DefenderForOffice365Report.ps1
```

---

### Optional Parameters

```powershell
# Auto-connect if not already connected
.\scripts\Export-DefenderForOffice365Report.ps1 -ConnectIfNeeded

# Specify output folder
.\scripts\Export-DefenderForOffice365Report.ps1 -OutputFolder .\output
```

---

### Output Files

* `DefenderForOffice365-<timestamp>.json`
* `DefenderForOffice365-<timestamp>.html`

---

### Notes

* Export script is **read-only**
* Does not modify tenant configuration
* Safe to run anytime

---

## ⚙️ Deployment Behavior

### Policies

* Created if missing
* Updated if existing

### Rules

* Created if missing
* Updated if existing
* **Forced to Disabled after deployment**

This ensures:

* No immediate mail flow disruption
* Safe staged rollout
* Controlled enablement

---

## 🗂️ Project Structure

```text
.
├── README.md
├── .gitignore
├── docs/
│   ├── change-log.md
│   └── deployment-notes.md
├── scripts/
│   ├── DefenderforOffice365DeploymentV1.ps1
│   └── Export-DefenderForOffice365Report.ps1
├── examples/
│   └── sample-output.json
└── tests/
    └── smoke-test-checklist.md
```

---

## ✅ Validation

After deployment, confirm:

* Anti-Phish rule → **Disabled**
* Inbound spam rule → **Disabled**
* Outbound spam rule → **Disabled**
* Safe Links rule → **Disabled**
* Safe Attachments rule → **Disabled**
* Anti-Malware rule → **Disabled**

---

## 🧭 Roadmap

* Config-driven deployments (JSON input)
* Logging and reporting enhancements
* Policy comparison/diff mode
* CI/CD pipeline integration

---

## 📜 Change Log

See: `docs/change-log.md`

---

## ⚠️ Notes

* Required quarantine policies:

  * `AdminOnlyAccessPolicy`
  * `DefaultFullAccesswithNotificationPolicy`
* Some settings may require tenant-specific tuning
* Script uses compatibility checks for different EXO versions

---

## ⚠️ Disclaimer

This tool is provided for **educational, testing, and security validation purposes only**.

Use of this tool should be limited to:
- Authorized environments  
- Lab or approved enterprise systems  

The author assumes **no liability or responsibility** for:
- Misuse of this tool  
- Damage to systems  
- Unauthorized or improper use  

By using this tool, you agree to use it in a lawful and responsible manner.
---

This project is not affiliated with or endorsed by Microsoft.
---


## ⚖️ Professional Disclaimer

This project is an independent work developed in a personal capacity.

The views, opinions, code, and content expressed in this repository are solely my own and do not reflect the views, policies, or positions of any current or future employer, client, or affiliated organization.

No employer, past, present, or future, has reviewed, approved, endorsed, or is in any way associated with these works.

This project was developed outside the scope of any employment and without the use of proprietary, confidential, or restricted resources.

All code/language in this repository is provided under the terms of the included MIT License.

---

## 🤝 Contributing (Optional Future)

Contributions, improvements, and feedback are welcome as the project evolves.
