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
│   └── DefenderforOffice365DeploymentV1.ps1
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

* JSON export/reporting module
* Deploy-all wrapper script
* Config-driven deployments (JSON input)
* Logging and reporting enhancements
* Policy comparison/diff mode

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

## 🛡️ Disclaimer

This tool applies security policies that may impact mail flow.
Always test in a controlled environment before enabling rules in production.

---

## 🤝 Contributing (Optional Future)

Contributions, improvements, and feedback are welcome as the project evolves.
