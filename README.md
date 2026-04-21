# 🛡️ DFO365 Deployment Tool

![PowerShell](https://img.shields.io/badge/PowerShell-5.1+-blue)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey)
![License](https://img.shields.io/badge/License-MIT-green)
![Version](https://img.shields.io/badge/Version-V1.0.0-purple)
![Status](https://img.shields.io/badge/Status-Stable-success)

---

## 📑 Table of Contents

- [📌 Current Version](#-current-version)
- [🎯 Purpose](#-purpose)
- [⚠️ Important Behavior](#️-important-behavior)
- [🧱 Project Structure](#-project-structure)
- [⚙️ Requirements](#️-requirements)
- [🔐 Permissions Required](#-permissions-required)
- [🚀 Quick Start](#-quick-start)
- [🧪 Deployment Workflow](#-deployment-workflow)
- [✅ Validation](#-validation)
- [📤 Export & Reporting](#-export--reporting)
- [🧠 Design Principles](#-design-principles)
- [🔄 Versioning](#-versioning)
- [🚧 Roadmap](#-roadmap)
- [💬 Notes](#-notes)
- [📌 Disclaimer](#-disclaimer)
- [⭐ Summary](#-summary)
- [🙌 Contributions](#-contributions)

---

## 📌 Current Version

**V1.0.0 — Baseline Release**

This tool provides a **repeatable, safe deployment** of Microsoft Defender for Office 365 security policies aligned to a Zero Trust approach.

---

## 🎯 Purpose

The DFO365 Deployment Tool simplifies and standardizes the deployment of:

- Anti-Phishing (P2)
- Anti-Spam (Inbound & Outbound)
- Safe Links
- Safe Attachments
- Anti-Malware

---

## ⚠️ Important Behavior

- Policies are created or updated  
- Rules are **always deployed disabled**  
- No impact to mail flow during deployment  
- Safe to run multiple times  

---

## 🧱 Project Structure


repo/
├── scripts/
│ ├── DFO365_V1.ps1
│ ├── Export-DefenderForOffice365Report.ps1
│ └── Test-DFO365DeploymentValidation.ps1
│
├── tests/
│ ├── smoke-test-checklist.md
│ └── validation-scenarios.md


---

## ⚙️ Requirements

- PowerShell 5.1 or later  
- ExchangeOnlineManagement module  

Install module:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
```
---

# 🔐 Permissions Required
Global Administrator or
Security Administrator

# 🚀 Quick Start

```Powershell
Connect-ExchangeOnline
.\scripts\DFO365_V1.ps1
```

# 🧪 Deployment Workflow

Launch tool
Connect to tenant
Run Quick Build: All Baselines
Validate deployment
Export configuration (optional)


# ✅ Validation
🔹 Smoke Test (Manual)

Location:

tests/smoke-test-checklist.md

🔹 Automated Validation

.\scripts\Test-DFO365DeploymentValidation.ps1

Validates:

Policy existence
Rule existence
Rule disabled state
🔹 Validation Scenarios (Real Testing)

Location:

tests/validation-scenarios.md

Includes:

Phishing simulation
Safe Links behavior
Safe Attachments behavior
Spam filtering
Malware testing (EICAR)

# 📤 Export & Reporting
.\scripts\Export-DefenderForOffice365Report.ps1

Outputs:

JSON configuration
Optional HTML report

# 🧠 Design Principles
Idempotent deployment
Safe-by-default (no rules enabled)
Clear UI feedback
Minimal tenant impact
Repeatable baseline configuration

# 🔄 Versioning
Version	Description
V1.0.0	Baseline deployment tool
V1.1	Validation + reporting improvements (planned)
V2	JSON-driven configuration engine (planned)

# 🚧 Roadmap
V1.1
Enhanced validation script
Improved reporting
Logging to file
V2
JSON-driven configuration
Multiple deployment profiles
GUI config loading
Config vs tenant comparison
💬 Notes

Some Exchange Online rules may default to enabled on creation.
Rules are explicitly set to disabled during deployment.

# 📌 Disclaimer

This tool is provided as-is for deployment acceleration and standardization.
Always validate in a test tenant before production use.

# ⭐ Summary

DFO365 Deployment Tool V1 delivers:

Reliable Defender for Office 365 baseline deployment
Safe, repeatable execution
Clear validation and testing approach

# 🙌 Contributions

Feedback and improvements are welcome.


---

## 🔥 Tip

When you paste this into GitHub:
- Make sure the file is named `README.md`
- View it in **preview mode** to confirm badges render

---

If you want next:
👉 I can add **clickable badges (linking to releases/downloads)**  
👉 or add a **GIF demo section (very strong for GitHub visibility)**
