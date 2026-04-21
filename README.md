# 🛡️ DFO365 Deployment Tool

## 🚀 Defender for Office 365 Baseline Deployment (V1)

---

## 📌 Current Version

**V1.0.0 — Baseline Release**

This tool provides a **repeatable, safe deployment** of Microsoft Defender for Office 365 security policies aligned to a Zero Trust approach.

All policies are deployed with:

* ✅ Correct security settings applied
* ✅ Rules created but **disabled by default**
* ✅ Safe re-run behavior (idempotent deployment)

---

## 🎯 Purpose

The DFO365 Deployment Tool simplifies and standardizes the deployment of:

* Anti-Phishing (P2)
* Anti-Spam (Inbound & Outbound)
* Safe Links
* Safe Attachments
* Anti-Malware

---

## ⚠️ Important Behavior

* Policies are created or updated
* Rules are **always deployed disabled**
* No impact to mail flow during deployment
* Safe to run multiple times

---

## 🧱 Project Structure

```text
repo/
├── scripts/
│   ├── DFO365_V1.ps1
│   ├── Export-DefenderForOffice365Report.ps1
│   └── Test-DFO365DeploymentValidation.ps1
│
├── tests/
│   ├── smoke-test-checklist.md
│   └── validation-scenarios.md
```

---

## ⚙️ Requirements

* PowerShell 5.1 or later
* ExchangeOnlineManagement module

Install module:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
```

---

## 🔐 Permissions Required

* Global Administrator **or**
* Security Administrator

---

## 🚀 Quick Start

```powershell
# Connect to Exchange Online
Connect-ExchangeOnline

# Run the tool
.\scripts\DFO365_V1.ps1
```

---

## 🧪 Deployment Workflow

1. Launch tool
2. Connect to tenant
3. Run **Quick Build: All Baselines**
4. Validate deployment
5. Export configuration (optional)

---

## ✅ Validation

### 🔹 Smoke Test (Manual)

Located in:

```text
tests/smoke-test-checklist.md
```

Provides a quick verification checklist after deployment.

---

### 🔹 Automated Validation

```powershell
.\scripts\Test-DFO365DeploymentValidation.ps1
```

Validates:

* Policy existence
* Rule existence
* Rule disabled state

---

### 🔹 Validation Scenarios (Real Testing)

Located in:

```text
tests/validation-scenarios.md
```

Includes:

* Phishing simulation
* Safe Links behavior
* Safe Attachments behavior
* Spam filtering
* Malware testing (EICAR)

---

## 📤 Export & Reporting

The tool includes export capabilities:

```powershell
.\scripts\Export-DefenderForOffice365Report.ps1
```

Outputs:

* JSON configuration
* Optional HTML report

---

## 🧠 Design Principles

* Idempotent deployment
* Safe-by-default (no rules enabled)
* Clear UI feedback
* Minimal tenant impact
* Repeatable baseline configuration

---

## 🔄 Versioning

| Version | Description                                   |
| ------- | --------------------------------------------- |
| V1.0.0  | Baseline deployment tool                      |
| V1.1    | Validation + reporting improvements (planned) |
| V2      | JSON-driven configuration engine (planned)    |

---

## 🚧 Roadmap

### 🔹 V1.1

* Enhanced validation script
* Improved reporting
* Logging to file

### 🔹 V2

* JSON-driven configuration
* Multiple deployment profiles
* GUI config loading
* Config vs tenant comparison

---

## 💬 Notes

> Some Exchange Online rules may default to enabled on creation.
> Rules are explicitly set to disabled during deployment.

---

## 📌 Disclaimer

This tool is provided as-is for deployment acceleration and standardization.
Always validate in a test tenant before production use.

---

## ⭐ Summary

DFO365 Deployment Tool V1 delivers:

* Reliable Defender for Office 365 baseline deployment
* Safe, repeatable execution
* Clear validation and testing approach

---

## 🙌 Contributions

Feedback and improvements are welcome.
