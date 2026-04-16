# Defender for Office 365 Deployment Tool

## Overview

This PowerShell-based deployment tool automates the configuration of Microsoft Defender for Office 365 security policies using a Zero Trust-aligned approach.

The script deploys core threat protection policies and ensures they are configured but **disabled by default** for safe, staged rollout.

---

## Current Version

**V1 (Baseline)**
Script: `DefenderforOffice365DeploymentV1.ps1`

---

## Features

* Anti-Phishing policy deployment (P2-aligned)
* Anti-Spam (Inbound & Outbound) policy deployment
* Safe Links policy deployment
* Safe Attachments policy deployment
* Anti-Malware policy deployment
* Policies configured with:

  * Quarantine enforcement
  * Admin-only access where applicable
* **Rules deployed in disabled state by default**
* Designed for repeatable, idempotent execution

---

## Prerequisites

* Exchange Online PowerShell Module
* Microsoft Defender for Office 365 licensing
* Required roles:

  * Security Administrator
  * Exchange Administrator (recommended)

---

## Setup

### Connect to Exchange Online

```powershell
Connect-ExchangeOnline
```

---

## Usage

### Run Full Deployment

```powershell
.\scripts\DefenderforOffice365DeploymentV1.ps1
```

---

## Deployment Behavior

* Policies are **created or updated**
* Rules are:

  * Created if missing
  * Updated if existing
  * **Forced to Disabled state after execution**

This allows:

* Safe validation
* Controlled enablement
* Reduced risk of mail disruption

---

## Project Structure

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

## Notes

* Quarantine policies must exist in the tenant:

  * `AdminOnlyAccessPolicy`
  * `DefaultFullAccesswithNotificationPolicy`
* Some settings (users/domains) may require customization per environment
* Script uses compatibility checks to support multiple EXO module versions

---

## Roadmap

* JSON export/reporting module
* Deploy-all wrapper script
* Config-driven deployment (JSON input)
* Logging & reporting enhancements

---

## Change Log

See `docs/change-log.md`
