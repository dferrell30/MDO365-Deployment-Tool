# Defender for Office 365 Deployment Tool

## Overview
PowerShell deployment tool for Microsoft Defender for Office 365 policy configuration.

## Current Version
V1

## Features
- Anti-Phish policy deployment
- Anti-Spam inbound/outbound deployment
- Safe Links deployment
- Safe Attachments deployment
- Anti-Malware deployment
- Export configuration to JSON
- Rules deploy disabled by default

## Prerequisites
- Exchange Online PowerShell
- Security & Compliance / Defender permissions
- Required admin roles

## Files
- `scripts/DefenderforOffice365DeploymentV1.ps1` - main deployment script
- `scripts/deploy-all.ps1` - full deployment runner
- `scripts/export-config-json.ps1` - export current config to JSON

## Usage

### Deploy all
```powershell
.\scripts\deploy-all.ps1
