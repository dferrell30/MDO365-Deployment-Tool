# 🔐 Security Policy

## 📌 Supported Versions

The following versions of the DFO365 Deployment Tool are currently supported:

| Version | Supported |
|--------|-----------|
| V1.1.x | ✅ Yes |
| V1.0.x | ⚠️ No |

---

## 🚨 Reporting a Vulnerability

If you discover a security issue, please **do not open a public GitHub issue**.

Instead, report it responsibly:

- 📧 Email: YOUR-EMAIL-HERE
- 📨 Or open a private security advisory (if enabled)

Please include:
- Description of the issue  
- Steps to reproduce  
- Potential impact  
- Suggested fix (if available)  

---

## 🔍 Scope

This repository includes:

- PowerShell scripts for Defender for Office 365 deployment  
- JSON configuration files  
- Validation and reporting logic  

Security considerations include:
- Authentication to Exchange Online  
- Policy configuration changes  
- Handling of tenant-specific data  

---

## ⚠️ Security Considerations

- This tool executes **administrative actions** in Microsoft 365  
- It requires **privileged roles** (Global Admin / Security Admin)  
- Always test in a **non-production tenant first**  
- Review JSON configurations before applying  

---

## 🛡️ Best Practices

- Use **least privilege access** where possible  
- Avoid storing sensitive data in configuration files  
- Monitor audit logs after deployment  
- Validate configurations regularly  

---

## 📌 Disclaimer

This tool is provided as-is without warranty.

Use at your own risk. Always validate changes in a controlled environment before deploying to production.

---

## 🙌 Acknowledgements

Thanks to the community and security practitioners who help improve tooling and practices around Microsoft Defender and Zero Trust.
