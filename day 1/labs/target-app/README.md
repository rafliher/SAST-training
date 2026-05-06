# Vulnerable Test Application

Intentionally vulnerable PHP application for SAST training demos.

**DO NOT deploy this application on any public-facing server.**

## Vulnerabilities

| # | Type | Function | Line | OWASP 2025 |
|---|------|----------|------|------------|
| 1 | SQL Injection | `login()` | 13 | A05 Injection |
| 2 | XSS (Stored) | `show_profile()` | 33-35 | A05 Injection |
| 3 | Unrestricted File Upload | `upload_avatar()` | 41 | A02 Security Misconfiguration |
| 4 | IDOR / Missing Authorization | `delete_user()` | 48 | A01 Broken Access Control |
| 5 | OS Command Injection | `ping_host()` | 54 | A05 Injection |
| 6 | Weak Crypto (MD5 token) | `reset_password()` | 60 | A04 Cryptographic Failures |

## Usage in Training

```bash
# Scan with LLM-SAST
cd ../llm-sast
python llm_sast_scanner.py ../target-app/vulnerable_app.php
```
