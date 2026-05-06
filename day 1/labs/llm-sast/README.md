# LLM-Powered SAST Scanner

Offline source code vulnerability scanner using LLM (Ollama).

## Prerequisites

```bash
# 1. Install Ollama (skip if already installed)
curl -fsSL https://ollama.com/install.sh | sh
# macOS/Windows: download from https://ollama.com

# 2. Pull the code analysis model (~4.7GB)
ollama pull qwen2.5-coder:7b

# 3. Install Python dependency
pip install ollama
```

## Usage

```bash
# Scan a single file
python llm_sast_scanner.py ../target-app/vulnerable_app.php

# Scan with a different model
python llm_sast_scanner.py ../target-app/vulnerable_app.php qwen2.5:7b-instruct

# Scan multiple files
find ../target-app -name "*.php" | xargs -I{} python llm_sast_scanner.py {}
```

## Files

| File | Description |
|------|-------------|
| `llm_sast_scanner.py` | The scanner script (uses Ollama API) |
| `prompt.txt` | System prompt for security analysis |

## Example Output

```
============================================================
Scanning: vulnerable_app.php
Model: qwen2.5-coder:7b
============================================================

Found 6 vulnerabilities:

[High] SQL Injection (CWE-89) @ line 10
  Explanation: User inputs directly concatenated into SQL query
  Fix: Use prepared statements with bound parameters

[High] XSS (CWE-79) @ line 25
  Explanation: User input echoed without sanitization
  Fix: Use htmlspecialchars()

[High] Unrestricted File Upload (CWE-400) @ line 39
[High] IDOR / Missing Auth (CWE-200) @ line 53
[High] OS Command Injection (CWE-78) @ line 63
[Medium] Weak Token / MD5 (CWE-781) @ line 68
```
