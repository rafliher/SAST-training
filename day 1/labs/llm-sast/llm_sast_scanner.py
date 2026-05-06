#!/usr/bin/env python3
"""
LLM-powered SAST scanner using Ollama (offline) or Claude API (cloud).
"""

import ollama
import json
import sys
import pathlib


SYSTEM_PROMPT = pathlib.Path(
    pathlib.Path(__file__).parent / "prompt.txt"
).read_text()


def scan_file(filepath, model='qwen2.5-coder:7b'):
    """Scan a source code file for vulnerabilities using LLM."""
    code = pathlib.Path(filepath).read_text()
    filename = pathlib.Path(filepath).name

    print(f"\n{'='*60}")
    print(f"Scanning: {filename}")
    print(f"Model: {model}")
    print(f"{'='*60}\n")

    response = ollama.chat(
        model=model,
        messages=[
            {'role': 'system', 'content': SYSTEM_PROMPT},
            {'role': 'user', 'content': f'Analyze this code:\n```\n{code}\n```'}
        ],
        format='json'
    )

    raw = response['message']['content']

    try:
        results = json.loads(raw)
    except json.JSONDecodeError:
        print(f"[!] Failed to parse JSON response. Raw output:")
        print(raw[:500])
        return None

    vulns = results.get('vulnerabilities', [])
    print(f"Found {len(vulns)} vulnerabilities:\n")

    for v in vulns:
        sev = v.get('severity', '?')
        typ = v.get('type', '?')
        cwe = v.get('cwe', '?')
        line = v.get('line', '?')
        explanation = v.get('explanation', '')
        fix = v.get('fix', '')

        # Color coding for terminal
        colors = {
            'Critical': '\033[91m',  # Red
            'High': '\033[93m',      # Yellow
            'Medium': '\033[96m',    # Cyan
            'Low': '\033[92m',       # Green
        }
        reset = '\033[0m'
        color = colors.get(sev, '')

        print(f"{color}[{sev}]{reset} {typ} (CWE-{cwe}) @ line {line}")
        print(f"  Explanation: {explanation}")
        if fix:
            print(f"  Fix: {fix[:200]}")
        print()

    return results


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python llm_sast_scanner.py <source_file> [model]")
        print("  model defaults to 'qwen2.5-coder:7b'")
        sys.exit(1)

    filepath = sys.argv[1]
    model = sys.argv[2] if len(sys.argv) > 2 else 'qwen2.5-coder:7b'
    # Fallback: if coder model not available, use instruct
    try:
        models = [m.model for m in ollama.list().models]
        if model not in models and 'qwen2.5:7b-instruct' in models:
            print(f"[*] {model} not found, using qwen2.5:7b-instruct")
            model = 'qwen2.5:7b-instruct'
    except Exception:
        pass

    if not pathlib.Path(filepath).exists():
        print(f"Error: {filepath} not found")
        sys.exit(1)

    scan_file(filepath, model)
