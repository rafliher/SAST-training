#!/usr/bin/env python3
"""Add 20 new questions to the Google Form pretest."""

import json
import subprocess
import sys

FORM_ID = "18H-LJ0qJ2AT9mP1Cxj3djfI0-Co1wkfPi2fgkBwocec"
PROJECT = "credible-skill-479602-v1"

# 20 new questions — topics not covered by existing 30
QUESTIONS = [
    {
        "title": "What does the CIA Triad stand for in information security?",
        "options": [
            "Confidentiality, Integrity, Availability",
            "Control, Identity, Authentication",
            "Compliance, Isolation, Authorization",
            "Certification, Infrastructure, Access"
        ],
        "answer": 0
    },
    {
        "title": "In the STRIDE threat model, what does the 'T' stand for?",
        "options": [
            "Theft",
            "Tampering",
            "Token Abuse",
            "Traversal"
        ],
        "answer": 1
    },
    {
        "title": "What is the core idea behind 'Shift-Left Security'?",
        "options": [
            "Moving security testing to the earliest stages of the SDLC",
            "Shifting firewall rules to the left side of the network",
            "Prioritizing left-to-right code review",
            "Moving production security controls to staging"
        ],
        "answer": 0
    },
    {
        "title": "Which type of Application Security Testing analyzes source code without executing it?",
        "options": [
            "DAST (Dynamic Application Security Testing)",
            "IAST (Interactive Application Security Testing)",
            "SAST (Static Application Security Testing)",
            "RASP (Runtime Application Self-Protection)"
        ],
        "answer": 2
    },
    {
        "title": "What is the Principle of Least Privilege (PoLP)?",
        "options": [
            "Users should have admin access during onboarding",
            "Users should only have the minimum access needed to do their job",
            "Privileges should be granted based on seniority",
            "All users share the same access level for simplicity"
        ],
        "answer": 1
    },
    {
        "title": "Which attack tricks a user's browser into making an unwanted request to a site where they are authenticated?",
        "options": [
            "XSS",
            "SQL Injection",
            "CSRF (Cross-Site Request Forgery)",
            "SSRF"
        ],
        "answer": 2
    },
    {
        "title": "A file upload feature allows users to upload .php files that are then served directly. What type of vulnerability is this?",
        "options": [
            "Path Traversal",
            "Unrestricted File Upload leading to Remote Code Execution",
            "Insecure Deserialization",
            "SSRF"
        ],
        "answer": 1
    },
    {
        "title": "In the OWASP Top 10 2025, which is a NEW category that did not exist in 2021?",
        "options": [
            "Broken Access Control",
            "Injection",
            "Software Supply Chain Failures",
            "Cryptographic Failures"
        ],
        "answer": 2
    },
    {
        "title": "What is the main advantage of using an LLM for source code analysis compared to traditional SAST?",
        "options": [
            "LLMs are always faster than rule-based scanners",
            "LLMs can semantically understand code logic and detect business logic flaws",
            "LLMs never produce false positives",
            "LLMs don't need access to source code"
        ],
        "answer": 1
    },
    {
        "title": "Which tool allows you to run LLM models locally and offline for private code analysis?",
        "options": [
            "SonarQube",
            "Burp Suite",
            "Ollama",
            "Nmap"
        ],
        "answer": 2
    },
    {
        "title": "What is a key limitation of using LLMs for SAST?",
        "options": [
            "They cannot read source code",
            "They may hallucinate (report non-existent vulnerabilities)",
            "They only work with Java",
            "They require physical hardware tokens"
        ],
        "answer": 1
    },
    {
        "title": "What is the BEST practice when combining traditional SAST and LLM-powered analysis?",
        "options": [
            "Only use LLM, traditional tools are obsolete",
            "Only use traditional SAST, LLMs are unreliable",
            "Use traditional SAST for speed and LLM for depth — defense in depth",
            "Alternate between them on each release"
        ],
        "answer": 2
    },
    {
        "title": "Which PHP function is the primary defense against Cross-Site Scripting (XSS) when outputting user data?",
        "options": [
            "mysql_real_escape_string()",
            "htmlspecialchars()",
            "base64_encode()",
            "urlencode()"
        ],
        "answer": 1
    },
    {
        "title": "Which PHP configuration directive should be set to OFF to prevent Remote File Inclusion (RFI)?",
        "options": [
            "display_errors",
            "allow_url_include",
            "max_execution_time",
            "session.auto_start"
        ],
        "answer": 1
    },
    {
        "title": "What does the 'SameSite' cookie attribute protect against?",
        "options": [
            "SQL Injection",
            "XSS",
            "CSRF (Cross-Site Request Forgery)",
            "Path Traversal"
        ],
        "answer": 2
    },
    {
        "title": "In DAST testing, the application is tested:",
        "options": [
            "By reading source code directly",
            "By simulating attacks against a running application",
            "By analyzing compiled bytecode only",
            "By reviewing design documents"
        ],
        "answer": 1
    },
    {
        "title": "Which new OWASP Top 10 2025 item replaces the old 'Server-Side Request Forgery (SSRF)' at position A10?",
        "options": [
            "Broken Access Control",
            "Software Supply Chain Failures",
            "Mishandling of Exceptional Conditions",
            "Authentication Failures"
        ],
        "answer": 2
    },
    {
        "title": "What is the purpose of a session.cookie_httponly = 1 setting in PHP?",
        "options": [
            "Encrypts the cookie value",
            "Prevents JavaScript from accessing the session cookie (mitigates XSS)",
            "Makes the cookie expire faster",
            "Allows cookies only over HTTP, not HTTPS"
        ],
        "answer": 1
    },
    {
        "title": "Which of the following is an example of 'Security Misconfiguration' according to OWASP 2025?",
        "options": [
            "Using parameterized queries",
            "Enabling directory listing on a production web server",
            "Hashing passwords with Argon2",
            "Implementing rate limiting on login"
        ],
        "answer": 1
    },
    {
        "title": "When using an LLM for SAST in a CI/CD pipeline, what is the recommended trigger?",
        "options": [
            "Only on the annual security audit",
            "On every pull request, scanning changed files",
            "Only after a security incident",
            "Once per year before deployment"
        ],
        "answer": 1
    },
]

def get_token():
    result = subprocess.run(['gcloud', 'auth', 'print-access-token'],
                          capture_output=True, text=True)
    return result.stdout.strip()

def add_questions():
    token = get_token()
    headers = [
        '-H', f'Authorization: Bearer {token}',
        '-H', f'x-goog-user-project: {PROJECT}',
        '-H', 'Content-Type: application/json',
    ]

    # Build batchUpdate request — add all 20 questions
    requests = []
    for i, q in enumerate(QUESTIONS):
        options = []
        for j, opt in enumerate(q['options']):
            option = {"value": opt}
            if j == q['answer']:
                option["isCorrect"] = True
            options.append(option)

        req = {
            "createItem": {
                "item": {
                    "title": q['title'],
                    "questionItem": {
                        "question": {
                            "required": True,
                            "choiceQuestion": {
                                "type": "RADIO",
                                "options": options,
                                "shuffle": True
                            },
                            "grading": {
                                "pointValue": 1,
                                "correctAnswers": {
                                    "answers": [{"value": q['options'][q['answer']]}]
                                }
                            }
                        }
                    }
                },
                "location": {
                    "index": 5 + 30 + i  # After 5 info fields + 30 existing questions
                }
            }
        }
        requests.append(req)

    payload = json.dumps({"requests": requests})

    url = f"https://forms.googleapis.com/v1/forms/{FORM_ID}:batchUpdate"

    result = subprocess.run(
        ['curl', '-s', '-X', 'POST', url] + headers + ['-d', payload],
        capture_output=True, text=True, timeout=30
    )

    response = json.loads(result.stdout)
    if 'error' in response:
        print(f"ERROR: {json.dumps(response['error'], indent=2)}")
        return False

    replies = response.get('replies', [])
    print(f"Successfully added {len(replies)} questions!")
    return True

if __name__ == '__main__':
    ok = add_questions()
    if not ok:
        sys.exit(1)
