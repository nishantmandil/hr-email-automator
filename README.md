# 📬 HR Email Automator

> Automated cold outreach tool for job seekers — verifies email addresses before sending, attaches resume, and logs every result to a persistent Excel report.

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python&logoColor=white)
![Gmail](https://img.shields.io/badge/Gmail-SMTP-red?style=flat-square&logo=gmail&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Status](https://img.shields.io/badge/Status-Active-brightgreen?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=flat-square)

---

## 🚀 What Problem Does This Solve?

Most job seekers send cold emails blindly from a spreadsheet — many bounce, waste time, damage sender reputation, and flood your inbox with **"Address not found"** failure notices.

**HR Email Automator** fixes this by validating every address before touching Gmail:

| Without This Tool | With This Tool |
|---|---|
| Emails sent to dead domains | ❌ Dead domains caught via DNS — skipped instantly |
| "Address not found" bounces | ❌ Invalid formats rejected before sending |
| No record of what was sent | ✅ Every action logged to Excel with timestamp |
| Script crash = lost progress | ✅ Each row saved to disk immediately |
| Spam-flagged for sending too fast | ✅ 45s rate limit between sends |

---

## ✨ Features

- **3-layer email verification** — format check → fake address check → DNS/MX record check, all without needing port 25
- **Persistent Excel report** — appends to existing report across multiple runs, never overwrites history
- **Crash-safe logging** — every result is written to disk instantly, so stopping mid-run loses nothing
- **Timestamped rows** — know exactly when each email was sent
- **Auto-reconnect** — reconnects to Gmail automatically if SMTP drops mid-run
- **Rate limiting** — 45-second delay between sends to avoid Gmail spam detection
- **Resume attachment** — PDF attached automatically to every outgoing email
- **Two-file architecture** — email content lives in `email_template.py`, logic stays untouched in `main.py`

---

## 📁 Project Structure

```
hr-email-automator/
│
├── main.py                  # Core automation + verification logic
├── email_template.py        # Email subject + body (only file you edit for content)
├── all_hr_emails.xlsx       # Input: HR email list with Company Name + HR Email columns
├── Nishant_SRE_Resume.pdf   # Resume attached to every email
└── email_report.xlsx        # Auto-generated output report (appends across runs)
```

---

## ⚙️ Setup

### 1. Clone the repo

```bash
git clone https://github.com/nishantmandil/hr-email-automator.git
cd hr-email-automator
```

### 2. Install dependencies

```bash
pip install dnspython pandas openpyxl
```

### 3. Prepare your Excel file

Your input file must have exactly these two columns:

| Company Name | HR Email |
|---|---|
| Google | hr@google.com |
| Razorpay | talent@razorpay.com |
| | careers@infosys.com |

> Blank company names are handled gracefully — email still sends with a generic greeting.

### 4. Add your resume

Place your resume PDF in the root folder and update `main.py`:

```python
resume_path = "Your_Name_Resume.pdf"
```

### 5. Set your Gmail credentials

In `main.py`:

```python
EMAIL = "your_email@gmail.com"
PASSWORD = "your_app_password"
```

> ⚠️ Use a [Gmail App Password](https://support.google.com/accounts/answer/185833), **not** your real Gmail password. Requires 2FA enabled on your account.

---

## ▶️ Run

```bash
python main.py
```

---

## 🔍 How Verification Works

Before any email is sent, three checks run in sequence:

```
Email Address
     │
     ▼
┌─────────────────────────┐
│  1. Format Check        │  regex — catches malformed addresses like abc@ or @gmail
│     (regex)             │
└────────────┬────────────┘
             │ pass
             ▼
┌─────────────────────────┐
│  2. Fake Address Check  │  blocks noreply@, test@, admin@, info@, etc.
│     (blocklist)         │
└────────────┬────────────┘
             │ pass
             ▼
┌─────────────────────────┐
│  3. MX Record Check     │  DNS lookup — confirms domain can receive email
│     (DNS only)          │  works without port 25, never blocked by ISP
└────────────┬────────────┘
             │ pass
             ▼
        ✅ Send Email
```

**Why DNS-only and not SMTP port 25?**
Port 25 is blocked by most ISPs in India and corporate networks. Earlier versions of this tool used SMTP handshake verification — every check timed out and fell back to "assume valid", causing bounce emails. The current approach uses DNS MX record lookups (port 53) which are never blocked and reliably confirm whether a domain can receive email at all.

---

## 📊 Output Report

`email_report.xlsx` is created on first run and **appended on every subsequent run** — it grows into a complete outreach history:

| Company Name | HR Email | Status | Error | Sent At |
|---|---|---|---|---|
| Google | hr@google.com | Sent | | 2025-03-14 10:22:01 |
| Razorpay | bad@nodomain.xyz | Skipped | Domain has no MX record | 2025-03-14 10:22:03 |
| Infosys | noreply@infosys.com | Skipped | Generic/placeholder email address | 2025-03-14 10:22:04 |
| EY | invalid-email | Skipped | Invalid email format | 2025-03-14 10:22:04 |
| Colt | talent@colt.net | Sent | | 2025-03-15 09:10:44 |

> Script stopped mid-run? Every row processed before the stop is already saved. Resume from where you left off next run.

---

## ✏️ Customizing Your Email

All content lives in `email_template.py` — change subject or body without touching any logic:

```python
def get_email_subject():
    return "Your Custom Subject Line Here"

def get_email_body(company=None):
    body = """Hi Hiring Team,

Your email content here...
"""
    return body
```

`main.py` never needs to change for content updates.

---

## 🛡️ Responsible Usage

- Built for **personal job outreach only**
- Respects Gmail's sending limits (~500 emails/day for regular accounts)
- 45-second delay between sends prevents spam flagging
- Do not use for bulk unsolicited marketing or spam

---

## 🧰 Tech Stack

| Library | Purpose |
|---|---|
| `smtplib` | Gmail SMTP login and sending |
| `dns.resolver` | MX record lookups for domain verification |
| `re` | Email format validation via regex |
| `pandas` | Reading Excel input, writing report |
| `openpyxl` | Excel file engine for pandas |
| `email.message` | Building MIME emails with PDF attachment |
| `os` + `datetime` | Report file management + timestamps |

---

## 👤 Author

**Nishant Mandil**
Site Reliability Engineer @ Colt Technology Services

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/nishant-mandil-07b165159/)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-black?style=flat-square&logo=github)](https://github.com/nishantmandil)
[![Portfolio](https://img.shields.io/badge/Portfolio-Visit-orange?style=flat-square&logo=safari)](https://portfolio.taurbykaur.co.in/)

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
