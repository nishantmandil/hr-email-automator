# 📬 HR Email Automator

> Automated cold outreach tool for job seekers — verifies email addresses via SMTP before sending, attaches resume, and logs every result to Excel.

![Python](https://img.shields.io/badge/Python-3.8+-blue?style=flat-square&logo=python)
![License](https://img.shields.io/badge/License-MIT-green?style=flat-square)
![Status](https://img.shields.io/badge/Status-Active-brightgreen?style=flat-square)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey?style=flat-square)

---

## 🚀 What It Does

Most job seekers send cold emails blindly — many bounce, waste time, and hurt sender reputation.

**HR Email Automator** solves this by:

- ✅ **Verifying each email address** via SMTP handshake before sending — no bounces
- 📎 **Attaching your resume** automatically to every email
- 🧠 **Personalizing each email** with the company name from your Excel sheet
- ⏱️ **Rate-limiting sends** to avoid Gmail spam detection (45s between emails)
- 📊 **Logging every result** — Sent / Skipped / Failed — into a clean Excel report
- 🔁 **Auto-reconnecting** if Gmail drops the SMTP connection mid-run

---

## 📁 Project Structure

```
hr-email-automator/
│
├── main.py               # Core automation logic
├── email_template.py     # Email subject + body (edit this to change your message)
├── all_hr_emails.xlsx  # Input: list of HR emails + company names
├── Nishant_SRE_Resume.pdf  # Your resume (replace with your own)
└── email_report.xlsx     # Output: auto-generated after run
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

Your `all_hr_emails.xlsx` must have these two columns:

| Company Name | HR Email              |
|--------------|-----------------------|
| Google       | hr@google.com         |
| Razorpay     | talent@razorpay.com   |

### 4. Add your resume

Place your resume PDF in the root folder and update this line in `main.py`:

```python
resume_path = "Your_Resume.pdf"
```

### 5. Configure your Gmail credentials

In `main.py`, update:

```python
EMAIL = "your_email@gmail.com"
PASSWORD = "your_app_password"   # Use Gmail App Password, not your actual password
```

> ⚠️ **Important:** Use a [Gmail App Password](https://support.google.com/accounts/answer/185833), not your real Gmail password. Enable 2FA on your account first.

---

## ▶️ Run

```bash
python main.py
```

---

## 🔍 How Email Verification Works

Before sending any email, the tool performs a silent **SMTP handshake**:

```
1. Resolve MX record for the email's domain
2. Open SMTP connection to the mail server (port 25)
3. Issue RCPT TO command
4. If server returns 250 OK → send the email
   If server returns 550/551/553 → skip silently (no sleep, no send)
```

This happens **without sending any actual email** — it's just a connection check.

| Scenario | Action |
|----------|--------|
| SMTP returns 250 OK | ✅ Send email + wait 45s |
| SMTP returns 550 (not found) | ⏭️ Skip instantly, no sleep |
| Port 25 blocked / timeout | ✅ Assume valid, send anyway |
| Empty email cell in Excel | ⏭️ Skip with log entry |
| Gmail drops connection mid-run | 🔁 Auto-reconnect and continue |

---

## 📊 Output Report

After the run, `email_report.xlsx` is generated automatically:

| Company Name | HR Email           | Status  | Error                        |
|--------------|--------------------|---------|------------------------------|
| Google       | hr@google.com      | Sent    |                              |
| Razorpay     | bad@fake.com       | Skipped | Address not found / SMTP rejected |
| Infosys      | talent@infosys.com | Failed  | SMTPRecipientsRefused        |

---

## ✏️ Customizing Your Email

All email content lives in `email_template.py` — you never need to touch `main.py` for content changes.

```python
# email_template.py

def get_email_subject():
    return "Your Custom Subject Line"

def get_email_body(company):
    # greeting is auto-personalized from Excel
    ...
    return body
```

Just edit `email_template.py` and re-run. The logic in `main.py` stays untouched.

---

## 🛡️ Responsible Usage

- This tool is built for **personal job outreach only**
- Always respect the recipient's privacy
- Do not use for spam or bulk unsolicited marketing
- Gmail has daily sending limits (~500 emails/day for regular accounts)

---

## 🧰 Tech Stack

| Tool | Purpose |
|------|---------|
| `smtplib` | Sending emails + SMTP verification |
| `dnspython` | Resolving MX records for email domains |
| `pandas` | Reading Excel input, writing report |
| `openpyxl` | Excel file engine for pandas |
| `email.message` | Building MIME emails with attachments |

---

## 👤 Author

**Nishant Mandil**
Site Reliability Engineer @ Colt Technology Services

[![LinkedIn](https://img.shields.io/badge/LinkedIn-Connect-blue?style=flat-square&logo=linkedin)](https://www.linkedin.com/in/nishant-mandil-07b165159/)
[![GitHub](https://img.shields.io/badge/GitHub-Follow-black?style=flat-square&logo=github)](https://github.com/nishantmandil)
[![Portfolio](https://img.shields.io/badge/Portfolio-Visit-orange?style=flat-square)](https://portfolio.taurbykaur.co.in/)

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
