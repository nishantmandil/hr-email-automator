import pandas as pd
import smtplib
import re
import os
import dns.resolver
from email.message import EmailMessage
from email_template import get_email_subject, get_email_body
from datetime import datetime
import time

# -----------------------
# CONFIG
# -----------------------

EMAIL = ""       # Your Gmail Address Here
PASSWORD = ""    # Your App Password Here

resume_path = "Nishant_SRE_Resume.pdf"   # Place resume PDF in same directory
excel_file = "Book 2.xlsx"         # Excel file with "Company Name" and "HR Email" columns
report_file = "email_report.xlsx"

# -----------------------
# REPORT HELPER
# -----------------------

def save_single_result(result):
    """
    Saves one result row immediately to disk after every action.
    If report exists — appends. If not — creates fresh.
    Ensures nothing is lost if script is stopped or crashes mid-run.
    """
    result["Sent At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_row = pd.DataFrame([result])

    if os.path.exists(report_file):
        existing_df = pd.read_excel(report_file)
        combined_df = pd.concat([existing_df, new_row], ignore_index=True)
        combined_df.to_excel(report_file, index=False)
    else:
        new_row.to_excel(report_file, index=False)

# -----------------------
# EMAIL VERIFICATION
# -----------------------

def smtp_verify_email(email, sender_email):
    """
    Step 1: Validates email format via regex.
    Step 2: Checks domain MX record via DNS (no port 25 needed).
    Step 3: Attempts SMTP handshake on port 25 if reachable.
    Falls back gracefully if port 25 is blocked.
    """
    # --- Format check ---
    pattern = r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
    if not re.match(pattern, email):
        print(f"  ↳ Invalid email format ❌")
        return False

    try:
        domain = email.strip().split("@")[1]

        # --- MX record check (DNS only, always works) ---
        try:
            mx_records = dns.resolver.resolve(domain, "MX")
            mx_host = str(sorted(mx_records, key=lambda r: r.preference)[0].exchange)
            print(f"  ↳ MX record found for '{domain}' ✅")
        except dns.resolver.NXDOMAIN:
            print(f"  ↳ Domain '{domain}' does not exist ❌")
            return False  # Hard fail — domain is completely fake
        except dns.resolver.NoAnswer:
            print(f"  ↳ Domain '{domain}' has no MX records ❌")
            return False  # Hard fail — domain can't receive email
        except dns.resolver.Timeout:
            print(f"  ↳ DNS timeout for '{domain}' — assuming valid ⚠️")
            return True   # Soft fail — don't skip on timeout
        except Exception as e:
            print(f"  ↳ DNS error for '{domain}': {e} — assuming valid ⚠️")
            return True   # Soft fail — unknown DNS error

        # --- SMTP handshake on port 25 (best effort) ---
        try:
            with smtplib.SMTP(mx_host, 25, timeout=10) as smtp:
                smtp.helo("gmail.com")
                smtp.mail(sender_email)
                code, message = smtp.rcpt(email)

                if code == 250:
                    print(f"  ↳ SMTP verified OK ✅")
                    return True
                else:
                    print(f"  ↳ SMTP rejected: {code} {message.decode()} ❌")
                    return False

        except smtplib.SMTPConnectError:
            print(f"  ↳ Port 25 blocked — relying on MX check only ⚠️")
            return True   # MX was valid, port just blocked
        except smtplib.SMTPServerDisconnected:
            print(f"  ↳ SMTP disconnected during check — relying on MX check only ⚠️")
            return True
        except TimeoutError:
            print(f"  ↳ SMTP timeout — relying on MX check only ⚠️")
            return True
        except Exception as e:
            print(f"  ↳ SMTP check failed ({type(e).__name__}: {e}) — relying on MX check only ⚠️")
            return True   # MX was valid, SMTP just unreachable

    except Exception as e:
        print(f"  ↳ Verification failed ({type(e).__name__}: {e}) — assuming valid ⚠️")
        return True

# -----------------------
# LOAD EXCEL
# -----------------------

df = pd.read_excel(excel_file)
print(f"Total HR emails found: {len(df)}\n")

results = []

# -----------------------
# SMTP LOGIN
# -----------------------

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
server.login(EMAIL, PASSWORD)
print("Logged into Gmail\n")

# -----------------------
# SEND EMAILS
# -----------------------

for index, row in df.iterrows():

    company = ""
    email = ""

    try:
        company = row["Company Name"]
        email = row["HR Email"]

        # Guard: empty email cell
        if pd.isna(email) or str(email).strip() == "":
            print(f"{index+1}. Skipping row {index+1} — no email address found")
            entry = {
                "Company Name": company,
                "HR Email": email,
                "Status": "Skipped",
                "Error": "Empty email address"
            }
            results.append(entry)
            save_single_result(entry)  # ← save immediately
            continue

        email = str(email).strip()

        # -----------------------
        # VERIFY BEFORE SENDING
        # -----------------------

        print(f"{index+1}. Verifying {email}...")
        is_valid = smtp_verify_email(email, EMAIL)

        if not is_valid:
            print(f"  ↳ Skipping {email} — address not found (no sleep)\n")
            entry = {
                "Company Name": company,
                "HR Email": email,
                "Status": "Skipped",
                "Error": "Address not found / SMTP rejected"
            }
            results.append(entry)
            save_single_result(entry)  # ← save immediately
            continue

        # -----------------------
        # BUILD EMAIL
        # -----------------------

        if pd.isna(company) or str(company).strip() == "":
            greeting = "Hi Hiring Team,"
        else:
            greeting = f"Hi {str(company).strip()} Team,"

        msg = EmailMessage()
        msg["Subject"] = get_email_subject()
        msg["From"] = EMAIL
        msg["To"] = email
        msg.set_content(get_email_body(company))

        with open(resume_path, "rb") as f:
            msg.add_attachment(
                f.read(),
                maintype="application",
                subtype="pdf",
                filename="Nishant_Mandil_SRE_Resume.pdf"
            )

        # -----------------------
        # SEND
        # -----------------------

        server.send_message(msg)
        print(f"  ↳ Sent to {email} ✅\n")

        entry = {
            "Company Name": company,
            "HR Email": email,
            "Status": "Sent",
            "Error": ""
        }
        results.append(entry)
        save_single_result(entry)  # ← save immediately

        time.sleep(45)  # Sleep only after successful send

    except smtplib.SMTPRecipientsRefused:
        print(f"  ↳ Recipient refused by Gmail: {email} (no sleep)\n")
        entry = {
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": "SMTPRecipientsRefused — address rejected by Gmail"
        }
        results.append(entry)
        save_single_result(entry)  # ← save immediately

    except smtplib.SMTPServerDisconnected:
        print(f"  ↳ SMTP server disconnected. Reconnecting...\n")
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(EMAIL, PASSWORD)
            print("  ↳ Reconnected successfully.")
        except Exception as reconnect_err:
            print(f"  ↳ Reconnect failed: {reconnect_err}")
        entry = {
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": "SMTPServerDisconnected — reconnected, retry manually"
        }
        results.append(entry)
        save_single_result(entry)  # ← save immediately

    except Exception as e:
        print(f"  ↳ Failed for {email}: {type(e).__name__}: {e}\n")
        entry = {
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": f"{type(e).__name__}: {str(e)}"
        }
        results.append(entry)
        save_single_result(entry)  # ← save immediately

# -----------------------
# CLOSE SERVER
# -----------------------

try:
    server.quit()
except Exception:
    pass

print("All emails processed.")
print(f"Report saved/updated at: {report_file}")
