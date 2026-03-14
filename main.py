import pandas as pd
import smtplib
import imaplib
import email as email_lib
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

EMAIL = "" # Use your Gmail address here (app password must be set up for this to work)
PASSWORD = "" # Use an app password generated from your Google Account security settings.

resume_path = "Nishant_SRE_Resume.pdf" # Make sure the resume file is in the same directory as this script, or provide the correct path.
excel_file = "all_hr_emails.xlsx" # This Excel file should have columns "Company Name" and "HR Email" with the relevant data.
report_file = "email_report.xlsx" # This will be created/updated with results as the script runs.

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


def update_last_result(sent_email, new_status, new_error):
    """
    Updates the most recent row for a given email in the report.
    Used to change 'Sent' to 'Bounced' after bounce is detected.
    """
    if not os.path.exists(report_file):
        return

    df = pd.read_excel(report_file)

    # Find the last row matching this email
    mask = df["HR Email"] == sent_email
    if mask.any():
        last_index = df[mask].index[-1]
        df.at[last_index, "Status"] = new_status
        df.at[last_index, "Error"] = new_error
        df.to_excel(report_file, index=False)
        print(f"  ↳ Report updated — marked {sent_email} as {new_status} ⚠️")

# -----------------------
# BOUNCE CHECKER
# -----------------------

def check_bounce_replies(mail_user, mail_password, sent_email):
    """
    Connects to Gmail via IMAP during the 45s sleep window.
    Looks for unread bounce/failure emails from mailer-daemon
    that mention the address we just sent to.
    Returns True if a bounce was detected.
    """
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(mail_user, mail_password)
        mail.select("inbox")

        # Search for unread emails from mailer-daemon
        status, messages = mail.search(None, '(UNSEEN FROM "mailer-daemon@googlemail.com")')

        if status != "OK" or not messages[0]:
            mail.logout()
            return False

        bounce_found = False

        for msg_id in messages[0].split():
            status, msg_data = mail.fetch(msg_id, "(RFC822)")
            if status != "OK":
                continue

            raw = msg_data[0][1]
            msg = email_lib.message_from_bytes(raw)

            # Extract full body text
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body += part.get_payload(decode=True).decode(errors="ignore")
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")

            # Check if bounce mentions our sent email address
            if sent_email.lower() in body.lower():
                print(f"  ↳ Bounce detected for {sent_email} in inbox ⚠️")
                mail.store(msg_id, "+FLAGS", "\\Seen")  # Mark as read
                bounce_found = True
                break

        mail.logout()
        return bounce_found

    except Exception as e:
        print(f"  ↳ IMAP bounce check failed ({type(e).__name__}: {e}) — skipping check")
        return False

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
            return False
        except dns.resolver.NoAnswer:
            print(f"  ↳ Domain '{domain}' has no MX records ❌")
            return False
        except dns.resolver.Timeout:
            print(f"  ↳ DNS timeout for '{domain}' — assuming valid ⚠️")
            return True
        except Exception as e:
            print(f"  ↳ DNS error for '{domain}': {e} — assuming valid ⚠️")
            return True

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
            return True
        except smtplib.SMTPServerDisconnected:
            print(f"  ↳ SMTP disconnected during check — relying on MX check only ⚠️")
            return True
        except TimeoutError:
            print(f"  ↳ SMTP timeout — relying on MX check only ⚠️")
            return True
        except Exception as e:
            print(f"  ↳ SMTP check failed ({type(e).__name__}: {e}) — relying on MX check only ⚠️")
            return True

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
            save_single_result(entry)
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
            save_single_result(entry)
            continue

        # -----------------------
        # BUILD EMAIL
        # -----------------------

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
        print(f"  ↳ Sent to {email} ✅")

        entry = {
            "Company Name": company,
            "HR Email": email,
            "Status": "Sent",
            "Error": ""
        }
        results.append(entry)
        save_single_result(entry)  # ← save as Sent immediately

        # -----------------------
        # SLEEP + BOUNCE CHECK
        # -----------------------

        print(f"  ↳ Waiting 45s and checking inbox for bounce replies...")
        time.sleep(45)

        bounce_detected = check_bounce_replies(EMAIL, PASSWORD, email)

        if bounce_detected:
            # Update the report row from Sent → Bounced
            update_last_result(
                sent_email=email,
                new_status="Bounced",
                new_error="Address not found — bounce received from mailer-daemon"
            )
            print()
        else:
            print(f"  ↳ No bounce detected ✅\n")

    except smtplib.SMTPRecipientsRefused:
        print(f"  ↳ Recipient refused by Gmail: {email} (no sleep)\n")
        entry = {
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": "SMTPRecipientsRefused — address rejected by Gmail"
        }
        results.append(entry)
        save_single_result(entry)

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
        save_single_result(entry)

    except Exception as e:
        print(f"  ↳ Failed for {email}: {type(e).__name__}: {e}\n")
        entry = {
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": f"{type(e).__name__}: {str(e)}"
        }
        results.append(entry)
        save_single_result(entry)

# -----------------------
# CLOSE SERVER
# -----------------------

try:
    server.quit()
except Exception:
    pass

print("All emails processed.")
print(f"Report saved/updated at: {report_file}")
