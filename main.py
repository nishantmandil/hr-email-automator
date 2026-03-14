import pandas as pd
import smtplib
import dns.resolver
from email.message import EmailMessage
import time

from email_template import get_email_subject, get_email_body  # ← import from File 1

# -----------------------
# CONFIG
# -----------------------

EMAIL = "" # Your Gmail Address Here
PASSWORD = "" # our App Password Here

resume_path = "Nishant_SRE_Resume.pdf" # Make sure to place your resume PDF in the same directory as this script or provide the correct path
excel_file = "all_hr_emails.xlsx" # Make sure to have the Excel file with columns "Company Name" and "HR Email" in the same directory or provide the correct path

# -----------------------
# EMAIL VERIFICATION
# -----------------------

def smtp_verify_email(email, sender_email):
    try:
        domain = email.strip().split("@")[1]

        try:
            mx_records = dns.resolver.resolve(domain, "MX")
            mx_host = str(sorted(mx_records, key=lambda r: r.preference)[0].exchange)
        except Exception as e:
            print(f"  ↳ No MX record for domain '{domain}': {e}")
            return False

        with smtplib.SMTP(mx_host, 25, timeout=10) as smtp:
            smtp.helo("gmail.com")
            smtp.mail(sender_email)
            code, message = smtp.rcpt(email)

            if code == 250:
                print(f"  ↳ SMTP verified OK: {email}")
                return True
            else:
                print(f"  ↳ SMTP rejected {email}: {code} {message.decode()}")
                return False

    except smtplib.SMTPConnectError:
        print(f"  ↳ Port 25 blocked for {email} — assuming valid")
        return True
    except smtplib.SMTPServerDisconnected:
        print(f"  ↳ Server disconnected during check for {email} — assuming valid")
        return True
    except TimeoutError:
        print(f"  ↳ SMTP check timed out for {email} — assuming valid")
        return True
    except Exception as e:
        print(f"  ↳ SMTP check failed for {email} ({type(e).__name__}: {e}) — assuming valid")
        return True

# -----------------------
# LOAD EXCEL
# -----------------------

df = pd.read_excel(excel_file)
print(f"Total HR emails found: {len(df)}")

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

        if pd.isna(email) or str(email).strip() == "":
            print(f"{index+1}. Skipping row {index+1} — no email address found")
            results.append({
                "Company Name": company,
                "HR Email": email,
                "Status": "Skipped",
                "Error": "Empty email address"
            })
            continue

        email = str(email).strip()

        # -----------------------
        # VERIFY BEFORE SENDING
        # -----------------------

        print(f"{index+1}. Verifying {email}...")
        is_valid = smtp_verify_email(email, EMAIL)

        if not is_valid:
            print(f"  ↳ Skipping {email} — address not found (no sleep)\n")
            results.append({
                "Company Name": company,
                "HR Email": email,
                "Status": "Skipped",
                "Error": "Address not found / SMTP rejected"
            })
            continue

        # -----------------------
        # BUILD EMAIL
        # -----------------------

        msg = EmailMessage()
        msg["Subject"] = get_email_subject()        # ← from email_template.py
        msg["From"] = EMAIL
        msg["To"] = email
        msg.set_content(get_email_body(company))    # ← from email_template.py

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
        print(f"  ↳ Sent to {email}\n")

        results.append({
            "Company Name": company,
            "HR Email": email,
            "Status": "Sent",
            "Error": ""
        })

        time.sleep(45)

    except smtplib.SMTPRecipientsRefused:
        print(f"  ↳ Recipient refused by Gmail: {email} (no sleep)\n")
        results.append({
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": "SMTPRecipientsRefused — address rejected by Gmail"
        })

    except smtplib.SMTPServerDisconnected:
        print(f"  ↳ SMTP server disconnected. Reconnecting...\n")
        try:
            server = smtplib.SMTP("smtp.gmail.com", 587)
            server.starttls()
            server.login(EMAIL, PASSWORD)
            print("  ↳ Reconnected successfully.")
        except Exception as reconnect_err:
            print(f"  ↳ Reconnect failed: {reconnect_err}")
        results.append({
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": "SMTPServerDisconnected — reconnected, retry manually"
        })

    except Exception as e:
        print(f"  ↳ Failed for {email}: {type(e).__name__}: {e}\n")
        results.append({
            "Company Name": company,
            "HR Email": email,
            "Status": "Failed",
            "Error": f"{type(e).__name__}: {str(e)}"
        })

# -----------------------
# CLOSE SERVER
# -----------------------

try:
    server.quit()
except Exception:
    pass

print("All emails processed.")

# -----------------------
# SAVE REPORT
# -----------------------

report_df = pd.DataFrame(results)
report_df.to_excel("email_report.xlsx", index=False)
print("Report saved as email_report.xlsx")
