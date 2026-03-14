from flask import Flask, render_template, request, jsonify, Response
import pandas as pd
import smtplib
import imaplib
import email as email_lib
import re
import os
import ast
import dns.resolver
from email.message import EmailMessage
from datetime import datetime
import time
import threading
import json
import queue

app = Flask(__name__)

# Global state
job_queue = queue.Queue()
is_running = False
stop_flag = False
log_queue = queue.Queue()
stats = {"sent": 0, "skipped": 0, "bounced": 0, "failed": 0, "total": 0, "current": 0}

report_file = "email_report.xlsx"

# -----------------------
# READ CONFIG FROM main.py
# -----------------------

def read_config_from_main():
    """
    Parses main.py and extracts EMAIL, PASSWORD, resume_path,
    excel_file values. Returns a dict with whatever it finds.
    Falls back to empty string if not found or file missing.
    """
    config = {
        "email": "",
        "password": "",
        "resume": "",
        "excel": ""
    }

    main_path = os.path.join(os.path.dirname(__file__), "main.py")
    if not os.path.exists(main_path):
        return config

    try:
        with open(main_path, "r", encoding="utf-8") as f:
            source = f.read()

        tree = ast.parse(source)

        for node in ast.walk(tree):
            if isinstance(node, ast.Assign):
                for target in node.targets:
                    if isinstance(target, ast.Name):
                        name = target.id
                        # Only extract simple string assignments
                        if isinstance(node.value, ast.Constant) and isinstance(node.value.value, str):
                            val = node.value.value.strip()
                            if name == "EMAIL":
                                config["email"] = val
                            elif name == "PASSWORD":
                                config["password"] = val
                            elif name == "resume_path":
                                config["resume"] = val
                            elif name == "excel_file":
                                config["excel"] = val
    except Exception:
        pass  # If parsing fails, return whatever we have

    return config


# -----------------------
# REPORT HELPERS
# -----------------------

def save_single_result(result):
    result["Sent At"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    new_row = pd.DataFrame([result])
    if os.path.exists(report_file):
        existing_df = pd.read_excel(report_file)
        combined_df = pd.concat([existing_df, new_row], ignore_index=True)
        combined_df.to_excel(report_file, index=False)
    else:
        new_row.to_excel(report_file, index=False)


def update_last_result(sent_email, new_status, new_error):
    if not os.path.exists(report_file):
        return
    df = pd.read_excel(report_file)
    mask = df["HR Email"] == sent_email
    if mask.any():
        last_index = df[mask].index[-1]
        df.at[last_index, "Status"] = new_status
        df.at[last_index, "Error"] = new_error
        df.to_excel(report_file, index=False)


# -----------------------
# LOGGING
# -----------------------

def log(msg, level="info"):
    timestamp = datetime.now().strftime("%H:%M:%S")
    log_queue.put(json.dumps({"time": timestamp, "msg": msg, "level": level}))


# -----------------------
# BOUNCE CHECKER
# -----------------------

def check_bounce_replies(mail_user, mail_password, sent_email):
    try:
        mail = imaplib.IMAP4_SSL("imap.gmail.com")
        mail.login(mail_user, mail_password)
        mail.select("inbox")
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
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body += part.get_payload(decode=True).decode(errors="ignore")
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore")
            if sent_email.lower() in body.lower():
                mail.store(msg_id, "+FLAGS", "\\Seen")
                bounce_found = True
                break
        mail.logout()
        return bounce_found
    except Exception as e:
        log(f"IMAP bounce check failed: {e}", "warn")
        return False


# -----------------------
# EMAIL VERIFICATION
# -----------------------

def smtp_verify_email(email, sender_email):
    pattern = r'^[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}$'
    if not re.match(pattern, email):
        return False
    try:
        domain = email.strip().split("@")[1]
        try:
            mx_records = dns.resolver.resolve(domain, "MX")
            mx_host = str(sorted(mx_records, key=lambda r: r.preference)[0].exchange)
        except dns.resolver.NXDOMAIN:
            return False
        except dns.resolver.NoAnswer:
            return False
        except (dns.resolver.Timeout, Exception):
            return True
        try:
            with smtplib.SMTP(mx_host, 25, timeout=10) as smtp:
                smtp.helo("gmail.com")
                smtp.mail(sender_email)
                code, _ = smtp.rcpt(email)
                return code == 250
        except Exception:
            return True
    except Exception:
        return True


# -----------------------
# MAIN SEND JOB
# -----------------------

def run_send_job(config):
    global is_running, stop_flag, stats

    EMAIL = config["email"]
    PASSWORD = config["password"]
    resume_path = config["resume"]
    excel_file = config["excel"]
    subject = config["subject"]
    body = config["body"]

    is_running = True
    stop_flag = False
    stats = {"sent": 0, "skipped": 0, "bounced": 0, "failed": 0, "total": 0, "current": 0}

    try:
        df = pd.read_excel(excel_file)
        stats["total"] = len(df)
        log(f"Loaded {len(df)} emails from {excel_file}", "info")
    except Exception as e:
        log(f"Failed to load Excel file: {e}", "error")
        is_running = False
        return

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL, PASSWORD)
        log("Logged into Gmail successfully", "success")
    except Exception as e:
        log(f"Gmail login failed: {e}", "error")
        is_running = False
        return

    for index, row in df.iterrows():
        if stop_flag:
            log("Job stopped by user", "warn")
            break

        stats["current"] = index + 1
        company = ""
        email_addr = ""

        try:
            company = row["Company Name"]
            email_addr = row["HR Email"]

            if pd.isna(email_addr) or str(email_addr).strip() == "":
                log(f"Row {index+1}: Skipped — empty email", "warn")
                entry = {"Company Name": company, "HR Email": email_addr, "Status": "Skipped", "Error": "Empty email address"}
                save_single_result(entry)
                stats["skipped"] += 1
                continue

            email_addr = str(email_addr).strip()
            log(f"[{index+1}/{stats['total']}] Verifying {email_addr}...", "info")

            is_valid = smtp_verify_email(email_addr, EMAIL)

            if not is_valid:
                log(f"Skipped {email_addr} — domain invalid or rejected", "warn")
                entry = {"Company Name": company, "HR Email": email_addr, "Status": "Skipped", "Error": "Address not found / SMTP rejected"}
                save_single_result(entry)
                stats["skipped"] += 1
                continue

            msg = EmailMessage()
            msg["Subject"] = subject
            msg["From"] = EMAIL
            msg["To"] = email_addr
            msg.set_content(body)

            if resume_path and os.path.exists(resume_path):
                with open(resume_path, "rb") as f:
                    msg.add_attachment(f.read(), maintype="application", subtype="pdf",
                                       filename=os.path.basename(resume_path))

            server.send_message(msg)
            log(f"Sent to {email_addr}", "success")

            entry = {"Company Name": company, "HR Email": email_addr, "Status": "Sent", "Error": ""}
            save_single_result(entry)
            stats["sent"] += 1

            log(f"Waiting 45s — checking for bounces...", "info")
            for _ in range(45):
                if stop_flag:
                    break
                time.sleep(1)

            if not stop_flag:
                bounce = check_bounce_replies(EMAIL, PASSWORD, email_addr)
                if bounce:
                    update_last_result(email_addr, "Bounced", "Address not found — bounce from mailer-daemon")
                    stats["sent"] -= 1
                    stats["bounced"] += 1
                    log(f"Bounce detected for {email_addr}", "warn")
                else:
                    log(f"No bounce for {email_addr} ✓", "success")

        except smtplib.SMTPRecipientsRefused:
            log(f"Gmail refused recipient: {email_addr}", "error")
            entry = {"Company Name": company, "HR Email": email_addr, "Status": "Failed", "Error": "SMTPRecipientsRefused"}
            save_single_result(entry)
            stats["failed"] += 1

        except smtplib.SMTPServerDisconnected:
            log("SMTP disconnected — reconnecting...", "warn")
            try:
                server = smtplib.SMTP("smtp.gmail.com", 587)
                server.starttls()
                server.login(EMAIL, PASSWORD)
                log("Reconnected successfully", "success")
            except Exception as e:
                log(f"Reconnect failed: {e}", "error")
            entry = {"Company Name": company, "HR Email": email_addr, "Status": "Failed", "Error": "SMTPServerDisconnected"}
            save_single_result(entry)
            stats["failed"] += 1

        except Exception as e:
            log(f"Error for {email_addr}: {e}", "error")
            entry = {"Company Name": company, "HR Email": email_addr, "Status": "Failed", "Error": str(e)}
            save_single_result(entry)
            stats["failed"] += 1

    try:
        server.quit()
    except Exception:
        pass

    log("All emails processed", "success")
    is_running = False


# -----------------------
# ROUTES
# -----------------------

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/config")
def get_config():
    """Called by frontend on load — returns config parsed from main.py."""
    return jsonify(read_config_from_main())


@app.route("/api/start", methods=["POST"])
def start_job():
    global is_running
    if is_running:
        return jsonify({"error": "Job already running"}), 400
    data = request.json
    thread = threading.Thread(target=run_send_job, args=(data,), daemon=True)
    thread.start()
    return jsonify({"status": "started"})


@app.route("/api/stop", methods=["POST"])
def stop_job():
    global stop_flag
    stop_flag = True
    return jsonify({"status": "stopping"})


@app.route("/api/stats")
def get_stats():
    return jsonify({**stats, "running": is_running})


@app.route("/api/logs")
def stream_logs():
    def generate():
        while True:
            try:
                msg = log_queue.get(timeout=30)
                yield f"data: {msg}\n\n"
            except queue.Empty:
                yield f"data: {json.dumps({'ping': True})}\n\n"
    return Response(generate(), mimetype="text/event-stream")


@app.route("/api/report")
def get_report():
    if not os.path.exists(report_file):
        return jsonify([])
    df = pd.read_excel(report_file)
    df = df.fillna("")
    return jsonify(df.tail(50).to_dict(orient="records"))


if __name__ == "__main__":
    os.makedirs("templates", exist_ok=True)
    app.run(debug=False, port=5000, threaded=True)
