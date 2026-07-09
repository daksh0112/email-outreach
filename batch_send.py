#!/usr/bin/env python3
"""
batch_send.py — Gmail SMTP batch email sender for job outreach
Run: PYTHONUNBUFFERED=1 python batch_send.py
"""

import json
import os
import sys
import time
import smtplib
import imaplib
import email as email_lib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from datetime import datetime

# ── CONFIGURATION ─────────────────────────────────────────────────────────────
# Supports both local PC and GitHub Actions (env vars take priority)
_ON_GITHUB = os.environ.get("GITHUB_ACTIONS") == "true"

GMAIL_USER         = os.environ.get("GMAIL_USER",         "dakshmanuarya@gmail.com")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD", "fxzofaoxgaxdavjj")

if _ON_GITHUB:
    # Relative paths — files live in the repo root on GitHub Actions
    RESUME_PATH    = "Daksh_Resume.pdf"
    CONTACTS_JSON  = "contacts.json"
    SENT_LIST_FILE = "sent_emails_list.txt"
    PROGRESS_FILE  = "send_progress.txt"
    FAILED_FILE    = "failed_remaining.json"
else:
    # Local PC paths
    RESUME_PATH    = r"C:\Users\user\Downloads\Daksh_Resume (1).pdf"
    CONTACTS_JSON  = r"C:\Users\user\email_outreach\contacts.json"
    SENT_LIST_FILE = r"C:\Users\user\email_outreach\sent_emails_list.txt"
    PROGRESS_FILE  = r"C:\Users\user\email_outreach\send_progress.txt"
    FAILED_FILE    = r"C:\Users\user\email_outreach\failed_remaining.json"

EXCLUDED_DOMAINS   = []

SENDER_NAME        = "Daksh Manu Arya"
LINKEDIN_URL       = "https://www.linkedin.com/in/daksh-manu-arya-88b573192/"
PHONE              = "+91-9416129284"

DELAY_BETWEEN      = 10    # seconds between each email
PAUSE_AFTER        = 450   # pause after this many emails
PAUSE_DURATION     = 60    # pause duration in seconds
MAX_RETRIES        = 3
SMTP_TIMEOUT       = 30
PROGRESS_PRINT_EVERY = 25
DAILY_LIMIT        = 8     # 8 per run × 2 runs/day = 16/day → 237 contacts in 15 days


# ── EMAIL SUBJECT ──────────────────────────────────────────────────────────────
def build_subject(company, referred_by=None):
    base = f"Backend Developer for {company}"
    if referred_by:
        base = f"Backend Developer for {company} — Referred by {referred_by}"
    return base


# ── HTML BODY ──────────────────────────────────────────────────────────────────
def build_html(first_name=None, company=None, referred_by=None, custom_note=None):
    if first_name and first_name.lower() not in ("—", "hiring", "university", "careers", "campus", "ta"):
        greeting = f"Hi {first_name},"
    elif company:
        greeting = f"Hi Hiring Team at {company},"
    else:
        greeting = "Hi there,"

    referral_note = f'<p style="margin:0 0 14px;padding:10px 14px;background:#f0f6ff;border-left:3px solid #3b82f6;border-radius:0 4px 4px 0;font-size:14px;">I was referred to you by <strong>{referred_by}</strong>.</p>' if referred_by else ""
    custom_block  = f'<p style="margin:0 0 14px;">{custom_note}</p>' if custom_note else ""
    company_name  = company or "your company"

    return f"""<!DOCTYPE html>
<html>
<head><meta charset="UTF-8"></head>
<body style="margin:0;padding:24px 0;background:#f3f4f6;font-family:Arial,Helvetica,sans-serif;">

<!-- Outer wrapper -->
<table width="100%" cellpadding="0" cellspacing="0" border="0">
<tr><td align="center">
<table width="600" cellpadding="0" cellspacing="0" border="0" style="max-width:600px;background:#ffffff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,0.08);">

  <!-- ── DARK HEADER ── -->
  <tr>
    <td style="background:#0f2044;padding:28px 32px 24px;">
      <p style="margin:0 0 6px;font-size:11px;font-weight:700;letter-spacing:2px;color:#93b4e8;text-transform:uppercase;">Job Application</p>
      <p style="margin:0 0 6px;font-size:26px;font-weight:700;color:#ffffff;line-height:1.2;">Daksh Manu Arya</p>
      <p style="margin:0;font-size:13px;color:#a8c0e0;line-height:1.5;">Backend Engineer &nbsp;|&nbsp; Distributed Systems &nbsp;|&nbsp; Kubernetes &nbsp;|&nbsp; Python &amp; Rust</p>
    </td>
  </tr>

  <!-- ── BODY ── -->
  <tr>
    <td style="padding:28px 32px 0;">

      <p style="margin:0 0 16px;font-size:15px;color:#1a1a1a;line-height:1.6;">{greeting}</p>

      {referral_note}
      {custom_block}

      <p style="margin:0 0 20px;font-size:15px;color:#1a1a1a;line-height:1.7;">
        I'm a Backend Engineer at <strong>E2E Networks</strong>, where I build distributed systems and Kubernetes-native infrastructure.
        I came across <strong>{company_name}</strong> and wanted to reach out directly about backend opportunities — I've attached my resume below.
      </p>

      <!-- ── TWO-COLUMN TABLE ── -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:20px;border-radius:6px;overflow:hidden;border:1px solid #e5e7eb;">
        <tr style="background:#f8f9fb;">
          <td width="50%" style="padding:12px 16px;border-right:1px solid #e5e7eb;">
            <p style="margin:0 0 10px;font-size:12px;font-weight:700;letter-spacing:1px;color:#6b7280;text-transform:uppercase;">Role Fit</p>
            <p style="margin:0;font-size:13px;color:#1a1a1a;line-height:1.8;">
              &bull; Backend Software Engineer (SDE)<br>
              &bull; Platform / Infrastructure Engineer<br>
              &bull; Cloud Backend Developer<br>
              &bull; Distributed Systems Engineer
            </p>
          </td>
          <td width="50%" style="padding:12px 16px;">
            <p style="margin:0 0 10px;font-size:12px;font-weight:700;letter-spacing:1px;color:#6b7280;text-transform:uppercase;">Tech Stack</p>
            <p style="margin:0;font-size:13px;color:#1a1a1a;line-height:1.8;">
              &bull; Python, Rust, C/C++<br>
              &bull; Django, gRPC, Celery<br>
              &bull; Kubernetes, Docker, Temporal<br>
              &bull; ClickHouse, MySQL, NATS JetStream<br>
              &bull; OpenTelemetry, Grafana, Vector
            </p>
          </td>
        </tr>
      </table>

      <!-- ── KEY HIGHLIGHTS ── -->
      <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom:20px;border-radius:6px;overflow:hidden;border:1px solid #e5e7eb;">
        <tr style="background:#f8f9fb;">
          <td style="padding:12px 16px 6px;">
            <p style="margin:0 0 12px;font-size:12px;font-weight:700;letter-spacing:1px;color:#6b7280;text-transform:uppercase;">Key Highlights</p>

            <p style="margin:0 0 10px;font-size:13px;color:#1a1a1a;line-height:1.6;">
              <strong style="color:#0f2044;">Kubernetes Observability Platform:</strong>
              Processes <strong>1M+ events/min</strong> via a centralized OTLP gateway backed by ClickHouse and NATS JetStream.
            </p>
            <p style="margin:0 0 10px;font-size:13px;color:#1a1a1a;line-height:1.6;">
              <strong style="color:#0f2044;">Rust gRPC Auth Service:</strong>
              1M-entry token cache achieving <strong>p99 &lt;5ms</strong> cache-hit latency with fail-closed authorization.
            </p>
            <p style="margin:0 0 10px;font-size:13px;color:#1a1a1a;line-height:1.6;">
              <strong style="color:#0f2044;">VM Autoscaling Framework:</strong>
              KubeVirt-based ASG with CPU, memory, and custom-metric-based scaling policies.
            </p>
            <p style="margin:0 0 6px;font-size:13px;color:#1a1a1a;line-height:1.6;">
              <strong style="color:#0f2044;">K8s Deployment Framework:</strong>
              Supports 50+ one-click apps with tenant isolation and Temporal lifecycle management.
            </p>
          </td>
        </tr>
      </table>

      <!-- ── BUTTONS ── -->
      <table cellpadding="0" cellspacing="0" border="0" style="margin-bottom:24px;">
        <tr>
          <td style="padding-right:8px;">
            <a href="mailto:dakshmanuarya@gmail.com" style="display:inline-block;padding:9px 18px;background:#1d4ed8;color:#ffffff;font-size:13px;font-weight:600;text-decoration:none;border-radius:5px;">Contact Me</a>
          </td>
          <td style="padding-right:8px;">
            <a href="{LINKEDIN_URL}" style="display:inline-block;padding:9px 18px;background:#0a66c2;color:#ffffff;font-size:13px;font-weight:600;text-decoration:none;border-radius:5px;">LinkedIn</a>
          </td>
          <td>
            <a href="https://github.com/daksh0112" style="display:inline-block;padding:9px 18px;background:#24292e;color:#ffffff;font-size:13px;font-weight:600;text-decoration:none;border-radius:5px;">GitHub</a>
          </td>
        </tr>
      </table>

    </td>
  </tr>

  <!-- ── FOOTER ── -->
  <tr>
    <td style="padding:20px 32px 28px;border-top:1px solid #e5e7eb;">
      <p style="margin:0 0 6px;font-size:14px;color:#374151;line-height:1.6;">
        I'd love to be referred for any open backend or SDE roles at <strong>{company_name}</strong>.
        My resume is attached — and if you have any feedback on my profile or suggestions on how I could improve my chances, I'd genuinely love to hear it. Any guidance would mean a lot.
      </p>
      <p style="margin:14px 0 4px;font-size:14px;color:#374151;">Thank you for your time and consideration.</p>
      <p style="margin:0;font-size:14px;font-weight:700;color:#0f2044;">Daksh Manu Arya</p>
      <p style="margin:2px 0 0;font-size:13px;color:#6b7280;">Noida, Uttar Pradesh</p>
      <p style="margin:6px 0 0;font-size:13px;color:#374151;">
        <a href="tel:+919416129284" style="color:#1d4ed8;text-decoration:none;">+91 9416129284</a>
        &nbsp;|&nbsp;
        <a href="mailto:dakshmanuarya@gmail.com" style="color:#1d4ed8;text-decoration:none;">dakshmanuarya@gmail.com</a>
        &nbsp;|&nbsp;
        <a href="{LINKEDIN_URL}" style="color:#1d4ed8;text-decoration:none;">LinkedIn</a>
      </p>
    </td>
  </tr>

</table>
</td></tr>
</table>

</body>
</html>"""


# ── SMTP CONNECTION ────────────────────────────────────────────────────────────
def connect_smtp():
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            server = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=SMTP_TIMEOUT)
            server.login(GMAIL_USER, GMAIL_APP_PASSWORD)
            return server
        except Exception as e:
            print(f"  [SMTP] Connection attempt {attempt}/{MAX_RETRIES} failed: {e}", flush=True)
            if attempt < MAX_RETRIES:
                time.sleep(5)
    return None


# ── BUILD MIME MESSAGE ─────────────────────────────────────────────────────────
def build_message(to_email, to_name, company, referred_by=None, custom_note=None):
    msg = MIMEMultipart("alternative")
    msg["From"]    = f"{SENDER_NAME} <{GMAIL_USER}>"
    msg["To"]      = to_email
    msg["Subject"] = build_subject(company, referred_by)

    first_name = to_name.split()[0] if to_name else None
    html = build_html(first_name, company, referred_by, custom_note)

    msg_related = MIMEMultipart("related")
    msg_related.attach(MIMEText(html, "html"))

    msg.attach(MIMEText("Please view this email in an HTML-capable client.", "plain"))
    msg.attach(msg_related)

    # Attach resume
    if RESUME_PATH and Path(RESUME_PATH).exists():
        with open(RESUME_PATH, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f'attachment; filename="{Path(RESUME_PATH).name}"')
        msg.attach(part)

    return msg


# ── SEND ONE EMAIL (with retry) ───────────────────────────────────────────────
def send_one(server, to_email, to_name, company, referred_by=None, custom_note=None):
    msg = build_message(to_email, to_name, company, referred_by, custom_note)
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            server.sendmail(GMAIL_USER, to_email, msg.as_string())
            return True, None
        except smtplib.SMTPServerDisconnected:
            print(f"  [SMTP] Disconnected on attempt {attempt}, reconnecting...", flush=True)
            new_server = connect_smtp()
            if new_server is None:
                return False, "SMTP reconnect failed"
            server.__dict__.update(new_server.__dict__)
        except Exception as e:
            if attempt == MAX_RETRIES:
                return False, str(e)
            time.sleep(3)
    return False, "Max retries exceeded"


# ── LOAD SENT LIST ─────────────────────────────────────────────────────────────
def load_sent():
    if Path(SENT_LIST_FILE).exists():
        return set(Path(SENT_LIST_FILE).read_text().splitlines())
    return set()


def log_sent(email_addr, name, company):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(SENT_LIST_FILE, "a") as f:
        f.write(f"{email_addr}\n")
    with open(PROGRESS_FILE, "a") as f:
        f.write(f"{timestamp} | {email_addr} | {name} | {company}\n")


# ── MAIN BATCH SEND ────────────────────────────────────────────────────────────
def run_batch():
    contacts = json.loads(Path(CONTACTS_JSON).read_text(encoding="utf-8"))
    sent     = load_sent()

    # Filter: not already sent, not excluded domain
    todo = [
        c for c in contacts
        if c["email"] not in sent
        and not any(c["email"].lower().endswith(d) for d in EXCLUDED_DOMAINS)
    ]

    print(f"\n{'='*55}", flush=True)
    print(f"  Total contacts : {len(contacts)}", flush=True)
    print(f"  Already sent   : {len(sent)}", flush=True)
    print(f"  To send now    : {min(len(todo), DAILY_LIMIT)}", flush=True)
    print(f"{'='*55}\n", flush=True)

    # Clear progress file for this run
    Path(PROGRESS_FILE).write_text("")

    server = connect_smtp()
    if server is None:
        print("ERROR: Could not connect to Gmail. Check credentials.", flush=True)
        sys.exit(1)

    sent_count  = 0
    failed_list = []

    for i, contact in enumerate(todo):
        if sent_count >= DAILY_LIMIT:
            print(f"\n[LIMIT] Reached daily limit of {DAILY_LIMIT}. Stopping.", flush=True)
            break

        to_email = contact["email"]
        to_name  = contact.get("name") or contact.get("first_name", "")
        company  = contact.get("company", "your company")

        ok, err = send_one(server, to_email, to_name, company)

        if ok:
            log_sent(to_email, to_name, company)
            sent_count += 1
            if sent_count % PROGRESS_PRINT_EVERY == 0:
                print(f"  [{sent_count}/{min(len(todo), DAILY_LIMIT)}] Sent to {to_email} ({company})", flush=True)
        else:
            print(f"  [FAIL] {to_email} — {err}", flush=True)
            failed_list.append({**contact, "error": err})

        # Rate limiting
        if sent_count > 0 and sent_count % PAUSE_AFTER == 0:
            print(f"\n[PAUSE] {sent_count} emails sent — pausing {PAUSE_DURATION}s to respect Gmail limits...\n", flush=True)
            time.sleep(PAUSE_DURATION)
        else:
            time.sleep(DELAY_BETWEEN)

    try:
        server.quit()
    except Exception:
        pass

    # Save failed
    if failed_list:
        Path(FAILED_FILE).write_text(json.dumps(failed_list, indent=2), encoding="utf-8")

    print(f"\n{'='*55}", flush=True)
    print(f"  Run complete", flush=True)
    print(f"  Sent    : {sent_count}", flush=True)
    print(f"  Failed  : {len(failed_list)} (see {FAILED_FILE})", flush=True)
    print(f"  Remaining: {max(0, len(todo) - sent_count - len(failed_list))}", flush=True)
    print(f"{'='*55}\n", flush=True)


# ── ONE-OFF SEND ───────────────────────────────────────────────────────────────
def send_single(to_email, to_name, company, referred_by=None, custom_note=None):
    sent = load_sent()
    if to_email in sent:
        print(f"[SKIP] {to_email} already in sent list.", flush=True)
        return

    server = connect_smtp()
    if server is None:
        print("ERROR: Could not connect to Gmail.", flush=True)
        return

    ok, err = send_one(server, to_email, to_name, company, referred_by, custom_note)
    if ok:
        log_sent(to_email, to_name, company)
        print(f"[OK] Sent to {to_email} at {company}", flush=True)
    else:
        print(f"[FAIL] {to_email} — {err}", flush=True)
    try:
        server.quit()
    except Exception:
        pass


# ── SAVE TO DRAFTS ─────────────────────────────────────────────────────────────
def save_draft(to_email, to_name, company):
    msg = build_message(to_email, to_name, company)
    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        imap.select('"[Gmail]/Drafts"')
        imap.append(
            '"[Gmail]/Drafts"', "\\Draft",
            imaplib.Time2Internaldate(time.time()),
            msg.as_bytes()
        )
        imap.logout()
        print(f"[DRAFT] Saved draft for {to_email}", flush=True)
    except Exception as e:
        print(f"[DRAFT ERROR] {e}", flush=True)


# ── FETCH & CATEGORIZE REPLIES ─────────────────────────────────────────────────
def fetch_replies():
    try:
        imap = imaplib.IMAP4_SSL("imap.gmail.com")
        imap.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        imap.select("INBOX")
        _, data = imap.search(None, 'ALL')
        msg_ids = data[0].split()[-50:]   # last 50 messages

        real_replies, ooo, auto, undeliverable, job_alerts = [], [], [], [], []

        for mid in msg_ids:
            _, msg_data = imap.fetch(mid, "(RFC822)")
            raw = msg_data[0][1]
            msg = email_lib.message_from_bytes(raw)
            subject = str(msg.get("Subject", "")).lower()
            sender  = str(msg.get("From", "")).lower()
            body    = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode(errors="ignore").lower()
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="ignore").lower()

            entry = {"from": sender, "subject": msg.get("Subject"), "snippet": body[:200]}

            if any(k in body for k in ["out of office", "ooo", "on leave", "will be back", "auto-reply"]):
                ooo.append(entry)
            elif any(k in subject for k in ["undelivered", "delivery failed", "mailer-daemon"]):
                undeliverable.append(entry)
            elif any(k in subject for k in ["job alert", "new jobs", "recommended jobs"]):
                job_alerts.append(entry)
            elif any(k in body for k in ["noreply", "no-reply", "donotreply"]):
                auto.append(entry)
            else:
                real_replies.append(entry)

        imap.logout()

        print(f"\n[INBOX REPORT]", flush=True)
        print(f"  Real replies    : {len(real_replies)}", flush=True)
        print(f"  OOO / Referrals : {len(ooo)}", flush=True)
        print(f"  Auto-replies    : {len(auto)}", flush=True)
        print(f"  Undeliverable   : {len(undeliverable)}", flush=True)
        print(f"  Job alerts      : {len(job_alerts)}", flush=True)

        if real_replies:
            print("\n  -- Real Replies --")
            for r in real_replies:
                print(f"  From: {r['from']}\n  Subject: {r['subject']}\n  Snippet: {r['snippet'][:100]}\n")
        if ooo:
            print("\n  -- OOO / Referrals (check for redirected contact) --")
            for r in ooo:
                print(f"  From: {r['from']}\n  Subject: {r['subject']}\n  Snippet: {r['snippet'][:150]}\n")

    except Exception as e:
        print(f"[INBOX ERROR] {e}", flush=True)


# ── PROGRESS CHECK ─────────────────────────────────────────────────────────────
def check_progress():
    sent = load_sent()
    contacts = json.loads(Path(CONTACTS_JSON).read_text(encoding="utf-8"))
    remaining = [c for c in contacts if c["email"] not in sent]
    print(f"\n[PROGRESS]", flush=True)
    print(f"  Sent      : {len(sent)}", flush=True)
    print(f"  Total     : {len(contacts)}", flush=True)
    print(f"  Remaining : {len(remaining)}", flush=True)
    if Path(PROGRESS_FILE).exists():
        lines = Path(PROGRESS_FILE).read_text().splitlines()
        if lines:
            print(f"  Last sent : {lines[-1]}", flush=True)


# ── ENTRY POINT ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="Batch email sender for job outreach")
    parser.add_argument("--run",      action="store_true",  help="Run batch send")
    parser.add_argument("--progress", action="store_true",  help="Show progress")
    parser.add_argument("--replies",  action="store_true",  help="Fetch and categorize inbox replies")
    parser.add_argument("--single",   metavar="EMAIL",      help="Send to one email")
    parser.add_argument("--name",     default="",           help="Name for single send")
    parser.add_argument("--company",  default="",           help="Company for single send")
    parser.add_argument("--referred-by", default=None,      help="Referred by (for OOO redirects)")
    parser.add_argument("--draft",    metavar="EMAIL",      help="Save to drafts instead of sending")
    parser.add_argument("--note",     default=None,         help="Custom note to insert in body")
    args = parser.parse_args()

    if args.progress:
        check_progress()
    elif args.replies:
        fetch_replies()
    elif args.single:
        send_single(args.single, args.name, args.company, args.referred_by, args.note)
    elif args.draft:
        save_draft(args.draft, args.name, args.company)
    elif args.run:
        run_batch()
    else:
        parser.print_help()
