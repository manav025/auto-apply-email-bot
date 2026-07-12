#!/usr/bin/env python3
"""
Job Application Email Sender
=============================

Reads your resume, asks Claude to (1) suggest improvements and
(2) draft a personalized outreach email for each contact at a company,
then sends the email from your Gmail account with your resume attached.

SETUP (see README.md for full details):
  1. pip install -r requirements.txt
  2. Copy .env.example to .env and fill in:
       GMAIL_ADDRESS, GMAIL_APP_PASSWORD, ANTHROPIC_API_KEY
  3. Fill in recipients.csv with the people you want to email
  4. Put your resume at resume.pdf or resume.docx (or pass --resume path)

USAGE:
  # Step 1: get resume feedback only (no emails touched)
  python job_email_sender.py --analyze-resume

  # Step 2: preview the emails it would send (nothing is sent)
  python job_email_sender.py --dry-run

  # Step 3: actually send, after you've reviewed the previews
  python job_email_sender.py --send
"""

import argparse
import csv
import os
import smtplib
import ssl
import sys
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from pathlib import Path
from datetime import datetime

from dotenv import load_dotenv
from ai_client import call_ai

load_dotenv()

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------

GMAIL_ADDRESS = os.getenv("GMAIL_ADDRESS")
GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
GROQ_API_KEY = os.getenv("GROQ_API_KEY")

SENDER_NAME = os.getenv("SENDER_NAME", "")
TARGET_ROLE = os.getenv("TARGET_ROLE", "")   # e.g. "Software Engineer"
TARGET_COMPANY = os.getenv("TARGET_COMPANY", "")  # e.g. "Acme Corp"

SECONDS_BETWEEN_EMAILS = int(os.getenv("SECONDS_BETWEEN_EMAILS", "45"))

RECIPIENTS_CSV = Path("recipients.csv")
SUGGESTIONS_FILE = Path("resume_suggestions.md")
SENT_LOG = Path("sent_log.csv")


# ---------------------------------------------------------------------------
# Resume reading
# ---------------------------------------------------------------------------

def extract_resume_text(path: Path) -> str:
    suffix = path.suffix.lower()
    if suffix == ".pdf":
        import pdfplumber
        text = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                text.append(page.extract_text() or "")
        return "\n".join(text)
    elif suffix == ".docx":
        import docx
        doc = docx.Document(path)
        return "\n".join(p.text for p in doc.paragraphs)
    elif suffix == ".txt":
        return path.read_text()
    else:
        raise ValueError(f"Unsupported resume format: {suffix}")


def find_resume(explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            sys.exit(f"Resume not found at {p}")
        return p
    for candidate in ("resume.pdf", "resume.docx", "resume.txt"):
        p = Path(candidate)
        if p.exists():
            return p
    sys.exit("No resume found. Put resume.pdf / resume.docx in this folder, "
              "or pass --resume path/to/file")


# ---------------------------------------------------------------------------
# Claude calls
# ---------------------------------------------------------------------------

def get_resume_suggestions(resume_text: str) -> str:
    prompt = f"""You are an expert technical resume reviewer and career coach.

Here is a candidate's resume:
---
{resume_text}
---

Give specific, actionable suggestions to improve this resume for {TARGET_ROLE or "a competitive role"} \
positions{f" at companies like {TARGET_COMPANY}" if TARGET_COMPANY else ""}. Cover:
1. Overall structure / formatting issues
2. Weak or vague bullet points, with a rewritten stronger version for each
3. Missing quantification / impact metrics
4. Keywords likely missing for ATS screening
5. Anything that should be cut

Be concrete. Do not invent facts about the candidate's experience -- only rewrite/tighten what's there.
"""
    return call_ai(prompt, max_tokens=2000)


def draft_email(resume_text: str, recipient: dict) -> tuple[str, str]:
    """Returns (subject, body) for one recipient."""
    name = recipient.get("name", "").strip()
    role = recipient.get("role", "").strip()
    company = recipient.get("company", "").strip() or TARGET_COMPANY
    job_title = recipient.get("job_title", "").strip() or TARGET_ROLE

    prompt = f"""Write a concise, professional job-outreach email from a candidate to a contact at a company.

Candidate resume (for context on real background -- do not fabricate anything not in here):
---
{resume_text}
---

Recipient: {name or "Hiring contact"}{f", {role}" if role else ""} at {company or "the company"}
Target position: {job_title or "a relevant open role"}
Sender's name to sign as: {SENDER_NAME or "[Your Name]"}

Requirements:
- Subject line: short, specific, not spammy
- Body: 120-180 words, warm but professional, no generic filler ("I am writing to express...")
- Reference 1-2 concrete, real details from the resume that make the candidate relevant
- Mention the resume is attached
- End with a clear, low-pressure call to action
- No markdown, no placeholders left unfilled except sign-off name if not provided
- Output format exactly:
SUBJECT: <subject line>
BODY:
<email body>
"""
    text = call_ai(prompt, max_tokens=600)

    subject, body = "", text
    if "SUBJECT:" in text and "BODY:" in text:
        subject = text.split("SUBJECT:", 1)[1].split("BODY:", 1)[0].strip()
        body = text.split("BODY:", 1)[1].strip()
    return subject, body


# ---------------------------------------------------------------------------
# Email sending
# ---------------------------------------------------------------------------

def send_email(to_addr: str, subject: str, body: str, resume_path: Path):
    msg = MIMEMultipart()
    msg["From"] = f"{SENDER_NAME} <{GMAIL_ADDRESS}>" if SENDER_NAME else GMAIL_ADDRESS
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    with open(resume_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=resume_path.name)
    part["Content-Disposition"] = f'attachment; filename="{resume_path.name}"'
    msg.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(GMAIL_ADDRESS, GMAIL_APP_PASSWORD)
        server.sendmail(GMAIL_ADDRESS, to_addr, msg.as_string())


def log_sent(to_addr: str, subject: str):
    is_new = not SENT_LOG.exists()
    with open(SENT_LOG, "a", newline="") as f:
        writer = csv.writer(f)
        if is_new:
            writer.writerow(["timestamp", "to", "subject"])
        writer.writerow([datetime.now().isoformat(timespec="seconds"), to_addr, subject])


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def load_recipients() -> list[dict]:
    if not RECIPIENTS_CSV.exists():
        sys.exit(f"{RECIPIENTS_CSV} not found. Create it with columns: "
                  f"email,name,role,company,job_title")
    with open(RECIPIENTS_CSV, newline="") as f:
        rows = list(csv.DictReader(f))
    if not rows:
        sys.exit(f"{RECIPIENTS_CSV} is empty.")
    for r in rows:
        if not r.get("email", "").strip():
            sys.exit(f"Row missing email: {r}")
    return rows


def check_config(need_send: bool):
    if not GROQ_API_KEY:
        sys.exit("Missing GROQ_API_KEY in .env. Get a free key at "
                  "https://console.groq.com/keys")
    if need_send and not (GMAIL_ADDRESS and GMAIL_APP_PASSWORD):
        sys.exit("Missing GMAIL_ADDRESS / GMAIL_APP_PASSWORD in .env")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--resume", help="Path to resume (pdf/docx/txt)")
    parser.add_argument("--analyze-resume", action="store_true",
                         help="Only generate resume improvement suggestions, then exit")
    parser.add_argument("--dry-run", action="store_true",
                         help="Generate and print/save all emails without sending anything")
    parser.add_argument("--send", action="store_true",
                         help="Actually send the emails (after you've reviewed a dry run)")
    args = parser.parse_args()

    if not (args.analyze_resume or args.dry_run or args.send):
        parser.print_help()
        sys.exit("\nPick one of --analyze-resume, --dry-run, or --send")

    check_config(need_send=args.send)

    resume_path = find_resume(args.resume)
    print(f"Reading resume: {resume_path}")
    resume_text = extract_resume_text(resume_path)
    if len(resume_text.strip()) < 50:
        sys.exit("Couldn't extract meaningful text from the resume file. "
                  "Try a different file or --resume path.")

    if args.analyze_resume:
        print("Asking Claude for resume feedback...")
        suggestions = get_resume_suggestions(resume_text)
        SUGGESTIONS_FILE.write_text(suggestions)
        print(f"\nSaved suggestions to {SUGGESTIONS_FILE}\n")
        print(suggestions)
        return

    recipients = load_recipients()
    print(f"Loaded {len(recipients)} recipient(s) from {RECIPIENTS_CSV}")

    drafts = []
    for r in recipients:
        print(f"Drafting email for {r['email']}...")
        subject, body = draft_email(resume_text, r)
        drafts.append((r["email"], subject, body))

    print("\n" + "=" * 70)
    for to_addr, subject, body in drafts:
        print(f"TO: {to_addr}")
        print(f"SUBJECT: {subject}")
        print("-" * 70)
        print(body)
        print("=" * 70)

    if args.dry_run:
        print(f"\nDry run only -- nothing was sent. {len(drafts)} email(s) previewed above.")
        return

    if args.send:
        confirm = input(f"\nType SEND to actually email these {len(drafts)} people now: ")
        if confirm.strip() != "SEND":
            print("Aborted. Nothing was sent.")
            return
        for i, (to_addr, subject, body) in enumerate(drafts):
            try:
                send_email(to_addr, subject, body, resume_path)
                log_sent(to_addr, subject)
                print(f"Sent to {to_addr}")
            except Exception as e:
                print(f"FAILED to send to {to_addr}: {e}")
            if i < len(drafts) - 1:
                time.sleep(SECONDS_BETWEEN_EMAILS)
        print(f"\nDone. See {SENT_LOG} for the send log.")


if __name__ == "__main__":
    main()
