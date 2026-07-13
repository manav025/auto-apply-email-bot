#!/usr/bin/env python3
"""
Career Page Auto-Apply Bot
===========================

Opens each URL in career_urls.txt, finds the application form, and fills in
whatever it confidently can from your resume and profile.json. It does NOT
submit anything on its own -- it leaves the filled form open (or, with
--auto-submit, asks you to type SUBMIT for each one individually) so you can
check it over first.

Why it doesn't blind-submit by default:
  - Application forms often include legal/consequential questions (work
    authorization, sponsorship, salary, EEO/demographic self-identification)
    that this tool will never guess on your behalf.
  - Every ATS (Greenhouse, Lever, Workday, custom sites...) is laid out
    differently, so field-detection is heuristic and can occasionally
    mis-map a field. A human check before hitting submit is cheap insurance
    against applying with a mistake in it.
  - Some sites use CAPTCHAs. This tool does not attempt to solve or bypass
    them -- it will pause and tell you to finish that step yourself.

SETUP:
  pip install -r requirements.txt -r requirements-apply-bot.txt
  playwright install chromium
  cp profile.json.example profile.json   # then fill in your real info
  # add target job posting URLs to career_urls.txt, one per line

USAGE:
  python apply_bot.py                 # fill forms, screenshot, leave open for you to submit
  python apply_bot.py --auto-submit   # after you review each one, type SUBMIT to submit it
  python apply_bot.py --headless      # run without a visible browser window (still won't submit)
"""

import argparse
import json
import sys
import time
from pathlib import Path

from dotenv import load_dotenv
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

from ai_client import call_ai
from job_email_sender import extract_resume_text, find_resume

load_dotenv()

PROFILE_FILE = Path("profile.json")
URLS_FILE = Path("career_urls.txt")
SCREENSHOT_DIR = Path("screenshots")

SECONDS_BETWEEN_APPLICATIONS = 20

# Field name/label keywords -> profile.json key. Checked against the
# input's name, id, placeholder, and associated <label> text (lowercased).
FIELD_MAP = {
    "first_name": ["first name", "firstname", "fname", "given name"],
    "last_name": ["last name", "lastname", "lname", "surname", "family name"],
    "full_name": ["full name", "your name"],
    "email": ["email"],
    "phone": ["phone", "mobile", "telephone"],
    "linkedin_url": ["linkedin"],
    "portfolio_url": ["portfolio", "website", "personal site", "github"],
    "location": ["location", "city", "current city", "address"],
}

# Anything matching these keywords is left for the human -- never auto-filled,
# never guessed by the AI. This includes legal, financial, and demographic
# self-identification questions.
SENSITIVE_KEYWORDS = [
    "authorized to work", "work authorization", "sponsorship", "visa",
    "salary", "compensation expectation", "race", "ethnicity", "gender",
    "veteran", "disability", "military", "criminal", "felony",
    "social security", "ssn", "date of birth", "sexual orientation",
]

LONG_TEXT_KEYWORDS = [
    "cover letter", "why do you want", "why are you interested",
    "tell us about", "additional information", "anything else",
]


def load_profile() -> dict:
    if not PROFILE_FILE.exists():
        sys.exit(f"{PROFILE_FILE} not found. Copy profile.json.example to "
                  f"profile.json and fill in your real info.")
    return json.loads(PROFILE_FILE.read_text())


def load_urls() -> list[str]:
    if not URLS_FILE.exists():
        sys.exit(f"{URLS_FILE} not found. Add one job posting URL per line.")
    urls = [u.strip() for u in URLS_FILE.read_text().splitlines() if u.strip() and not u.startswith("#")]
    if not urls:
        sys.exit(f"{URLS_FILE} is empty.")
    return urls


def field_label_text(page, el) -> str:
    """Best-effort text describing what a form field is for."""
    parts = []
    for attr in ("name", "id", "placeholder", "aria-label"):
        v = el.get_attribute(attr)
        if v:
            parts.append(v)
    try:
        el_id = el.get_attribute("id")
        if el_id:
            label = page.query_selector(f'label[for="{el_id}"]')
            if label:
                parts.append(label.inner_text())
    except Exception:
        pass
    return " ".join(parts).lower()


def classify_field(label: str) -> str | None:
    for key, keywords in FIELD_MAP.items():
        if any(kw in label for kw in keywords):
            return key
    return None


def is_sensitive(label: str) -> bool:
    return any(kw in label for kw in SENSITIVE_KEYWORDS)


def is_long_text_prompt(label: str) -> bool:
    return any(kw in label for kw in LONG_TEXT_KEYWORDS)


def apply_to_job(page, url: str, profile: dict, resume_path: Path, resume_text: str, auto_submit: bool):
    print(f"\n{'=' * 70}\nOpening: {url}")
    page.goto(url, wait_until="domcontentloaded", timeout=30000)
    time.sleep(2)

    # Common "Apply" buttons that reveal a form on click
    for text in ["Apply Now", "Apply now", "Apply", "Apply for this job"]:
        btn = page.get_by_text(text, exact=False).first
        try:
            if btn and btn.is_visible():
                btn.click(timeout=3000)
                time.sleep(1.5)
                break
        except Exception:
            continue

    # CAPTCHA check -- do not attempt to bypass, just flag it
    if page.query_selector('iframe[src*="captcha"], iframe[title*="recaptcha" i], .g-recaptcha, #hcaptcha'):
        print("CAPTCHA detected on this page. Skipping auto-fill -- "
              "please complete this application manually.")
        return

    filled, skipped, ai_filled = [], [], []

    inputs = page.query_selector_all("input, textarea, select")
    job_desc_text = None

    for el in inputs:
        try:
            input_type = (el.get_attribute("type") or "text").lower()
            tag = el.evaluate("e => e.tagName.toLowerCase()")
        except Exception:
            continue

        if input_type in ("hidden", "submit", "button", "checkbox", "radio"):
            continue

        label = field_label_text(page, el)
        if not label:
            continue

        if is_sensitive(label):
            skipped.append(label)
            continue

        if input_type == "file" or "resume" in label or "cv" in label:
            try:
                el.set_input_files(str(resume_path))
                filled.append(f"[resume file] -> {label[:40]}")
            except Exception:
                skipped.append(f"(file upload failed) {label[:40]}")
            continue

        field_key = classify_field(label)
        if field_key and profile.get(field_key):
            try:
                el.fill(str(profile[field_key]))
                filled.append(f"{field_key} -> {label[:40]}")
            except Exception:
                skipped.append(label)
            continue

        if tag == "textarea" and is_long_text_prompt(label):
            if job_desc_text is None:
                job_desc_text = page.inner_text("body")[:3000]
            prompt = f"""Write a concise, specific answer (80-150 words, no markdown) for a job \
application question. The question/field is: "{label}"

Candidate resume:
---
{resume_text[:2500]}
---

Job posting page text (for context on the role):
---
{job_desc_text}
---

Extra notes from the candidate to weave in if relevant: {profile.get("cover_letter_notes", "")}

Only use real facts from the resume. Do not invent experience.
"""
            try:
                answer = call_ai(prompt, max_tokens=400)
                el.fill(answer)
                ai_filled.append(label[:60])
            except Exception as e:
                skipped.append(f"(AI draft failed: {e}) {label[:40]}")
            continue

        # Anything else we don't recognize -- leave for the human
        skipped.append(label[:60])

    SCREENSHOT_DIR.mkdir(exist_ok=True)
    shot_path = SCREENSHOT_DIR / f"{int(time.time())}.png"
    page.screenshot(path=str(shot_path), full_page=True)

    print(f"Filled automatically: {len(filled)}")
    for f in filled:
        print(f"  - {f}")
    print(f"AI-drafted answers (review these!): {len(ai_filled)}")
    for f in ai_filled:
        print(f"  - {f}")
    print(f"Left for you to fill in manually: {len(skipped)}")
    for s in skipped:
        print(f"  - {s}")
    print(f"Screenshot saved: {shot_path}")

    if not auto_submit:
        print("Not submitting (default mode). Review the open browser tab "
              "and submit it yourself when ready.")
        return

    confirm = input(f"\nReview the browser window for {url}. "
                     f"Type SUBMIT to submit this application, or anything else to skip: ")
    if confirm.strip() != "SUBMIT":
        print("Skipped -- not submitted.")
        return

    for text in ["Submit Application", "Submit application", "Submit", "Send Application"]:
        btn = page.get_by_text(text, exact=False).first
        try:
            if btn and btn.is_visible():
                btn.click(timeout=3000)
                print("Submitted.")
                return
        except Exception:
            continue
    print("Could not find a submit button automatically -- please submit manually.")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--resume", help="Path to resume (pdf/docx/txt)")
    parser.add_argument("--auto-submit", action="store_true",
                         help="After you review each filled form, type SUBMIT to submit it")
    parser.add_argument("--headless", action="store_true",
                         help="Run without a visible browser window")
    args = parser.parse_args()

    profile = load_profile()
    urls = load_urls()
    resume_path = find_resume(args.resume)
    resume_text = extract_resume_text(resume_path)

    print(f"Loaded profile, {len(urls)} job URL(s), resume: {resume_path}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=args.headless)
        context = browser.new_context()
        for i, url in enumerate(urls):
            page = context.new_page()
            try:
                apply_to_job(page, url, profile, resume_path, resume_text, args.auto_submit)
            except PWTimeout:
                print(f"Timed out loading {url} -- skipping.")
            except Exception as e:
                print(f"Error on {url}: {e}")
            if not args.auto_submit:
                pass  # leave tab open for manual review
            if i < len(urls) - 1:
                time.sleep(SECONDS_BETWEEN_APPLICATIONS)

        if not args.auto_submit:
            input("\nAll forms filled. Browser windows are left open for you to review "
                  "and submit manually. Press Enter here when you're done to close them...")
        browser.close()


if __name__ == "__main__":
    main()
