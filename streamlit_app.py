"""
Job Application Email Assistant -- Streamlit web app version.

Same functionality as job_email_sender.py, but as a browser UI you can run
locally (`streamlit run streamlit_app.py`) or host for free on Streamlit
Community Cloud, so nothing runs on your own laptop.

Credentials are entered in the sidebar and kept only in this browser
session -- they are never written to disk. If you deploy this publicly,
set APP_PASSWORD in Streamlit's secrets so random visitors can't use your
Gmail/Groq credentials.
"""

import io
import re
import time

import pandas as pd
import pdfplumber
import docx
import requests
import streamlit as st
from groq import Groq
import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

st.set_page_config(page_title="CV Analyzer & Job Finder", page_icon="📄", layout="wide")

st.markdown("""
<style>
.stApp { background: linear-gradient(180deg, #f7f6fd 0%, #efedfb 100%); }
.score-card { background:white; border-radius:20px; padding:28px 24px; box-shadow:0 2px 14px rgba(76,63,176,0.08); }
.gauge-wrap { display:flex; flex-direction:column; align-items:center; padding-top:8px; }
.chip { display:inline-block; padding:6px 14px; border-radius:999px; font-weight:600; font-size:0.85rem; margin:4px 6px 4px 0; }
.chip-green { background:#d1fae5; color:#047857; }
.chip-red { background:#fee2e2; color:#b91c1c; }
.bar-row { margin-bottom:16px; }
.bar-label { display:flex; justify-content:space-between; font-weight:600; color:#1e1b4b; margin-bottom:5px; }
.bar-track { background:#ece9f9; border-radius:8px; height:9px; overflow:hidden; }
.bar-fill { height:100%; border-radius:8px; }
.job-card { background:white; border-radius:16px; padding:18px 20px; box-shadow:0 1px 8px rgba(76,63,176,0.08); margin-bottom:14px; }
.badge { display:inline-block; padding:3px 11px; border-radius:999px; font-size:0.72rem; font-weight:700; background:#ede9fe; color:#5b21b6; margin-left:6px; vertical-align:middle; }
.badge-ok { background:#d1fae5; color:#047857; }
.badge-warn { background:#fef3c7; color:#92400e; }
.section-title { font-size:1.4rem; font-weight:800; color:#1e1b4b; margin:0.2em 0 0.4em 0; }
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Optional password gate -- set APP_PASSWORD in Streamlit secrets to enable.
# Without it, anyone with the app URL can open it (they'd still need their
# own Groq/Gmail credentials to actually do anything, since nothing is
# hardcoded server-side).
# ---------------------------------------------------------------------------
_app_password = st.secrets.get("APP_PASSWORD", "")
if _app_password:
    if not st.session_state.get("unlocked"):
        st.title("🔒 Job Email Assistant")
        pw = st.text_input("App password", type="password")
        if st.button("Unlock"):
            if pw == _app_password:
                st.session_state.unlocked = True
                st.rerun()
            else:
                st.error("Wrong password")
        st.stop()

st.title("📄 CV Analyzer & Job Finder")
st.caption("ATS scoring, JD match, AI-tailored resume, and outreach emails -- sent from your own Gmail, "
           "using your own free Groq key. Nothing here runs on your machine.")

# ---------------------------------------------------------------------------
# Sidebar: credentials (session-only, never saved to disk)
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Your credentials")
    st.caption("Kept only for this browser session.")
    groq_key = st.text_input("Groq API key", type="password",
                              value=st.secrets.get("GROQ_API_KEY", ""),
                              help="Free, no card required: console.groq.com/keys")
    groq_model = st.text_input("Groq model", value=st.secrets.get("GROQ_MODEL", "llama-3.3-70b-versatile"))
    gmail_address = st.text_input("Gmail address", value=st.secrets.get("GMAIL_ADDRESS", ""))
    gmail_app_password = st.text_input("Gmail app password", type="password",
                                        value=st.secrets.get("GMAIL_APP_PASSWORD", ""),
                                        help="Not your normal password -- create one at "
                                             "myaccount.google.com/apppasswords")
    sender_name = st.text_input("Your name (email sign-off)", value=st.secrets.get("SENDER_NAME", ""))
    seconds_between = st.slider("Seconds to wait between sends", 5, 90, 30)
    st.markdown("---")
    st.caption("Optional, powers 'Find open roles' below")
    adzuna_app_id = st.text_input("Adzuna App ID", value=st.secrets.get("ADZUNA_APP_ID", ""))
    adzuna_app_key = st.text_input("Adzuna App Key", type="password", value=st.secrets.get("ADZUNA_APP_KEY", ""))
    st.markdown("---")
    st.markdown("[Free Groq key](https://console.groq.com/keys) · "
                "[Gmail app passwords](https://myaccount.google.com/apppasswords) · "
                "[Free Adzuna API key](https://developer.adzuna.com/)")


def call_ai(prompt: str, max_tokens: int = 1000) -> str:
    if not groq_key:
        raise RuntimeError("Enter your Groq API key in the sidebar first.")
    client = Groq(api_key=groq_key)
    resp = client.chat.completions.create(
        model=groq_model,
        max_tokens=max_tokens,
        temperature=0.6,
        messages=[{"role": "user", "content": prompt}],
    )
    return resp.choices[0].message.content.strip()


def extract_resume_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    data = uploaded_file.getvalue()
    if name.endswith(".pdf"):
        text = []
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            for page in pdf.pages:
                text.append(page.extract_text() or "")
        return "\n".join(text)
    elif name.endswith(".docx"):
        d = docx.Document(io.BytesIO(data))
        return "\n".join(p.text for p in d.paragraphs)
    elif name.endswith(".txt"):
        return data.decode("utf-8", errors="ignore")
    else:
        raise ValueError("Unsupported resume format")


def send_email(to_addr: str, subject: str, body: str, resume_bytes: bytes, resume_name: str):
    msg = MIMEMultipart()
    msg["From"] = f"{sender_name} <{gmail_address}>" if sender_name else gmail_address
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))

    part = MIMEApplication(resume_bytes, Name=resume_name)
    part["Content-Disposition"] = f'attachment; filename="{resume_name}"'
    msg.attach(part)

    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(gmail_address, gmail_app_password)
        server.sendmail(gmail_address, to_addr, msg.as_string())


def name_from_email(email: str) -> str:
    """Best-effort guess at a person's name from their email address, e.g.
    'jane.doe@acme.com' -> 'Jane Doe', 'jdoe123@acme.com' -> 'Jdoe123'.
    Used only as a fallback when no name was supplied."""
    if not email or "@" not in email:
        return ""
    local = email.split("@")[0]
    local = re.sub(r"\d+", "", local)  # drop trailing numbers like jdoe123
    parts = re.split(r"[._\-+]", local)
    parts = [p for p in parts if p]
    if not parts:
        return ""
    return " ".join(p.capitalize() for p in parts)


def text_to_docx_bytes(title: str, body_text: str) -> bytes:
    """Wrap plain text into a simple, clean .docx (not a reproduction of the
    original resume's visual design -- just a readable, ATS-friendly draft)."""
    d = docx.Document()
    d.add_heading(title, level=1)
    for line in body_text.split("\n"):
        line = line.strip()
        if not line:
            d.add_paragraph("")
        elif line.endswith(":") and len(line) < 60:
            d.add_heading(line.rstrip(":"), level=2)
        elif line.startswith(("- ", "• ", "* ")):
            d.add_paragraph(line[2:].strip(), style="List Bullet")
        else:
            d.add_paragraph(line)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# ATS scoring engine -- part rule-based (deterministic, no AI needed), part
# AI-judged (content quality, keyword extraction). Categories sum to /100.
# ---------------------------------------------------------------------------
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+")
PHONE_RE = re.compile(r"(\+?\d[\d\-\s()]{7,}\d)")
URL_RE = re.compile(r"(https?://|www\.)\S+|\b\S+\.(?:com|io|dev|me|org|net)/\S*", re.I)
SECTION_PATTERNS = {
    "experience": re.compile(r"\b(experience|employment history|work experience)\b", re.I),
    "education": re.compile(r"\beducation\b", re.I),
    "skills": re.compile(r"\bskills\b", re.I),
    "summary": re.compile(r"\b(summary|objective|profile)\b", re.I),
}


def score_contact_info(text: str):
    found, score = [], 0
    if EMAIL_RE.search(text):
        score += 5; found.append("email")
    if PHONE_RE.search(text):
        score += 5; found.append("phone")
    if URL_RE.search(text):
        score += 5; found.append("link")
    return min(score, 15), found


def score_structure(text: str):
    found = [name for name, pat in SECTION_PATTERNS.items() if pat.search(text)]
    missing = [name for name in SECTION_PATTERNS if name not in found]
    score = round(len(found) / len(SECTION_PATTERNS) * 25)
    return min(score, 25), found, missing


def score_formatting(text: str):
    word_count = len(text.split())
    lines = [l for l in text.split("\n") if l.strip()]
    bullet_lines = [l for l in lines if l.strip().startswith(("-", "•", "*"))]
    bullet_ratio = (len(bullet_lines) / len(lines)) if lines else 0
    wc_score = 12 if 400 <= word_count <= 700 else (8 if 300 <= word_count <= 900 else 4)
    bullet_score = 8 if bullet_ratio >= 0.15 else (4 if bullet_ratio > 0 else 0)
    return min(wc_score + bullet_score, 20), word_count, bullet_ratio


def parse_json_loose(raw: str):
    import json
    text = re.sub(r"^```(json)?", "", raw.strip(), flags=re.I).strip()
    text = re.sub(r"```$", "", text).strip()
    try:
        return json.loads(text)
    except Exception:
        m = re.search(r"[\{\[].*[\}\]]", text, re.S)
        if m:
            try:
                return json.loads(m.group(0))
            except Exception:
                pass
        return {}


def analyze_resume_ai(resume_text: str, target_role: str, target_company: str, jd_text: str = ""):
    context = (f"Job description:\n---\n{jd_text}\n---" if jd_text.strip()
               else f"Target role: {target_role or 'a competitive role'}"
                    f"{f' at companies like {target_company}' if target_company else ''}")
    prompt = f"""You are an expert ATS resume reviewer. Analyze the resume below and respond with ONLY valid JSON \
(no markdown fences, no commentary) matching exactly this schema:

{{
  "content_quality_score": <integer 0-30>,
  "technical_keywords_found": [<specific technical/domain keywords or skills actually present in the resume, up to 20>],
  "issues": [<up to 5 short, specific problems, each under 12 words>],
  "suggestions": [<up to 6 short, specific actionable suggestions, each under 20 words>]
}}

Resume:
---
{resume_text}
---

{context}

Score content_quality_score out of 30 based on: strong action verbs, quantified achievements/impact, relevance \
and specificity (not generic filler). Do not invent facts about the candidate not present in the resume.
"""
    return parse_json_loose(call_ai(prompt, max_tokens=1200))


def run_ats_analysis(resume_text: str, target_role: str, target_company: str, jd_text: str = ""):
    contact_score, contact_found = score_contact_info(resume_text)
    structure_score, sections_found, sections_missing = score_structure(resume_text)
    formatting_score, word_count, bullet_ratio = score_formatting(resume_text)
    ai = analyze_resume_ai(resume_text, target_role, target_company, jd_text)

    content_score = max(0, min(30, int(ai.get("content_quality_score", 15) or 15)))
    keywords_found = [k for k in (ai.get("technical_keywords_found") or []) if isinstance(k, str)]
    keywords_score = min(10, round(len(keywords_found) / 15 * 10))
    overall = contact_score + structure_score + content_score + formatting_score + keywords_score

    issues = [str(i) for i in (ai.get("issues") or [])]
    for name in sections_missing:
        issues.append(f"'{name.capitalize()}' section not clearly labeled")
    if "email" not in contact_found:
        issues.append("No email address found")
    if "phone" not in contact_found:
        issues.append("No phone number found")

    suggestions = [str(s) for s in (ai.get("suggestions") or [])]
    if word_count < 400:
        suggestions.append(f"CV is {word_count} words -- aim for 400-700 words for a stronger ATS pass rate")
    elif word_count > 900:
        suggestions.append(f"CV is {word_count} words -- consider trimming toward 400-700")
    suggestions.append(f"{len(keywords_found)} technical keywords found -- try to reach 15+ for a strong ATS pass rate")

    return {
        "overall": overall,
        "categories": {
            "Contact Info": (contact_score, 15),
            "Structure": (structure_score, 25),
            "Content Quality": (content_score, 30),
            "Formatting": (formatting_score, 20),
            "Keywords": (keywords_score, 10),
        },
        "issues": issues[:8],
        "suggestions": suggestions[:8],
    }


def analyze_jd_match(resume_text: str, jd_text: str):
    prompt = f"""Extract the 10-15 most important, specific skills/qualifications/keywords from this job \
description. Respond with ONLY a JSON array of short keyword strings (1-3 words each), no commentary, e.g. \
["python","stakeholder management","aws"].

Job description:
---
{jd_text}
---
"""
    keywords = parse_json_loose(call_ai(prompt, max_tokens=400))
    if not isinstance(keywords, list):
        keywords = []
    resume_lower = resume_text.lower()
    matched = [k for k in keywords if isinstance(k, str) and k.lower().strip() in resume_lower]
    missing = [k for k in keywords if isinstance(k, str) and k.lower().strip() not in resume_lower]
    total = len(matched) + len(missing)
    pct = round(len(matched) / total * 100) if total else 0

    notes = []
    if missing:
        notes.append(f"Add these missing keywords to your CV: {', '.join(missing[:6])}")
    if total and len(matched) / total < 0.5:
        notes.append("Your CV matches less than half the JD keywords -- consider tailoring the skills section directly")
    yoe_match = re.search(r"(\d+)\+?\s*years?", jd_text, re.I)
    if yoe_match:
        notes.append(f"JD requires {yoe_match.group(1)}+ years experience -- make sure your experience duration is clearly stated")

    return {"pct": pct, "matched": matched, "missing": missing, "notes": notes}


def render_gauge(score: int, label: str = "ATS SCORE") -> str:
    score = max(0, min(100, int(score)))
    if score >= 80:
        color, tag, badge_cls = "#10b981", "Strong", "badge-ok"
    elif score >= 60:
        color, tag, badge_cls = "#e2a03f", "Average", "badge-warn"
    elif score >= 40:
        color, tag, badge_cls = "#f97316", "Needs Work", "badge-warn"
    else:
        color, tag, badge_cls = "#ef4444", "Weak", "badge-warn"
    deg = score * 3.6
    return f"""
    <div class="gauge-wrap">
      <div style="width:170px;height:170px;border-radius:50%;
           background:conic-gradient({color} {deg}deg, #ece9f9 0deg);
           display:flex;align-items:center;justify-content:center;">
        <div style="width:132px;height:132px;background:white;border-radius:50%;
             display:flex;flex-direction:column;align-items:center;justify-content:center;">
          <span style="font-size:2.4rem;font-weight:800;color:#1e1b4b;line-height:1;">{score}</span>
          <span style="color:#9ca3af;font-size:0.85rem;">/100</span>
        </div>
      </div>
      <div style="margin-top:10px;font-weight:700;color:#6b7280;letter-spacing:0.05em;font-size:0.78rem;">{label}</div>
      <span class="badge {badge_cls}" style="margin-top:6px;">{tag}</span>
    </div>
    """


def render_bar(label: str, score: int, max_score: int, color: str) -> str:
    pct = (score / max_score * 100) if max_score else 0
    return f"""
    <div class="bar-row">
      <div class="bar-label"><span>{label}</span><span style="color:{color};">{score}/{max_score}</span></div>
      <div class="bar-track"><div class="bar-fill" style="width:{pct}%;background:{color};"></div></div>
    </div>
    """


def render_chip(text: str, kind: str = "green") -> str:
    return f'<span class="chip {"chip-green" if kind == "green" else "chip-red"}">{text}</span>'


BAR_COLORS = ["#7c6ff0", "#10b981", "#e2a03f", "#3b82f6", "#ef4444"]


# ---------------------------------------------------------------------------
# 1. Resume
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# 1. Resume + ATS score
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">1. Your CV, objectively scored</div>', unsafe_allow_html=True)
st.caption("Upload your CV to get an ATS score, issues found, and suggestions -- powered by a mix of rule-based "
           "checks and AI review.")

resume_file = st.file_uploader("Upload your resume", type=["pdf", "docx", "txt"])
resume_text = None
if resume_file:
    try:
        resume_text = extract_resume_text(resume_file)
        st.success(f"Loaded {resume_file.name} ({len(resume_text)} characters extracted)")
    except Exception as e:
        st.error(f"Couldn't read that file: {e}")

col1, col2 = st.columns(2)
target_role = col1.text_input("Target role", value=st.secrets.get("TARGET_ROLE", ""))
target_company = col2.text_input("Default target company (optional -- can override per contact below)",
                                  value=st.secrets.get("TARGET_COMPANY", ""))

if st.button("Analyze CV", type="primary", disabled=not (resume_text and groq_key)):
    with st.spinner("Scoring your resume..."):
        st.session_state.ats = run_ats_analysis(resume_text, target_role, target_company)

if st.session_state.get("ats"):
    ats = st.session_state.ats
    gcol, scol = st.columns([1, 2])
    with gcol:
        st.markdown(f'<div class="score-card">{render_gauge(ats["overall"])}</div>', unsafe_allow_html=True)
    with scol:
        bars_html = "".join(
            render_bar(name, s, m, BAR_COLORS[i % len(BAR_COLORS)])
            for i, (name, (s, m)) in enumerate(ats["categories"].items())
        )
        st.markdown(f'<div class="score-card">{bars_html}</div>', unsafe_allow_html=True)

    icol, scol2 = st.columns(2)
    with icol:
        issues_html = "".join(f"<div style='margin-bottom:8px;'>❌ {i}</div>" for i in ats["issues"]) or "<div>No major issues found.</div>"
        st.markdown(f'<div class="score-card"><b style="color:#b91c1c;">⚠ ISSUES FOUND</b><div style="margin-top:12px;">{issues_html}</div></div>', unsafe_allow_html=True)
    with scol2:
        sugg_html = "".join(f"<div style='margin-bottom:8px;'>✅ {s}</div>" for s in ats["suggestions"]) or "<div>Looking solid.</div>"
        st.markdown(f'<div class="score-card"><b style="color:#5b21b6;">⚡ SUGGESTIONS</b><div style="margin-top:12px;">{sugg_html}</div></div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">2. Match to a Job Description</div>', unsafe_allow_html=True)
jd_text = st.text_area(
    "Paste any job posting -- see exactly how well your CV fits",
    height=160,
    help="Also powers resume tailoring and the 'why you're a fit' pitch below.",
)

if st.button("Analyze Match", type="primary", disabled=not (resume_text and jd_text.strip() and groq_key)):
    with st.spinner("Comparing your CV to this JD..."):
        st.session_state.jd_match = analyze_jd_match(resume_text, jd_text)

if st.session_state.get("jd_match"):
    jm = st.session_state.jd_match
    pct = jm["pct"]
    badge_cls = "badge-ok" if pct >= 60 else "badge-warn"
    st.markdown(
        f'<div class="score-card"><span style="font-weight:800;font-size:1.1rem;">Match score</span>'
        f'<span class="badge {badge_cls}" style="float:right;font-size:0.95rem;padding:6px 16px;">{pct}% match</span>'
        f'<div style="clear:both;"></div></div>',
        unsafe_allow_html=True,
    )
    hcol, mcol = st.columns(2)
    with hcol:
        chips = "".join(render_chip(k, "green") for k in jm["matched"]) or "<i>None matched</i>"
        st.markdown(f'<div class="score-card"><b>YOUR CV HAS</b><div style="margin-top:10px;">{chips}</div></div>', unsafe_allow_html=True)
    with mcol:
        chips = "".join(render_chip(k, "red") for k in jm["missing"]) or "<i>Nothing missing</i>"
        st.markdown(f'<div class="score-card"><b>MISSING FROM CV</b><div style="margin-top:10px;">{chips}</div></div>', unsafe_allow_html=True)
    if jm["notes"]:
        notes_html = "".join(f"<div style='margin-bottom:8px;'>✅ {n}</div>" for n in jm["notes"])
        st.markdown(f'<div class="score-card" style="margin-top:14px;">{notes_html}</div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">3. AI resume tailoring & pitch</div>', unsafe_allow_html=True)
tcol1, tcol2 = st.columns(2)

with tcol1:
    if st.button("Tailor resume to this JD", disabled=not (resume_text and jd_text.strip() and groq_key)):
        with st.spinner("Tailoring resume..."):
            prompt = f"""You are rewriting a candidate's resume to better match a specific job description.

Original resume:
---
{resume_text}
---

Job description:
---
{jd_text}
---

Rules -- follow these strictly:
- Only reword, reorder, re-emphasize, or tighten content that is already in the original resume.
- NEVER invent new skills, tools, employers, titles, dates, metrics, or achievements not already present or \
clearly implied by the original.
- Prioritize and surface the experience most relevant to this JD nearer the top of each section.
- Naturally work in JD keywords/phrasing only where they genuinely match real content.
- Keep the same overall sections (e.g. Experience, Education, Skills) the original has.
- Output the full tailored resume as plain text, ready to review -- no commentary, no markdown formatting, \
just the resume content with clear section headers.
"""
            st.session_state.tailored_resume = call_ai(prompt, max_tokens=2500)

with tcol2:
    if st.button("Generate 'why you're a fit' pitch", disabled=not (resume_text and groq_key)):
        with st.spinner("Writing pitch..."):
            prompt = f"""Write exactly 3 short sentences (not bullet points) explaining why this candidate is a \
strong fit for this specific role. This will be used as an opening pitch in outreach emails or pasted into an \
application's free-text "why are you interested" field.

Candidate resume:
---
{resume_text}
---

{"Job description:\n---\n" + jd_text + "\n---" if jd_text.strip() else f"Target role: {target_role or 'a relevant role'}"}

Only reference real, specific experience from the resume -- no generic claims like "I am a hard worker." \
Output only the 3 sentences, nothing else.
"""
            st.session_state.fit_pitch = call_ai(prompt, max_tokens=250)

if st.session_state.get("tailored_resume"):
    with st.expander("Tailored resume draft", expanded=True):
        st.warning("This is a content-only draft (plain formatting) for you to review -- it does not preserve "
                   "your original resume's visual design. Copy what's useful back into your real resume, or "
                   "download this version below.")
        st.text_area("Tailored resume text", value=st.session_state.tailored_resume, height=300, key="tailored_view")
        st.download_button(
            "Download as .docx",
            data=text_to_docx_bytes("Tailored Resume", st.session_state.tailored_resume),
            file_name="tailored_resume.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

if st.session_state.get("fit_pitch"):
    with st.expander("Why you're a fit (3 lines)", expanded=True):
        st.text_area("Copy this into emails or application forms", value=st.session_state.fit_pitch,
                      height=100, key="fit_pitch_view")

# ---------------------------------------------------------------------------
# 4. Find open roles (legitimate public job API -- not scraping LinkedIn/Indeed/Naukri)
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">4. Find open roles</div>', unsafe_allow_html=True)
st.caption("Searches via Adzuna's public API (aggregates many boards legitimately, official API -- no scraping). "
           "Add a free key in the sidebar to enable this.")

if not (adzuna_app_id and adzuna_app_key):
    st.info("Add your free Adzuna App ID/Key in the sidebar to search here -- "
            "https://developer.adzuna.com/")
else:
    scol1, scol2, scol3 = st.columns([2, 1, 1])
    search_keyword = scol1.text_input("Keyword", value=target_role, key="job_search_kw")
    search_country = scol2.selectbox("Country", ["us", "gb", "in", "ca", "au", "de", "fr"], index=0)
    search_page = scol3.number_input("Page", min_value=1, value=1)

    if st.button("Search open roles", type="primary"):
        with st.spinner("Searching..."):
            try:
                resp = requests.get(
                    f"https://api.adzuna.com/v1/api/jobs/{search_country}/search/{int(search_page)}",
                    params={
                        "app_id": adzuna_app_id,
                        "app_key": adzuna_app_key,
                        "what": search_keyword,
                        "content-type": "application/json",
                    },
                    timeout=15,
                )
                resp.raise_for_status()
                st.session_state.job_results = resp.json().get("results", [])
            except Exception as e:
                st.error(f"Search failed: {e}")
                st.session_state.job_results = []

    MANUAL_ONLY_DOMAINS = ["linkedin.com", "indeed.com", "naukri.com", "glassdoor.com"]

    results = st.session_state.get("job_results", [])
    if results:
        st.session_state.setdefault("saved_jobs", set())
        fcol1, fcol2 = st.columns([3, 1])
        fcol1.markdown(f"**Job Listings** &nbsp; <span style='color:#6b7280;'>{len(results)} results</span>",
                        unsafe_allow_html=True)
        show_saved_only = fcol2.checkbox("Saved only")

        display_results = [
            j for j in results
            if not show_saved_only or j.get("id") in st.session_state.saved_jobs
        ]

        cols = st.columns(3)
        for idx, job in enumerate(display_results):
            url = job.get("redirect_url", "")
            company = (job.get("company") or {}).get("display_name", "Unknown company")
            title = job.get("title", "Untitled role")
            location = (job.get("location") or {}).get("display_name", "")
            job_id = job.get("id")
            manual_only = any(d in url for d in MANUAL_ONLY_DOMAINS)
            badge_html = (f'<span class="badge badge-warn">MANUAL APPLY</span>' if manual_only
                          else f'<span class="badge badge-ok">AUTO-APPLY OK</span>')

            with cols[idx % 3]:
                st.markdown(
                    f"""<div class="job-card">
                        <div style="font-weight:800;color:#1e1b4b;font-size:1.05rem;">{title}</div>
                        <div style="color:#4b5563;margin:4px 0;">🏢 {company}</div>
                        <div style="color:#6b7280;font-size:0.9rem;margin-bottom:8px;">📍 {location}</div>
                        {badge_html}
                    </div>""",
                    unsafe_allow_html=True,
                )
                bcol1, bcol2 = st.columns(2)
                bcol1.link_button("Apply Now →", url, use_container_width=True)
                is_saved = job_id in st.session_state.saved_jobs
                if bcol2.button("★ Saved" if is_saved else "☆ Save", key=f"save_{job_id}_{idx}", use_container_width=True):
                    if is_saved:
                        st.session_state.saved_jobs.discard(job_id)
                    else:
                        st.session_state.saved_jobs.add(job_id)
                    st.rerun()
                if not manual_only:
                    if st.button("+ Add to auto-apply list", key=f"add_{job_id}_{idx}", use_container_width=True):
                        st.session_state.setdefault("career_urls", [])
                        if url not in st.session_state.career_urls:
                            st.session_state.career_urls.append(url)
                        st.success("Added below ↓")
                else:
                    st.caption("This platform prohibits auto-apply bots -- use the fit-pitch/tailored resume above, then apply yourself.")

    if st.session_state.get("career_urls"):
        st.subheader("Your auto-apply list (for apply_bot.py)")
        urls_text = "\n".join(st.session_state.career_urls)
        st.text_area("One URL per line -- copy into career_urls.txt", value=urls_text, height=100)
        st.download_button("Download career_urls.txt", data=urls_text, file_name="career_urls.txt")

# ---------------------------------------------------------------------------
# 5. Contacts
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">5. Contacts</div>', unsafe_allow_html=True)
st.caption("Add each person you want to email. Leave name/role/job_title blank to fall back to the defaults above "
           "-- if name is blank, it'll be guessed from the email address instead.")

upload_col, _ = st.columns([1, 2])
uploaded_contacts = upload_col.file_uploader(
    "Or upload a contact list (.csv or .xlsx)",
    type=["csv", "xlsx"],
    help="Needs at least an 'email' column. 'name', 'role', 'company', 'job_title' columns are optional.",
)
if uploaded_contacts is not None:
    try:
        if uploaded_contacts.name.lower().endswith(".xlsx"):
            up_df = pd.read_excel(uploaded_contacts)
        else:
            up_df = pd.read_csv(uploaded_contacts)
        up_df.columns = [c.strip().lower() for c in up_df.columns]
        if "email" not in up_df.columns:
            st.error("That file needs an 'email' column.")
        else:
            for col in ["name", "role", "company", "job_title"]:
                if col not in up_df.columns:
                    up_df[col] = ""
            up_df = up_df[["email", "name", "role", "company", "job_title"]].fillna("")
            st.session_state.recipients_df = up_df
            st.success(f"Loaded {len(up_df)} contact(s) from {uploaded_contacts.name}")
    except Exception as e:
        st.error(f"Couldn't read that file: {e}")

if "recipients_df" not in st.session_state:
    st.session_state.recipients_df = pd.DataFrame(
        [{"email": "", "name": "", "role": "", "company": "", "job_title": ""}]
    )

recipients_df = st.data_editor(
    st.session_state.recipients_df,
    num_rows="dynamic",
    use_container_width=True,
    key="recipients_editor",
)
st.session_state.recipients_df = recipients_df

# ---------------------------------------------------------------------------
# 5. Draft
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">6. Draft emails</div>', unsafe_allow_html=True)

if st.button("Generate drafts", disabled=not (resume_text and groq_key)):
    rows = recipients_df[recipients_df["email"].astype(str).str.strip() != ""]
    if rows.empty:
        st.warning("Add at least one contact with an email address above.")
    else:
        drafts = []
        progress = st.progress(0.0, text="Drafting emails...")
        for i, row in enumerate(rows.to_dict("records")):
            company = row.get("company") or target_company
            job_title = row.get("job_title") or target_role
            contact_name = row.get("name") or name_from_email(row["email"]) or "Hiring contact"
            prompt = f"""Write a concise, professional job-outreach email from a candidate to a contact at a company.

Candidate resume (for context on real background -- do not fabricate anything not in here):
---
{resume_text}
---
{f"Job description for the target role:\n---\n{jd_text}\n---\n" if jd_text.strip() else ""}
Recipient: {contact_name}{f", {row.get('role')}" if row.get("role") else ""} at {company or "the company"}
Target position: {job_title or "a relevant open role"}
Sender's name to sign as: {sender_name or "[Your Name]"}

Requirements:
- Subject line: short, specific, not spammy
- Body: 120-180 words, warm but professional, no generic filler ("I am writing to express...")
- Reference 1-2 concrete, real details from the resume that make the candidate relevant{" to the job description" if jd_text.strip() else ""}
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
            drafts.append({"email": row["email"], "subject": subject, "body": body})
            progress.progress((i + 1) / len(rows), text=f"Drafted {i + 1}/{len(rows)}")
        st.session_state.drafts = drafts
        st.rerun()

if st.session_state.get("drafts"):
    st.subheader("Review & edit before sending")
    for i, d in enumerate(st.session_state.drafts):
        with st.expander(f"✉️ {d['email']} — {d['subject']}"):
            d["subject"] = st.text_input("Subject", value=d["subject"], key=f"subj_{i}")
            d["body"] = st.text_area("Body", value=d["body"], height=220, key=f"body_{i}")

# ---------------------------------------------------------------------------
# 7. Send
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">7. Send</div>', unsafe_allow_html=True)

ready_to_send = bool(
    st.session_state.get("drafts") and resume_file and gmail_address and gmail_app_password
)
if not ready_to_send:
    st.info("Generate drafts above and fill in your Gmail credentials in the sidebar to enable sending.")

confirm = st.text_input('Type SEND (all caps) to enable the send button', value="")
send_clicked = st.button(
    f"Send {len(st.session_state.get('drafts', []))} email(s) now",
    disabled=not (ready_to_send and confirm == "SEND"),
    type="primary",
)

if send_clicked:
    resume_bytes = resume_file.getvalue()
    results = []
    progress = st.progress(0.0, text="Sending...")
    drafts = st.session_state.drafts
    for i, d in enumerate(drafts):
        try:
            send_email(d["email"], d["subject"], d["body"], resume_bytes, resume_file.name)
            results.append({"email": d["email"], "status": "sent"})
        except Exception as e:
            results.append({"email": d["email"], "status": f"FAILED: {e}"})
        progress.progress((i + 1) / len(drafts), text=f"Sent {i + 1}/{len(drafts)}")
        if i < len(drafts) - 1:
            time.sleep(seconds_between)
    st.success("Done.")
    st.table(pd.DataFrame(results))
