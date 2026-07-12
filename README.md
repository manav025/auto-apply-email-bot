# Job Search Automation Toolkit

Two local tools that share one free AI backend ([Groq](https://console.groq.com)):

1. **`job_email_sender.py`** -- reads your resume, gets AI feedback on it,
   drafts a personalized outreach email per contact, sends it from your
   Gmail with your resume attached.
2. **`apply_bot.py`** -- opens company career page job postings you give it,
   auto-fills the application form from your resume/profile, AI-drafts any
   "why do you want this job" style questions, and leaves it open for you
   to review and submit (or asks for a typed confirmation per application
   with `--auto-submit`).

Both run on your own computer using your own Gmail account and your own
free Groq API key -- nothing here costs money at normal job-search volume.

## 1. Install

```bash
pip install -r requirements.txt
playwright install chromium   # only needed for apply_bot.py
```

(Python 3.10+ recommended.)

## 2. Set up credentials

```bash
cp .env.example .env
```

Fill in `.env`:
- **`GROQ_API_KEY`** -- free, no credit card. Get one at
  https://console.groq.com/keys. Free tier covers this use case comfortably;
  if you ever hit a rate limit, the script will tell you and you just wait
  a bit or lower usage.
- **`GMAIL_ADDRESS` / `GMAIL_APP_PASSWORD`** -- only needed for the email
  sender. Turn on 2-Step Verification at
  https://myaccount.google.com/security, then create an app password at
  https://myaccount.google.com/apppasswords.

## 3. Add your resume

Put `resume.pdf`, `resume.docx`, or `resume.txt` in this folder (or pass
`--resume path/to/file` to either script).

## 4. Email tool workflow

```bash
cp recipients.example.csv recipients.csv   # edit with real contacts

python job_email_sender.py --analyze-resume   # AI feedback, nothing sent
python job_email_sender.py --dry-run          # preview drafted emails
python job_email_sender.py --send             # type SEND to confirm, then sends
```

## 5. Auto-apply tool workflow

```bash
cp profile.json.example profile.json         # fill in your real info
cp career_urls.example.txt career_urls.txt   # add job posting URLs, one per line

python apply_bot.py                # fills forms, screenshots, leaves browser open for you to submit
python apply_bot.py --auto-submit  # after you review each one, type SUBMIT to submit it
python apply_bot.py --headless     # no visible browser window (still won't submit without --auto-submit)
```

**What it will and won't do:**
- It fills in factual fields (name, email, phone, LinkedIn, resume upload)
  from `profile.json`, and AI-drafts open-ended questions ("why this role")
  from your resume + the job posting text.
- It **never** auto-fills work authorization, visa/sponsorship, salary, or
  EEO/demographic self-identification questions -- those are always left
  for you, since guessing on them can genuinely matter.
- It does **not** submit anything by default -- you get a screenshot and a
  summary, and the browser tab stays open for you to double check.
- It does **not** attempt to solve or bypass CAPTCHAs; it'll tell you to
  finish that one manually.
- Works best on standard ATS platforms (Greenhouse, Lever, Workday, and
  most custom company career pages). Some sites' terms of service restrict
  automated applications -- worth a quick check for companies you really
  care about, and go manual there if in doubt.

## 6. Run the email tool as a web app instead (recommended if your laptop is slow)

`streamlit_app.py` is the richer, browser-based version of the email tool,
and it can run entirely in the cloud for free -- nothing taxes your laptop.
It covers everything the CLI does, plus:

- **JD-aware resume feedback** -- paste a job description and it maps the
  JD's keywords against your resume, flags gaps, and suggests specific
  rewrites (not just "use more action verbs").
- **Resume tailoring** -- generates a content-tailored draft of your resume
  aimed at a specific JD (never invents experience -- only reorders/reworks
  what's already there), downloadable as .docx.
- **"Why you're a fit" 3-liner** -- a short pitch you can paste into cover
  letters, emails, or application free-text fields.
- **Find open roles** -- searches Adzuna's free public API (a legitimate
  aggregator, not a scraper) and flags each result as auto-apply-safe
  (direct company career page → hand off to `apply_bot.py`) or manual-only
  (LinkedIn/Indeed/Naukri/Glassdoor → these platforms' own terms prohibit
  automated applying, so you apply yourself using the fit-pitch/tailored
  resume above).
- **Bulk contacts via Excel/CSV upload** -- upload a sheet with just an
  `email` column and it fills in a guessed name from the address (e.g.
  `jane.doe@acme.com` → "Jane Doe") wherever you haven't provided one.

**Try it locally first:**
```bash
streamlit run streamlit_app.py
```
Opens at http://localhost:8501. Enter your Groq/Gmail credentials in the
sidebar (kept only for that browser session, never saved to disk). Adzuna
keys are optional -- get a free pair at https://developer.adzuna.com/ if you
want the "Find open roles" search.

**Host it for free on Streamlit Community Cloud:**
1. Push this repo to GitHub (steps below).
2. Go to https://share.streamlit.io, sign in with GitHub.
3. Click "New app", pick this repo, branch `main`, main file `streamlit_app.py`.
4. Deploy. You get a public `*.streamlit.app` URL that runs in Streamlit's
   cloud, not your machine.
5. **Set secrets** (optional but recommended): in the app's Settings ->
   Secrets, paste the contents of `.streamlit/secrets.toml.example` filled
   in with your real values. This pre-fills the sidebar so you don't
   retype credentials each visit, and lets you set `APP_PASSWORD` -- since
   deployed apps get a public URL, this stops random visitors from opening
   it and using your Gmail/Groq credentials.

Note: Community Cloud apps get 1 GB RAM and sleep when idle (a visit wakes
them back up in a few seconds) -- plenty for this tool, since it's just API
calls and sending mail, no heavy computation.

**What about LinkedIn / Indeed / Naukri auto-apply, and `apply_bot.py`?**
Two separate reasons these stay hands-off/local:

- LinkedIn, Indeed, and Naukri all explicitly prohibit bots or automated
  methods in their own terms of service (Indeed: *"Use of any automation,
  scripting, or bots to automate the Indeed Apply process... is
  prohibited"*; LinkedIn's User Agreement §8.2 bans automated access
  outright; Naukri prohibits automated extraction/access without written
  permission). All three actively detect and ban accounts for it. This
  tool won't automate those platforms -- use the tailored resume and fit
  pitch above, then apply yourself in a couple of clicks.
- `apply_bot.py` (for direct company career pages -- Greenhouse, Lever,
  Workday, custom sites, which don't carry that same platform-wide
  restriction) drives a real Chromium browser, which is too heavy for
  Streamlit Cloud's free tier and has no screen there for you to review
  the filled form before submitting. It's a quick, occasional task though,
  so keep running it locally -- your laptop speed matters much less for an
  occasional task than a continuous one. If you'd rather it be off your
  laptop entirely, GitHub Codespaces (free ~60 hrs/month) gives you a full
  cloud machine with a browser-based VS Code -- ask and I can set up a
  `.devcontainer` config for that.

## 7. Put this on GitHub

The `.gitignore` is already set up to exclude your secrets and personal
data (`.env`, `resume.*`, `profile.json`, `recipients.csv`,
`career_urls.txt`, logs, screenshots) -- only the code and `.example`
templates get committed.

```bash
git init
git add .
git commit -m "Initial commit: job search automation toolkit"
```

Then on github.com, create a new empty repository (no README/license, so it
doesn't conflict), and:

```bash
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO_NAME.git
git branch -M main
git push -u origin main
```

Double check `git status` before your first push and confirm `.env`,
`resume.*`, `profile.json`, `recipients.csv`, and `career_urls.txt` are
**not** listed as tracked -- those contain your personal info and API
credentials and should never end up in a public repo.

## Good practice

- **Personalize where it matters.** Several near-identical emails to
  contacts at the same company can look careless -- consider starting with
  one strong contact (recruiter or hiring manager).
- **Real, consented addresses only.** This is meant for a short, deliberate
  list of real contacts, not scraped bulk lists -- both for your sender
  reputation and because unsolicited bulk email has real anti-spam rules
  (e.g. CAN-SPAM in the US) attached to it at volume.
- **Review AI-drafted answers before submitting.** They're a strong first
  draft, not a guarantee of accuracy -- especially for the free-text
  "why this role" style questions.

## Files

| File | Purpose |
|---|---|
| `job_email_sender.py` | Cold-email tool (CLI) |
| `streamlit_app.py` | Cold-email tool (web UI, can run free on Streamlit Cloud) |
| `apply_bot.py` | Career-page auto-fill tool |
| `ai_client.py` | Shared Groq API wrapper (used by the CLI tools) |
| `.env` | Your secrets for CLI use (create from `.env.example`) |
| `.streamlit/secrets.toml` | Your secrets for the web app (create from `.streamlit/secrets.toml.example`) |
| `recipients.csv` | People to email (create from `.example`) |
| `career_urls.txt` | Job postings to auto-apply to (create from `.example`) |
| `profile.json` | Your factual info for auto-fill (create from `.example`) |
| `resume_suggestions.md` | Generated resume feedback |
| `sent_log.csv` | Record of emails actually sent |
| `screenshots/` | Screenshot of each filled application |
