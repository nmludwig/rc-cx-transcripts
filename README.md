# RingCentral ACE Transcript Downloader

A web app that lets any RingCentral admin log in and export all their ACE call transcripts to Excel and PDF with one click.

**Live demo:** https://rc-transcripts.onrender.com

---

## What It Does

1. Admin visits the URL and clicks **Login with RingCentral**
2. They authenticate with their RingCentral admin credentials (handled directly by RingCentral — no credentials are stored)
3. They enter a customer/account name and pick a date range (7, 30, 90 days or custom)
4. The app pulls every recorded call from the call log, fetches the ACE AI transcript, summary, and sentiment for each one, and shows a live progress log
5. When complete, the admin downloads:
   - **Excel spreadsheet** — Summary tab, All Calls tab, Transcripts tab
   - **PDF report** — formatted call-by-call transcripts with speaker labels, sentiment badges, and AI summaries

Downloads take about 1.5 seconds per call due to RingCentral API rate limits.

---

## Files

```
├── app.py              ← Flask server (all backend logic)
├── requirements.txt    ← Python dependencies
├── README.md           ← This file
├── templates/
│   ├── index.html      ← Single-page UI (3-step wizard)
│   └── error.html      ← OAuth error page
└── outputs/            ← Generated files saved here (auto-created)
```

---

## Setup (Local)

```bash
pip install -r requirements.txt
gunicorn -w 2 -k gthread --threads 4 --timeout 300 -b 0.0.0.0:5000 app:app
```

Open **http://localhost:5000**

---

## Deploy on Render

1. Push this repo to GitHub
2. **Render → New Web Service** → connect your repo
3. Set these values:
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn -w 2 -k gthread --threads 4 --timeout 300 -b 0.0.0.0:$PORT app:app`
4. Add environment variables:

| Variable | Description |
|---|---|
| `RC_CLIENT_ID` | Your RingCentral app Client ID |
| `RC_CLIENT_SECRET` | Your RingCentral app Client Secret |
| `RC_REDIRECT_URI` | `https://your-app.onrender.com/oauth/callback` |
| `FLASK_SECRET` | Any random string for session signing |

---

## RingCentral App Setup (developers.ringcentral.com)

Your RingCentral app needs:
- **Auth type:** 3-legged OAuth — Server-side web app
- **OAuth Redirect URI:** `https://your-app.onrender.com/oauth/callback`
- **Scopes:** Analytics, Read Accounts, Read Call Log, Read Call Recording, Read Contacts, RingSense

---

## Auth Flow

- Users authenticate via **RingCentral OAuth** (Authorization Code flow)
- The server stores the access token and refresh token in a server-side session — never exposed to the browser
- Tokens are automatically refreshed every 50 calls during long downloads to prevent expiry on large accounts
- Credentials are never saved to disk

---

## Notes

- Only calls with ACE/RingSense licenses assigned will have transcripts — calls without licenses show "No transcript available"
- The RingSense public API must be enabled for the account by RingCentral support
- Output files are served via `/api/download/<job_id>/<type>` and exist for the duration of the server session