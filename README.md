# RingCentral RingCX Transcript Downloader

A web app that lets any RingCX admin log in and export all their contact-center call transcripts to Excel and PDF with one click.

Forked from [nmludwig/rc-transcripts](https://github.com/nmludwig/rc-transcripts) (RingEX/RingSense version).

---

## What It Does

1. Admin visits the URL and clicks **Login with RingCentral**
2. They authenticate with their RingCentral admin credentials (handled directly by RingCentral — no credentials stored)
3. They select a **RingCX sub-account**, enter a customer/account name, and pick a date range
4. The app queries the **RingCX Integration API** to fetch interaction metadata, then pulls each transcript
5. When complete, the admin downloads:
   - **Excel spreadsheet** — Summary tab, All Interactions tab, Transcripts tab
   - **PDF report** — formatted interaction-by-interaction transcripts with speaker labels

Downloads take about 0.5 seconds per interaction segment.

---

## How It Differs from the RingEX Version

| | RingEX version | This (RingCX) version |
|---|---|---|
| Call discovery | `platform.ringcentral.com` call log | `ringcx.ringcentral.com` interaction-metadata API |
| Transcript source | RingSense AI insights API | RingCX transcript API |
| Auth scope needed | RingSense, ReadCallLog, etc. | `ReadAccounts` only |
| Sub-account | Not required | Required — auto-discovered |
| AI requirement | RingSense license per user | "Enable AI Summaries" per queue in RingCX Admin |

---

## Files

```
├── app.py              ← Flask server (all backend logic)
├── requirements.txt    ← Python dependencies
├── README.md           ← This file
├── templates/
│   ├── index.html      ← Single-page UI
│   └── error.html      ← OAuth error page
└── outputs/            ← Generated files (auto-created)
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
- **Scopes:** `ReadAccounts`

> **Note:** Unlike the RingEX version, `ReadAccounts` is the only OAuth scope that has any effect on RingCX APIs. All other permissions are managed in the **RingCX Admin portal**, not in the developer console.

---

## RingCX Admin Prerequisites

Before transcripts will appear, you must enable AI transcription per queue:

1. Log in to **RingCX Admin**
2. Go to **Routing → Voice/Digital Queues & Skills**
3. Select your target queue → **AI Tools** section
4. Check **"Enable AI Summaries"**

You may also need your RingCentral representative to enable:
- **WEM (Workforce Engagement Management) access** on the account — required to call the integration/metadata API
- **Recording** — must be manually activated per account

---

## API Flow

```
OAuth login
    ↓
Discover sub-accounts  →  /cx/integration/v1/accounts/{rcAccountId}/sub-accounts
    ↓
Poll interaction metadata (1-hour windows)
    POST /cx/integration/v1/accounts/{rcAccountId}/sub-accounts/{subAccountId}/interaction-metadata
    ↓
Fetch transcript per segment
    GET /cx/integration/v1/accounts/{rcAccountId}/sub-accounts/{subAccountId}/transcripts/dialogs/{dialogId}/segments/{segmentId}
    ↓
Build Excel + PDF
```

---

## Rate Limits

- Interaction metadata endpoint: **2 calls/minute** — the app uses 1-hour time windows with a 600ms delay
- Transcript endpoint: **120 requests/minute** — the app uses a 500ms delay per segment
- Tokens are automatically refreshed every 50 calls

---

## Notes

- Only interactions where AI transcription is enabled on the queue will have transcript content
- Transcripts are available approximately **5 minutes** after a segment ends
- Output files are served via `/api/download/<job_id>/<type>` for the duration of the server session
