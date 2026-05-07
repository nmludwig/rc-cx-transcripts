# RingCentral ACE Transcript Downloader — Web App

A polished web UI wrapping the `download_transcripts.py` CLI script.  
Users enter their RingCentral credentials, choose a date range, watch live progress,  
then download finished Excel + PDF reports.

---

## What's Included

```
rc_transcripts/
├── app.py              ← Flask server (all backend logic)
├── requirements.txt    ← Python dependencies
├── README.md           ← This file
├── templates/
│   └── index.html      ← Full single-page UI (4-step wizard)
└── outputs/            ← Generated files saved here (auto-created)
```

---

## Quick Start (Local)

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the dev server
python app.py

# 3. Open http://localhost:5000
```

---

## Production Deploy

### Option A — Any Linux VPS (nginx + gunicorn)

```bash
# Install
pip install -r requirements.txt

# Run with gunicorn (4 workers)
gunicorn -w 4 -b 0.0.0.0:5000 app:app

# Or with a custom secret key
FLASK_SECRET=your-secret-here gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

**nginx config** (`/etc/nginx/sites-available/rc-transcripts`):

```nginx
server {
    listen 80;
    server_name yourdomain.com;

    client_max_body_size 50M;

    location / {
        proxy_pass         http://127.0.0.1:5000;
        proxy_set_header   Host $host;
        proxy_set_header   X-Real-IP $remote_addr;
        proxy_read_timeout 300s;   # allow long transcript fetches
        proxy_send_timeout 300s;
    }
}
```

```bash
sudo ln -s /etc/nginx/sites-available/rc-transcripts /etc/nginx/sites-enabled/
sudo nginx -t && sudo systemctl reload nginx
```

---

### Option B — Render.com (free tier)

1. Push this folder to a GitHub repo
2. New Web Service → Python → Build: `pip install -r requirements.txt`
3. Start command: `gunicorn -w 2 -b 0.0.0.0:$PORT app:app`
4. Set env var `FLASK_SECRET` to a random string

---

### Option C — Railway / Fly.io

Both auto-detect `requirements.txt`. Set `PORT` and `FLASK_SECRET` env vars.  
Start command: `gunicorn app:app`

---

## Environment Variables

| Variable       | Default                          | Description                  |
|----------------|----------------------------------|------------------------------|
| `PORT`         | `5000`                           | Port to listen on            |
| `FLASK_SECRET` | `rc-ace-demo-secret-change-in-prod` | Session signing key — **change this in production** |

---

## JWT Auth Flow

1. User enters Client ID, Client Secret, and JWT Token in the browser
2. The server POSTs to `https://platform.ringcentral.com/restapi/oauth/token`  
   using the `urn:ietf:params:oauth:grant-type:jwt-bearer` grant type
3. On success, the server stores the `access_token` in the **server-side Flask session**  
   (never returned to the browser)
4. Subsequent API calls use the stored token — credentials are never persisted to disk

---

## Notes

- **Credentials are never saved.** They exist only in memory for the duration of the session.
- Long-running jobs run in a background thread; the browser polls `/api/job/<id>` every 1.5 s.
- Output files are written to `outputs/` and served via `/api/download/<job_id>/<type>`.
- For production, consider adding auth in front of the app (e.g., Basic Auth via nginx).
