#!/usr/bin/env python3
"""
RingCentral ACE Transcript Downloader — Web Server
Wraps download_transcripts.py logic in a Flask web app with JWT auth flow.
"""

import sys
import os
import re
import json
import time
import threading
import uuid
from datetime import datetime, timezone
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file, session
try:
    from flask_cors import CORS
except ImportError:
    CORS = None

# ── Auto-install dependencies ─────────────────────────────────────────────────
def install(import_name, pip_name=None):
    try:
        __import__(import_name)
    except ImportError:
        pkg = pip_name or import_name
        import subprocess
        subprocess.run([sys.executable, "-m", "pip", "install", pkg, "-q"], check=True)

install("flask")
install("flask_cors")
install("requests")
install("openpyxl")
install("reportlab")
install("PIL", "pillow")

import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "rc-ace-demo-secret-change-in-prod")
if CORS:
    CORS(app)

# ── In-memory job store ───────────────────────────────────────────────────────
jobs = {}          # job_id -> { status, progress, log, records, files }
OUTPUT_DIR = Path(__file__).parent / "outputs"
OUTPUT_DIR.mkdir(exist_ok=True)


# ─────────────────────────────────────────────────────────────────────────────
#  ROUTES
# ─────────────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/authenticate", methods=["POST"])
def api_authenticate():
    """Validate credentials against RingCentral and return account info."""
    data = request.json or {}
    client_id     = data.get("client_id", "").strip()
    client_secret = data.get("client_secret", "").strip()
    jwt_token     = data.get("jwt_token", "").strip()

    if not all([client_id, client_secret, jwt_token]):
        return jsonify({"error": "All three fields are required."}), 400

    try:
        resp = requests.post(
            "https://platform.ringcentral.com/restapi/oauth/token",
            auth=(client_id, client_secret),
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
                "assertion":  jwt_token,
            },
            timeout=15,
        )
        resp.raise_for_status()
    except requests.exceptions.HTTPError as e:
        code = e.response.status_code if e.response else 0
        if code == 401:
            return jsonify({"error": "Authentication failed. Check your Client ID, Client Secret, and JWT Token."}), 401
        if code == 400:
            return jsonify({"error": "JWT Token may be expired. Please regenerate it at developers.ringcentral.com."}), 400
        return jsonify({"error": f"RingCentral returned HTTP {code}."}), 400
    except requests.exceptions.ConnectionError:
        return jsonify({"error": "Cannot reach RingCentral. Check your internet connection."}), 503

    rc_data = resp.json()
    token   = rc_data["access_token"]
    scope   = rc_data.get("scope", "")

    # Look up account ID
    try:
        acct_resp = requests.get(
            "https://platform.ringcentral.com/restapi/v1.0/account/~/call-log",
            headers={"Authorization": "Bearer " + token},
            params={"perPage": 1},
            timeout=15,
        )
        acct_resp.raise_for_status()
        match = re.search(r"account/(\d+)", acct_resp.json().get("uri", ""))
        account_id = match.group(1) if match else "~"
    except Exception:
        account_id = "~"

    has_ringsense = "RingSense" in scope or "ringsense" in scope.lower()

    # Store in session (server-side only – never sent back raw)
    session["token"]      = token
    session["account_id"] = account_id
    session["scope"]      = scope

    return jsonify({
        "success":       True,
        "account_id":    account_id,
        "has_ringsense": has_ringsense,
        "scope_preview": scope[:120] + ("…" if len(scope) > 120 else ""),
    })


@app.route("/api/start-job", methods=["POST"])
def api_start_job():
    """Kick off background download job."""
    if "token" not in session:
        return jsonify({"error": "Not authenticated. Please go back and re-enter credentials."}), 401

    data          = request.json or {}
    customer_name = data.get("customer_name", "Customer").strip() or "Customer"
    date_from     = data.get("date_from", "")
    date_to       = data.get("date_to", "")

    if not date_from or not date_to:
        return jsonify({"error": "date_from and date_to are required."}), 400

    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "status":    "running",
        "progress":  0,
        "log":       [],
        "records":   [],
        "files":     {},
        "error":     None,
    }

    token      = session["token"]
    account_id = session["account_id"]

    thread = threading.Thread(
        target=run_download_job,
        args=(job_id, token, account_id, customer_name, date_from, date_to),
        daemon=True,
    )
    thread.start()

    return jsonify({"job_id": job_id})


@app.route("/api/job/<job_id>", methods=["GET"])
def api_job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        return jsonify({"error": "Job not found"}), 404
    return jsonify({
        "status":   job["status"],
        "progress": job["progress"],
        "log":      job["log"][-50:],   # last 50 lines
        "files":    list(job["files"].keys()),
        "error":    job["error"],
        "summary":  job.get("summary"),
    })


@app.route("/api/download/<job_id>/<file_type>", methods=["GET"])
def api_download(job_id, file_type):
    job = jobs.get(job_id)
    if not job or file_type not in job["files"]:
        return jsonify({"error": "File not found"}), 404
    path = Path(job["files"][file_type])
    if not path.exists():
        return jsonify({"error": "File no longer available"}), 404
    return send_file(str(path), as_attachment=True, download_name=path.name)


# ─────────────────────────────────────────────────────────────────────────────
#  BACKGROUND JOB
# ─────────────────────────────────────────────────────────────────────────────

def job_log(job_id, msg, level="info"):
    jobs[job_id]["log"].append({"t": datetime.now().strftime("%H:%M:%S"), "msg": msg, "level": level})


def run_download_job(job_id, token, account_id, customer_name, date_from, date_to):
    job = jobs[job_id]
    try:
        job_log(job_id, f"Starting download for {customer_name}")
        job_log(job_id, f"Date range: {date_from[:10]} → {date_to[:10]}")

        # ── Fetch call logs ───────────────────────────────────────────────────
        job["progress"] = 5
        job_log(job_id, "Fetching call log from RingCentral…")
        records = []
        page    = 1
        while True:
            resp = requests.get(
                f"https://platform.ringcentral.com/restapi/v1.0/account/{account_id}/call-log",
                headers={"Authorization": "Bearer " + token},
                params={
                    "view":          "Detailed",
                    "dateFrom":      date_from + "T00:00:00Z",
                    "dateTo":        date_to   + "T23:59:59Z",
                    "type":          "Voice",
                    "withRecording": "true",
                    "perPage":       250,
                    "page":          page,
                },
                timeout=30,
            )
            resp.raise_for_status()
            data = resp.json()
            records.extend(data.get("records", []))
            if not data.get("navigation", {}).get("nextPage"):
                break
            page += 1
            time.sleep(0.25)

        job["progress"] = 15
        job_log(job_id, f"Found {len(records)} recorded calls", "ok")

        if not records:
            job["status"] = "done"
            job["progress"] = 100
            job["summary"] = {"total": 0, "transcripts": 0}
            job_log(job_id, "No recorded calls found in this date range.", "warn")
            return

        # ── Fetch transcripts ─────────────────────────────────────────────────
        total = len(records)
        job_log(job_id, f"Fetching RingSense transcripts for {total} calls…")
        transcript_records = []
        with_transcripts   = 0

        for i, call in enumerate(records):
            recording_id = call.get("recording", {}).get("id")
            if not recording_id:
                continue

            pct = 15 + int(70 * (i / max(total, 1)))
            job["progress"] = pct

            insights = None
            url = (
                "https://platform.ringcentral.com/ai/ringsense/v1/public"
                f"/accounts/~/domains/pbx/records/{recording_id}/insights"
            )
            while True:
                try:
                    r = requests.get(url, headers={"Authorization": "Bearer " + token}, timeout=30)
                    if r.status_code == 429:
                        job_log(job_id, "Rate limit hit — waiting 65 s…", "warn")
                        time.sleep(65)
                        continue
                    if r.status_code == 200:
                        insights = r.json()
                    break
                except Exception:
                    break

            if insights:
                with_transcripts += 1

            speaker_map = {}
            if insights:
                for sp in insights.get("speakerInfo", []):
                    sid  = sp.get("speakerId", "")
                    name = sp.get("name", "") or sp.get("phoneNumber", sid)
                    if sid and name:
                        speaker_map[sid] = name

            utterances = (insights or {}).get("insights", {}).get("Transcript", [])
            lines = []
            for u in utterances:
                sid   = u.get("speakerId", "?")
                name  = speaker_map.get(sid, sid)
                txt   = u.get("text", "").strip()
                start = u.get("start", 0)
                mm    = str(int(start // 60)).zfill(2)
                ss    = str(int(start % 60)).zfill(2)
                lines.append(f"[{mm}:{ss}] {name}: {txt}")

            summary_list   = (insights or {}).get("insights", {}).get("Summary", [])
            sentiment_list = (insights or {}).get("insights", {}).get("Sentiment", [])
            summary        = summary_list[0].get("value", "")   if summary_list   else ""
            sentiment      = sentiment_list[0].get("value", "") if sentiment_list else ""

            rec = call.get("recording", {})
            transcript_records.append({
                "call_id":        call.get("id", ""),
                "start_time":     call.get("startTime", ""),
                "duration_sec":   call.get("duration", 0),
                "direction":      call.get("direction", ""),
                "from_number":    call.get("from", {}).get("phoneNumber", ""),
                "from_name":      call.get("from", {}).get("name", ""),
                "to_number":      call.get("to",   {}).get("phoneNumber", ""),
                "to_name":        call.get("to",   {}).get("name", ""),
                "recording_id":   rec.get("id", ""),
                "has_transcript": insights is not None,
                "sentiment":      sentiment,
                "summary":        summary,
                "transcript":     "\n".join(lines),
            })

            if (i + 1) % 10 == 0:
                job_log(job_id, f"Processed {i+1}/{total} calls ({with_transcripts} transcripts so far)")

            time.sleep(1.5)

        job_log(job_id, f"{with_transcripts} of {total} calls have transcripts", "ok")

        # ── Build output files ────────────────────────────────────────────────
        job["progress"] = 88
        slug       = re.sub(r"[^a-z0-9]+", "_", customer_name.lower()).strip("_")
        date_stamp = datetime.now().strftime("%Y%m%d")

        job_log(job_id, "Building Excel spreadsheet…")
        xlsx_name = f"transcripts_{slug}_{date_stamp}.xlsx"
        xlsx_path = OUTPUT_DIR / xlsx_name
        build_excel(transcript_records, customer_name, date_from, date_to, xlsx_path)
        job["files"]["xlsx"] = str(xlsx_path)
        job_log(job_id, f"Excel saved: {xlsx_name}", "ok")

        job["progress"] = 94
        job_log(job_id, "Building PDF (may take 15–30 seconds)…")
        pdf_name  = f"transcripts_{slug}_{date_stamp}.pdf"
        pdf_path  = OUTPUT_DIR / pdf_name
        build_pdf(transcript_records, customer_name, date_from, date_to, pdf_path)
        job["files"]["pdf"] = str(pdf_path)
        job_log(job_id, f"PDF saved: {pdf_name}", "ok")

        job["progress"] = 100
        job["status"]   = "done"
        job["records"]  = transcript_records
        job["summary"]  = {
            "total":       len(records),
            "transcripts": with_transcripts,
            "customer":    customer_name,
            "date_from":   date_from[:10],
            "date_to":     date_to[:10],
        }
        job_log(job_id, "All done! Files are ready to download.", "ok")

    except Exception as e:
        job["status"] = "error"
        job["error"]  = str(e)
        job_log(job_id, f"Error: {e}", "error")


# ─────────────────────────────────────────────────────────────────────────────
#  EXCEL / PDF BUILDERS  (ported from download_transcripts.py)
# ─────────────────────────────────────────────────────────────────────────────

def build_excel(records, customer_name, date_from, date_to, out_path):
    wb = openpyxl.Workbook()

    RC_ORANGE   = "FF6A00"
    DARK        = "1A1A1A"
    WHITE       = "FFFFFF"
    LIGHT_GREY  = "F5F5F5"
    BORDER_CLR  = "CCCCCC"

    thin   = Side(style="thin", color=BORDER_CLR)
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # Summary sheet
    ws_sum = wb.active
    ws_sum.title = "Summary"

    ws_sum.merge_cells("A1:H1")
    c = ws_sum["A1"]
    c.value     = f"RingCentral ACE Transcript Report — {customer_name}"
    c.font      = Font(name="Calibri", size=16, bold=True, color=WHITE)
    c.fill      = PatternFill("solid", fgColor=RC_ORANGE)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws_sum.row_dimensions[1].height = 32

    ws_sum.merge_cells("A2:H2")
    c2 = ws_sum["A2"]
    c2.value     = (f"{date_from[:10]} to {date_to[:10]}   |   "
                    f"Generated {datetime.now().strftime('%B %d, %Y')}   |   "
                    "RingCentral AI Conversation Expert   |   Confidential")
    c2.font      = Font(name="Calibri", size=10, color="555555")
    c2.fill      = PatternFill("solid", fgColor=LIGHT_GREY)
    c2.alignment = Alignment(horizontal="center", vertical="center")

    with_trans = len([r for r in records if r["has_transcript"]])
    total_sec  = sum(r["duration_sec"] for r in records)
    hrs, rem   = divmod(total_sec, 3600)
    mins       = rem // 60

    stats = [
        ("Total Calls",         str(len(records))),
        ("With Transcripts",    str(with_trans)),
        ("Without Transcripts", str(len(records) - with_trans)),
        ("Total Talk Time",     f"{hrs}h {mins}m"),
        ("Date From",           date_from[:10]),
        ("Date To",             date_to[:10]),
    ]
    ws_sum.append([])
    for row_idx, (label, value) in enumerate(stats, start=4):
        ws_sum.cell(row=row_idx, column=1).value = label
        ws_sum.cell(row=row_idx, column=1).font  = Font(name="Calibri", size=11, bold=True, color=DARK)
        ws_sum.cell(row=row_idx, column=1).fill  = PatternFill("solid", fgColor=LIGHT_GREY)
        ws_sum.cell(row=row_idx, column=2).value = value
        ws_sum.cell(row=row_idx, column=2).font  = Font(name="Calibri", size=11, color=DARK)

    ws_sum.column_dimensions["A"].width = 24
    ws_sum.column_dimensions["B"].width = 20

    # All Calls sheet
    ws_calls = wb.create_sheet("All Calls")
    headers    = ["Date","Time","Direction","Duration","From Name","From Number","To Name","To Number","Has Transcript","Sentiment","AI Summary"]
    col_widths = [14, 10, 12, 10, 22, 18, 22, 18, 14, 12, 60]

    for col, (hdr, width) in enumerate(zip(headers, col_widths), start=1):
        cell = ws_calls.cell(row=1, column=col)
        cell.value     = hdr
        cell.font      = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=RC_ORANGE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = border
        ws_calls.column_dimensions[get_column_letter(col)].width = width
    ws_calls.row_dimensions[1].height = 24

    for row_idx, rec in enumerate(records, start=2):
        try:
            dt       = datetime.fromisoformat(rec["start_time"].replace("Z", "+00:00"))
            date_str = dt.strftime("%Y-%m-%d")
            time_str = dt.strftime("%I:%M %p")
        except Exception:
            date_str = rec["start_time"][:10]
            time_str = ""

        dur_m, dur_s = divmod(int(rec["duration_sec"]), 60)
        row_data = [
            date_str, time_str, rec["direction"], f"{dur_m}m {dur_s}s",
            rec["from_name"], rec["from_number"],
            rec["to_name"],   rec["to_number"],
            "Yes" if rec["has_transcript"] else "No",
            rec["sentiment"], rec["summary"],
        ]
        for col, val in enumerate(row_data, start=1):
            cell = ws_calls.cell(row=row_idx, column=col)
            cell.value     = val
            cell.font      = Font(name="Calibri", size=10, color=DARK)
            cell.alignment = Alignment(vertical="top", wrap_text=(col == 11))
            cell.border    = border
        ws_calls.row_dimensions[row_idx].height = 15

    # Transcripts sheet
    ws_tr = wb.create_sheet("Transcripts")
    tr_headers  = ["Date","From Name","To Name","Duration","Sentiment","AI Summary","Full Transcript"]
    tr_widths   = [14, 22, 22, 10, 12, 60, 100]

    for col, (hdr, width) in enumerate(zip(tr_headers, tr_widths), start=1):
        cell = ws_tr.cell(row=1, column=col)
        cell.value     = hdr
        cell.font      = Font(name="Calibri", size=11, bold=True, color=WHITE)
        cell.fill      = PatternFill("solid", fgColor=RC_ORANGE)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border    = border
        ws_tr.column_dimensions[get_column_letter(col)].width = width

    for row_idx, rec in enumerate([r for r in records if r["has_transcript"]], start=2):
        try:
            dt       = datetime.fromisoformat(rec["start_time"].replace("Z", "+00:00"))
            date_str = dt.strftime("%Y-%m-%d")
        except Exception:
            date_str = rec["start_time"][:10]
        dur_m, dur_s = divmod(int(rec["duration_sec"]), 60)
        row_data = [date_str, rec["from_name"], rec["to_name"], f"{dur_m}m {dur_s}s",
                    rec["sentiment"], rec["summary"], rec["transcript"]]
        for col, val in enumerate(row_data, start=1):
            cell = ws_tr.cell(row=row_idx, column=col)
            cell.value     = val
            cell.font      = Font(name="Calibri", size=10, color=DARK)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border    = border
        ws_tr.row_dimensions[row_idx].height = 60

    wb.save(out_path)


def build_pdf(records, customer_name, date_from, date_to, out_path):
    """Ported PDF builder from download_transcripts.py."""
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.units import inch
    from reportlab.lib.colors import HexColor
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, HRFlowable, KeepTogether)
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
    from reportlab.pdfgen import canvas as rc_canvas

    RC_ORANGE = HexColor("#FF6A00")
    DARK      = HexColor("#1A1A1A")
    GREY      = HexColor("#666666")
    BG_GREEN  = HexColor("#E8F8EF")
    BG_RED    = HexColor("#FEECEC")
    BG_GREY   = HexColor("#F5F5F5")
    BG_BLUE   = HexColor("#E6F1FB")
    BG_ORANGE = HexColor("#FFF3E8")
    TX_GREEN  = HexColor("#1B6B35")
    TX_RED    = HexColor("#8B1A1A")
    TX_GREY   = HexColor("#444444")
    TX_BLUE   = HexColor("#0B4F8A")
    RULE      = HexColor("#E0E0E0")

    def ps(name, **kw):
        return ParagraphStyle(name, **kw)

    def_s      = ps("D",  fontName="Helvetica", fontSize=10, textColor=DARK, leading=14)
    call_num_s = ps("CN", fontName="Helvetica", fontSize=8,  textColor=GREY, leading=12)
    call_title_s = ps("CT", fontName="Helvetica-Bold", fontSize=13, textColor=DARK, leading=16, spaceAfter=4)
    meta_s     = ps("ME", fontName="Helvetica", fontSize=9, textColor=GREY, leading=12, spaceAfter=6)
    sum_lbl_s  = ps("SL", fontName="Helvetica-Bold", fontSize=8, textColor=RC_ORANGE, leading=14, spaceBefore=10)
    sum_txt_s  = ps("ST", fontName="Helvetica", fontSize=10, textColor=DARK, leading=14)
    tr_lbl_s   = ps("TL", fontName="Helvetica-Bold", fontSize=8, textColor=GREY, leading=14, spaceBefore=10)
    no_tr_s    = ps("NT", fontName="Helvetica-Oblique", fontSize=9, textColor=GREY, leading=12)
    utt_s      = ps("UT", fontName="Helvetica", fontSize=9, textColor=DARK, leading=13, leftIndent=10, spaceAfter=4)
    sp_a_s     = ps("SA", fontName="Helvetica-Bold", fontSize=9, textColor=HexColor("#0B4F8A"), leading=12, spaceBefore=6)
    sp_b_s     = ps("SB", fontName="Helvetica-Bold", fontSize=9, textColor=HexColor("#8B1A1A"), leading=12, spaceBefore=6)
    sp_styles  = [sp_a_s, sp_b_s,
                  ps("SC", fontName="Helvetica-Bold", fontSize=9, textColor=HexColor("#1B6B35"), leading=12, spaceBefore=6),
                  ps("SD", fontName="Helvetica-Bold", fontSize=9, textColor=HexColor("#5A189A"), leading=12, spaceBefore=6)]

    class NumCanvas(rc_canvas.Canvas):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            self._saved_page_states = []

        def showPage(self):
            self._saved_page_states.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            num_pages = len(self._saved_page_states)
            for state in self._saved_page_states:
                self.__dict__.update(state)
                self._draw_page_number(num_pages)
                super().showPage()
            super().save()

        def _draw_page_number(self, page_count):
            self.setFont("Helvetica", 8)
            self.setFillColor(HexColor("#AAAAAA"))
            self.drawRightString(
                letter[0] - 0.6 * inch,
                0.4 * inch,
                f"Page {self._pageNumber} of {page_count}",
            )

    doc = SimpleDocTemplate(
        str(out_path),
        pagesize=letter,
        leftMargin=0.75*inch, rightMargin=0.75*inch,
        topMargin=0.75*inch,  bottomMargin=0.75*inch,
    )

    story = []

    # Cover header
    cover_data = [[
        Paragraph(f"<b>RingCentral ACE Transcript Report</b>", ps("H1",
            fontName="Helvetica-Bold", fontSize=18, textColor=HexColor("#FFFFFF"), leading=22)),
        Paragraph(customer_name, ps("H2",
            fontName="Helvetica-Bold", fontSize=14, textColor=HexColor("#FFD0A8"),
            leading=18, alignment=TA_RIGHT)),
    ]]
    cover_tbl = Table(cover_data, colWidths=[doc.width * 0.6, doc.width * 0.4])
    cover_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), RC_ORANGE),
        ("TOPPADDING",    (0,0), (-1,-1), 14),
        ("BOTTOMPADDING", (0,0), (-1,-1), 14),
        ("LEFTPADDING",   (0,0), (0,0),  18),
        ("RIGHTPADDING",  (-1,0),(-1,-1),18),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(cover_tbl)
    story.append(Spacer(1, 4))

    # Subtitle bar
    subtitle_txt = (f"{date_from[:10]} → {date_to[:10]}   |   "
                    f"Generated {datetime.now().strftime('%B %d, %Y')}   |   Confidential")
    sub_tbl = Table([[Paragraph(subtitle_txt, ps("SU",
        fontName="Helvetica", fontSize=8.5, textColor=HexColor("#555555"), alignment=TA_CENTER))]],
        colWidths=[doc.width])
    sub_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), HexColor("#F5F5F5")),
        ("TOPPADDING",    (0,0), (-1,-1), 7),
        ("BOTTOMPADDING", (0,0), (-1,-1), 7),
    ]))
    story.append(sub_tbl)
    story.append(Spacer(1, 16))

    calls_to_show = [r for r in records if r["has_transcript"]]
    if not calls_to_show:
        calls_to_show = records

    for idx, row in enumerate(calls_to_show, start=1):
        start_time = row["start_time"]
        duration   = row["duration_sec"]
        direction  = row["direction"]
        from_name  = row["from_name"] or row["from_number"] or "Unknown"
        to_name    = row["to_name"]   or row["to_number"]   or "Unknown"
        summary    = row["summary"]
        sentiment  = row["sentiment"]
        transcript = row["transcript"]

        try:
            dt        = datetime.fromisoformat(start_time.replace("Z", "+00:00"))
            date_str  = dt.strftime("%B %d, %Y")
            time_str  = dt.strftime("%I:%M %p")
        except Exception:
            date_str  = start_time[:10]
            time_str  = ""

        dur_m, dur_s = divmod(int(duration), 60)

        sl = sentiment.lower()
        if "positive"  in sl: sent_bg, sent_fg, sent_lbl = BG_GREEN, TX_GREEN, "Positive"
        elif "negative" in sl: sent_bg, sent_fg, sent_lbl = BG_RED,   TX_RED,   "Negative"
        else:                  sent_bg, sent_fg, sent_lbl = BG_GREY,  TX_GREY,  "Neutral"

        dir_bg = BG_BLUE   if direction == "Inbound" else BG_ORANGE
        dir_fg = TX_BLUE   if direction == "Inbound" else HexColor("#9B4A00")

        block = []
        block.append(Paragraph(f"CALL {idx} OF {len(calls_to_show)}", call_num_s))

        title = (f"Inbound call from <b>{from_name}</b>"
                 if direction == "Inbound"
                 else f"Outbound call to <b>{to_name}</b>")
        block.append(Paragraph(title, call_title_s))

        meta_parts = [p for p in [date_str, time_str, f"Duration: {dur_m}m {dur_s}s",
                                  f"From: {from_name}", f"To: {to_name}"] if p]
        block.append(Paragraph(" &nbsp;|&nbsp; ".join(meta_parts), meta_s))

        badges = [[
            Paragraph(f"<b>{direction}</b>", ps("DB", fontSize=8.5, fontName="Helvetica-Bold",
                      textColor=dir_fg, alignment=TA_CENTER)),
            Paragraph(f"<b>{sent_lbl}</b>", ps("SB2", fontSize=8.5, fontName="Helvetica-Bold",
                      textColor=sent_fg, alignment=TA_CENTER)),
            Paragraph("", ps("SP", fontSize=8)),
        ]]
        bt = Table(badges, colWidths=[1.0*inch, 1.0*inch, doc.width - 2.0*inch])
        bt.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (0,0), dir_bg),
            ("BACKGROUND",    (1,0), (1,0), sent_bg),
            ("TOPPADDING",    (0,0), (1,0), 5),
            ("BOTTOMPADDING", (0,0), (1,0), 5),
            ("LEFTPADDING",   (0,0), (1,0), 10),
            ("RIGHTPADDING",  (0,0), (1,0), 10),
        ]))
        block.append(Spacer(1, 4))
        block.append(bt)

        if summary:
            block.append(Paragraph("AI SUMMARY", sum_lbl_s))
            st2 = Table([[Paragraph(summary, sum_txt_s)]], colWidths=[doc.width])
            st2.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), BG_ORANGE),
                ("TOPPADDING",    (0,0), (-1,-1), 10),
                ("BOTTOMPADDING", (0,0), (-1,-1), 10),
                ("LEFTPADDING",   (0,0), (-1,-1), 14),
                ("RIGHTPADDING",  (0,0), (-1,-1), 14),
                ("LINEBEFORE",    (0,0), (0,-1),   3, RC_ORANGE),
            ]))
            block.append(st2)

        if transcript:
            block.append(Paragraph("FULL TRANSCRIPT", tr_lbl_s))
            unique_speakers = list(dict.fromkeys(
                line.split(": ")[0].split("] ")[-1].strip()
                for line in transcript.split("\n") if ": " in line
            ))
            sp_map = {sp: sp_styles[i % len(sp_styles)] for i, sp in enumerate(unique_speakers)}
            for line in transcript.split("\n"):
                if not line.strip():
                    continue
                m = re.match(r"^(\[\d+:\d+\])\s+(.+?):\s+(.+)$", line)
                if m:
                    ts2, speaker, text = m.group(1), m.group(2), m.group(3)
                    safe = text.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                    block.append(Paragraph(
                        f"<b>{speaker}</b> <font size='7' color='#BBBBBB'>{ts2}</font>",
                        sp_map.get(speaker, sp_a_s)))
                    block.append(Paragraph(safe, utt_s))
                else:
                    safe = line.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                    block.append(Paragraph(safe, utt_s))
        else:
            block.append(Spacer(1, 6))
            block.append(Paragraph(
                "No transcript available — extension may not have a RingSense license.", no_tr_s))

        block.append(Spacer(1, 0.15*inch))
        block.append(HRFlowable(width="100%", thickness=0.5, color=RULE))
        block.append(Spacer(1, 0.2*inch))

        story.append(KeepTogether(block[:7]))
        story.extend(block[7:])

    doc.build(story, canvasmaker=NumCanvas)


# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  RingCentral ACE Web App running at http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False)
