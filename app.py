from flask import Flask, render_template, request, jsonify
import sqlite3
from datetime import datetime, timedelta
from apscheduler.schedulers.background import BackgroundScheduler
from msal import ConfidentialClientApplication
import requests
import random
import logging

from config import *

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger(__name__)

app = Flask(__name__)
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"


# ─────────────────────────────────────────────────────────
#  Auth helper
# ─────────────────────────────────────────────────────────
def get_access_token():
    client = ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
    )
    result = client.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )
    if "access_token" not in result:
        raise Exception("Token error: " + result.get("error_description", "unknown"))
    return result["access_token"]


# ─────────────────────────────────────────────────────────
#  Email validation
# ─────────────────────────────────────────────────────────
def validate_email(access_token, email):
    email = email.strip().lower()
    if not any(email.endswith(d) for d in ALLOWED_DOMAINS):
        return False, f"Domain not allowed: {email}"
    r = requests.get(
        f"https://graph.microsoft.com/v1.0/users/{email}",
        headers={"Authorization": f"Bearer {access_token}"}, timeout=10
    )
    if r.status_code == 200:
        return True, "OK"
    return False, f"User not found in Azure AD: {email}"


# ─────────────────────────────────────────────────────────
#  PIN email
# ─────────────────────────────────────────────────────────
def send_pin_email(access_token, recipient, subject, start, end, pin_code):
    try:
        start_fmt = datetime.fromisoformat(start.replace("Z","")).strftime("%d %b %Y, %I:%M %p")
        end_fmt   = datetime.fromisoformat(end.replace("Z","")).strftime("%I:%M %p")
    except Exception:
        start_fmt, end_fmt = start, end

    body = {
        "message": {
            "subject": f"📋 Conference Room PIN — {subject}",
            "body": {
                "contentType": "HTML",
                "content": f"""
<html><body style="margin:0;padding:0;background:#f4f6f9;font-family:'Segoe UI',Arial,sans-serif;">
<table width="100%" cellpadding="0" cellspacing="0" style="padding:40px 0;">
<tr><td align="center">
<table width="520" cellpadding="0" cellspacing="0"
  style="background:#ffffff;border-radius:16px;overflow:hidden;box-shadow:0 4px 24px rgba(0,0,0,.08);">
  <tr><td style="background:linear-gradient(135deg,#1e3a5f,#2e6da4);padding:36px 40px;text-align:center;">
    <div style="font-size:36px;margin-bottom:8px;">🏢</div>
    <div style="color:#fff;font-size:22px;font-weight:700;">{ROOM_NAME}</div>
    <div style="color:#93c5fd;font-size:14px;margin-top:4px;">Booking Confirmed</div>
  </td></tr>
  <tr><td style="padding:32px 40px;">
    <div style="font-size:18px;font-weight:600;color:#1e3a5f;margin-bottom:6px;">{subject}</div>
    <div style="color:#64748b;font-size:14px;margin-bottom:24px;">⏰ {start_fmt} — {end_fmt}</div>
    <table width="100%" cellpadding="0" cellspacing="0"
      style="background:#f0f7ff;border:2px solid #bfdbfe;border-radius:12px;padding:28px;text-align:center;">
      <tr><td>
        <div style="font-size:11px;text-transform:uppercase;letter-spacing:2px;color:#64748b;margin-bottom:12px;">
          Your Access PIN
        </div>
        <div style="font-size:52px;font-weight:900;letter-spacing:14px;color:#1e3a5f;
                    font-family:'Courier New',monospace;">{pin_code}</div>
        <div style="font-size:12px;color:#94a3b8;margin-top:10px;">
          Enter this on the room display to confirm access
        </div>
      </td></tr>
    </table>
    <div style="margin-top:20px;padding:14px 16px;background:#fefce8;border-radius:8px;
                border-left:4px solid #f59e0b;font-size:13px;color:#78350f;">
      ⚠️ PIN is valid from 10 minutes before your meeting starts.
    </div>
  </td></tr>
  <tr><td style="padding:12px 40px 24px;text-align:center;">
    <div style="font-size:11px;color:#cbd5e1;">
      Automated message from {ROOM_NAME} system.
    </div>
  </td></tr>
</table>
</td></tr>
</table>
</body></html>"""
            },
            "toRecipients": [{"emailAddress": {"address": recipient}}]
        }
    }
    r = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{ROOM_EMAIL}/sendMail",
        headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
        json=body, timeout=15
    )
    if r.status_code in (200, 202):
        logger.info(f"PIN email sent → {recipient}")
    else:
        logger.error(f"PIN email failed {r.status_code}: {r.text[:200]}")


# ─────────────────────────────────────────────────────────
#  Sync  ← KEY FIX: read room calendar, deduplicate by
#          normalised start+end+subject fingerprint so the
#          same physical meeting never gets two DB rows
#          even if Graph returns it via different event_ids.
# ─────────────────────────────────────────────────────────
def sync_meetings():
    logger.info("Calendar sync starting…")
    try:
        token = get_access_token()

        now_utc  = datetime.utcnow()
        start_dt = now_utc.strftime("%Y-%m-%dT00:00:00Z")
        end_dt   = (now_utc + timedelta(days=30)).strftime("%Y-%m-%dT00:00:00Z")

        # Read the ROOM calendar (not a personal calendar) — this is the authoritative source
        url = (
            f"https://graph.microsoft.com/v1.0/users/{ROOM_EMAIL}/calendarView"
            f"?startDateTime={start_dt}&endDateTime={end_dt}"
            f"&$select=id,subject,organizer,start,end,isCancelled"
            f"&$top=100&$orderby=start/dateTime"
        )
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=15)
        events = r.json().get("value", [])
        logger.info(f"Graph returned {len(events)} events")

        conn   = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()

        # Collect event_ids that are still live in Graph
        live_event_ids = set()

        for ev in events:
            eid             = ev.get("id", "")
            is_cancelled    = ev.get("isCancelled", False)
            subject         = ev.get("subject", "No Subject").strip()
            organizer_name  = ev.get("organizer", {}).get("emailAddress", {}).get("name", "Unknown")
            organizer_email = ev.get("organizer", {}).get("emailAddress", {}).get("address", "").lower()
            start_str       = ev.get("start", {}).get("dateTime", "")
            end_str         = ev.get("end",   {}).get("dateTime", "")

            if is_cancelled:
                cursor.execute(
                    "UPDATE meetings SET status='CANCELLED' WHERE event_id=?", (eid,)
                )
                continue

            live_event_ids.add(eid)

            # ── Dedup fingerprint: same subject + same start minute = same meeting
            # Strip microseconds so "2026-05-27T09:43:00.0000000" → "2026-05-27T09:43"
            start_fp = start_str[:16]   # "YYYY-MM-DDTHH:MM"

            # Check if we already have a row with THIS event_id
            cursor.execute(
                "SELECT pin_code, status FROM meetings WHERE event_id=?", (eid,)
            )
            row_by_id = cursor.fetchone()

            if row_by_id:
                # Already tracked — just update mutable fields, preserve PIN & status
                if row_by_id[1] not in ("CANCELLED", "MANUAL"):
                    cursor.execute("""
                        UPDATE meetings
                        SET subject=?, organizer=?, organizer_email=?,
                            start_time=?, end_time=?, status='RESERVED'
                        WHERE event_id=?
                    """, (subject, organizer_name, organizer_email,
                          start_str, end_str, eid))
                continue

            # Check by fingerprint — did we already store a different event_id
            # for the SAME physical meeting? (happens when the same meeting
            # appears on both the organiser's and the room's calendar)
            cursor.execute("""
                SELECT event_id, pin_code FROM meetings
                WHERE substr(start_time,1,16)=?
                  AND status IN ('RESERVED','MANUAL')
            """, (start_fp,))
            row_by_fp = cursor.fetchone()

            if row_by_fp:
                # Duplicate event — just update the event_id to the room-calendar one,
                # keep existing PIN (no new email)
                logger.info(f"Dedup: updating event_id for '{subject}' at {start_fp}")
                cursor.execute(
                    "UPDATE meetings SET event_id=?, subject=?, organizer=?, "
                    "organizer_email=?, end_time=? WHERE event_id=?",
                    (eid, subject, organizer_name, organizer_email,
                     end_str, row_by_fp[0])
                )
                continue

            # Truly new meeting — generate PIN and send email
            pin = str(random.randint(100000, 999999))
            send_pin_email(token, organizer_email, subject, start_str, end_str, pin)
            cursor.execute("""
                INSERT INTO meetings
                    (event_id, subject, organizer, organizer_email,
                     start_time, end_time, pin_code, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, 'RESERVED')
            """, (eid, subject, organizer_name, organizer_email,
                  start_str, end_str, pin))
            logger.info(f"New: '{subject}' | PIN→{organizer_email}")

        # ── Delete rows that no longer exist in Graph
        #    (only for RESERVED ones — keep MANUAL bookings made through our form)
        cursor.execute(
            "SELECT event_id FROM meetings WHERE status='RESERVED'"
        )
        db_ids = {row[0] for row in cursor.fetchall()}
        stale  = db_ids - live_event_ids
        if stale:
            for sid in stale:
                cursor.execute("DELETE FROM meetings WHERE event_id=?", (sid,))
            logger.info(f"Cleaned {len(stale)} stale meetings")

        conn.commit()
        conn.close()
        logger.info("Sync complete")

    except Exception as e:
        logger.error(f"Sync error: {e}", exc_info=True)


# ─────────────────────────────────────────────────────────
#  Routes
# ─────────────────────────────────────────────────────────
@app.route("/")
def home():
    return render_template("display.html", room_name=ROOM_NAME, room_capacity=ROOM_CAPACITY)


@app.route("/api/status")
def api_status():
    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    now    = datetime.utcnow()

    # Only meetings that are active or in the future, and NOT cancelled
    cursor.execute("""
        SELECT subject, organizer, organizer_email, start_time, end_time, pin_code, status
        FROM meetings
        WHERE status IN ('RESERVED', 'MANUAL')
          AND end_time > ?
        ORDER BY start_time ASC
    """, (now.strftime("%Y-%m-%dT%H:%M:%S"),))
    rows = cursor.fetchall()
    conn.close()

    current  = None
    upcoming = []

    for row in rows:
        subject, organizer, _, start_str, end_str, _, _ = row
        try:
            s = datetime.fromisoformat(start_str[:19])
            e = datetime.fromisoformat(end_str[:19])
        except Exception:
            continue

        if s <= now <= e and current is None:
            current = {
                "subject":   subject,
                "organizer": organizer,
                "start":     s.strftime("%I:%M %p"),
                "end":       e.strftime("%I:%M %p"),
                "start_iso": start_str,
                "end_iso":   end_str,
            }
        elif s > now:
            upcoming.append({
                "subject":   subject,
                "organizer": organizer,
                "start":     s.strftime("%I:%M %p"),
                "end":       e.strftime("%I:%M %p"),
                "date":      s.strftime("%d %b"),
                "is_today":  s.date() == now.date(),
            })

    return jsonify({
        "room_name":     ROOM_NAME,
        "room_capacity": ROOM_CAPACITY,
        "current":       current,
        "upcoming":      upcoming[:5],
        "server_time":   now.strftime("%Y-%m-%dT%H:%M:%SZ"),
    })


@app.route("/api/verify_pin", methods=["POST"])
def verify_pin():
    data    = request.get_json(force=True)
    entered = (data.get("pin") or "").strip()
    now     = datetime.utcnow()

    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT subject, organizer, start_time, end_time
        FROM meetings
        WHERE pin_code=? AND status IN ('RESERVED','MANUAL')
        ORDER BY start_time ASC
    """, (entered,))
    row = cursor.fetchone()
    conn.close()

    if not row:
        return jsonify({"ok": False, "message": "Invalid PIN"})

    subject, organizer, start_str, end_str = row
    try:
        s = datetime.fromisoformat(start_str[:19])
        e = datetime.fromisoformat(end_str[:19])
    except Exception:
        return jsonify({"ok": False, "message": "Bad meeting data"})

    if not (s - timedelta(minutes=10) <= now <= e):
        return jsonify({"ok": False, "message": "PIN not valid at this time"})

    return jsonify({
        "ok": True, "subject": subject, "organizer": organizer,
        "start": s.strftime("%I:%M %p"), "end": e.strftime("%I:%M %p"),
    })


@app.route("/book", methods=["GET", "POST"])
def book_meeting():
    if request.method == "GET":
        return render_template("booking.html", room_name=ROOM_NAME)

    subject       = request.form.get("subject", "").strip()
    email         = request.form.get("email",   "").strip().lower()
    attendees_raw = request.form.get("attendees", "").strip()
    start         = request.form.get("start", "").strip()   # YYYY-MM-DDTHH:MM
    end           = request.form.get("end",   "").strip()

    def err(msg):
        return render_template("booking.html", room_name=ROOM_NAME, error=msg,
                               form=request.form)

    if not all([subject, email, start, end]):
        return err("Please fill in all required fields.")
    if start >= end:
        return err("End time must be after start time.")

    try:
        token = get_access_token()
    except Exception as e:
        return err(f"Auth error: {e}")

    valid, msg = validate_email(token, email)
    if not valid:
        return err(f"Organiser email: {msg}")

    attendees = [{"emailAddress": {"address": email}, "type": "required"}]
    if attendees_raw:
        for att in [a.strip().lower() for a in attendees_raw.split(",") if a.strip()]:
            valid, msg = validate_email(token, att)
            if not valid:
                return err(f"Attendee {att}: {msg}")
            attendees.append({"emailAddress": {"address": att}, "type": "required"})

    # Clash check (compare in IST-aware ISO strings)
    start_iso = start + ":00"
    end_iso   = end   + ":00"
    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT subject FROM meetings
        WHERE status IN ('RESERVED','MANUAL')
          AND start_time < ? AND end_time > ?
    """, (end_iso, start_iso))
    clash = cursor.fetchone()
    conn.close()
    if clash:
        return err(f"Room already booked: '{clash[0]}' overlaps your time slot.")

    payload = {
        "subject": subject,
        "body": {"contentType": "HTML",
                 "content": f"<p>Booked via <strong>{ROOM_NAME}</strong> room system.</p>"},
        "start": {"dateTime": start_iso, "timeZone": "Asia/Kolkata"},
        "end":   {"dateTime": end_iso,   "timeZone": "Asia/Kolkata"},
        "location": {"displayName": ROOM_NAME},
        "attendees": attendees + [
            {"emailAddress": {"address": ROOM_EMAIL}, "type": "resource"}
        ],
        "isOnlineMeeting": True,
        "onlineMeetingProvider": "teamsForBusiness",
    }

    r = requests.post(
        f"https://graph.microsoft.com/v1.0/users/{email}/calendar/events",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json=payload, timeout=15
    )

    if r.status_code == 201:
        event_id = r.json().get("id")
        pin      = str(random.randint(100000, 999999))

        conn   = sqlite3.connect(DB_PATH)
        cursor = conn.cursor()
        cursor.execute("""
            INSERT OR IGNORE INTO meetings
                (event_id, subject, organizer, organizer_email,
                 start_time, end_time, pin_code, status)
            VALUES (?, ?, ?, ?, ?, ?, ?, 'MANUAL')
        """, (event_id, subject, email, email, start_iso, end_iso, pin))
        conn.commit()
        conn.close()

        send_pin_email(token, email, subject, start_iso, end_iso, pin)
        logger.info(f"Manual booking: '{subject}' | {email} | PIN:{pin}")

        return render_template("booking_success.html", room_name=ROOM_NAME,
                               subject=subject, organizer=email,
                               start=start.replace("T", " "),
                               end=end.replace("T", " "), pin_code=pin)
    else:
        err_msg = r.json().get("error", {}).get("message", r.text)
        logger.error(f"Graph create failed {r.status_code}: {err_msg}")
        return err(f"Booking failed ({r.status_code}): {err_msg}")


@app.route("/admin")
def admin():
    conn   = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()
    cursor.execute("""
        SELECT subject, organizer, organizer_email, start_time, end_time, pin_code, status
        FROM meetings ORDER BY start_time DESC
    """)
    rows = cursor.fetchall()
    conn.close()
    return render_template("admin.html", meetings=rows, room_name=ROOM_NAME)


# ─────────────────────────────────────────────────────────
#  Scheduler
# ─────────────────────────────────────────────────────────
scheduler = BackgroundScheduler()
scheduler.add_job(sync_meetings, "interval", minutes=1)
scheduler.start()

if __name__ == "__main__":
    sync_meetings()
    app.run(host="0.0.0.0", port=5000, debug=False)