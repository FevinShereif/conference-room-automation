from flask import Flask, render_template
import sqlite3
from datetime import datetime
from apscheduler.schedulers.background import BackgroundScheduler
from msal import ConfidentialClientApplication
import requests
import random

from config import *

app = Flask(__name__)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"


def send_pin_email(
    access_token,
    recipient,
    subject,
    meeting_start,
    meeting_end,
    pin_code
):

    email_url = f"https://graph.microsoft.com/v1.0/users/{ROOM_EMAIL}/sendMail"

    email_body = {
        "message": {
            "subject": "Conference Room Access PIN",
            "body": {
                "contentType": "Text",
                "content": f"""
Conference Room Booking Confirmed

Meeting:
{subject}

Time:
{meeting_start} - {meeting_end}

Your Access PIN:
{pin_code}

Please use this PIN on the conference room display to unlock the room.
"""
            },
            "toRecipients": [
                {
                    "emailAddress": {
                        "address": recipient
                    }
                }
            ]
        }
    }

    requests.post(
        email_url,
        headers={
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        },
        json=email_body
    )


def sync_meetings():

    app_msal = ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )

    token_result = app_msal.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    access_token = token_result["access_token"]

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    url = f"https://graph.microsoft.com/v1.0/users/{ROOM_EMAIL}/calendar/events"

    response = requests.get(url, headers=headers)

    events = response.json().get("value", [])

    conn = sqlite3.connect("room.db")

    cursor = conn.cursor()

    for event in events:

        event_id = event.get("id")

        subject = event.get(
            "subject",
            "No Subject"
        )

        organizer = event.get(
            "organizer", {}
        ).get(
            "emailAddress", {}
        ).get(
            "name",
            "Unknown"
        )

        organizer_email = event.get(
            "organizer", {}
        ).get(
            "emailAddress", {}
        ).get(
            "address",
            ""
        )

        start_time = event.get(
            "start", {}
        ).get(
            "dateTime",
            ""
        )

        end_time = event.get(
            "end", {}
        ).get(
            "dateTime",
            ""
        )

        cursor.execute(
            "SELECT pin_code FROM meetings WHERE event_id=?",
            (event_id,)
        )

        existing = cursor.fetchone()

        if not existing:

            pin_code = str(
                random.randint(100000, 999999)
            )

            send_pin_email(
                access_token,
                organizer_email,
                subject,
                start_time,
                end_time,
                pin_code
            )

            cursor.execute("""
            INSERT INTO meetings (
                event_id,
                subject,
                organizer,
                organizer_email,
                start_time,
                end_time,
                pin_code
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                event_id,
                subject,
                organizer,
                organizer_email,
                start_time,
                end_time,
                pin_code
            ))

    conn.commit()

    conn.close()


scheduler = BackgroundScheduler()

scheduler.add_job(
    sync_meetings,
    "interval",
    minutes=1
)

scheduler.start()


@app.route("/")
def home():

    conn = sqlite3.connect("room.db")

    cursor = conn.cursor()

    cursor.execute("""
    SELECT
        subject,
        organizer,
        organizer_email,
        start_time,
        end_time,
        pin_code
    FROM meetings
    ORDER BY start_time ASC
    """)

    meetings = cursor.fetchall()

    conn.close()

    now = datetime.utcnow()

    current_meeting = None
    next_meeting = None

    for meeting in meetings:

        subject = meeting[0]
        organizer = meeting[1]
        organizer_email = meeting[2]
        start_time = meeting[3]
        end_time = meeting[4]

        start_dt = datetime.fromisoformat(
            start_time.replace("Z", "")
        )

        end_dt = datetime.fromisoformat(
            end_time.replace("Z", "")
        )

        if start_dt <= now <= end_dt:

            current_meeting = {
                "subject": subject,
                "organizer": organizer,
                "start": start_dt.strftime("%I:%M %p"),
                "end": end_dt.strftime("%I:%M %p")
            }

        elif start_dt > now and next_meeting is None:

            next_meeting = {
                "subject": subject,
                "start": start_dt.strftime("%I:%M %p")
            }

    current_time = datetime.now().strftime(
        "%d-%b-%Y %I:%M %p"
    )

    return render_template(
        "display.html",
        current_meeting=current_meeting,
        next_meeting=next_meeting,
        current_time=current_time
    )


if __name__ == "__main__":

    sync_meetings()

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )