from flask import Flask, render_template, request, redirect
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

    graph_event_ids = []

    for event in events:

        event_id = event.get("id")

        graph_event_ids.append(event_id)

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

        cursor.execute("""
        SELECT pin_code
        FROM meetings
        WHERE event_id=?
        """, (event_id,))

        existing = cursor.fetchone()

        if existing:

            cursor.execute("""
            UPDATE meetings
            SET
                subject=?,
                organizer=?,
                organizer_email=?,
                start_time=?,
                end_time=?
            WHERE event_id=?
            """, (
                subject,
                organizer,
                organizer_email,
                start_time,
                end_time,
                event_id
            ))

        else:

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

    cursor.execute("""
    SELECT event_id
    FROM meetings
    """)

    db_events = cursor.fetchall()

    for db_event in db_events:

        db_event_id = db_event[0]

        if db_event_id not in graph_event_ids:

            cursor.execute("""
            DELETE FROM meetings
            WHERE event_id=?
            """, (db_event_id,))

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
        "%d %b %Y, %I:%M:%S %p"
    )

    return render_template(
        "display.html",
        current_meeting=current_meeting,
        next_meeting=next_meeting,
        current_time=current_time
    )

def validate_email(access_token, email):

    headers = {
        "Authorization": f"Bearer {access_token}"
    }

    url = f"https://graph.microsoft.com/v1.0/users/{email}"

    response = requests.get(
        url,
        headers=headers
    )

    if response.status_code != 200:
        return False

    allowed_domains = [
        "@mahathiinfotech.com",
        "@mahathiinfotech.in"
    ]

    if not any(
        email.endswith(domain)
        for domain in allowed_domains
    ):
        return False

    return True


@app.route("/book", methods=["GET", "POST"])
def book_meeting():

    if request.method == "POST":

        subject = request.form["subject"]

        email = request.form["email"]

        attendees_input = request.form["attendees"]

        start = request.form["start"]

        end = request.form["end"]

        app_msal = ConfidentialClientApplication(
            CLIENT_ID,
            authority=AUTHORITY,
            client_credential=CLIENT_SECRET
        )

        token_result = app_msal.acquire_token_for_client(
            scopes=["https://graph.microsoft.com/.default"]
        )

        access_token = token_result["access_token"]

        if not validate_email(access_token, email):

            return """

            <h1 style='
                color:red;
                text-align:center;
                margin-top:100px;
            '>

                Invalid Organizer Email

            </h1>

            """

        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }

        attendees = []

        attendees.append({

            "emailAddress": {
                "address": email
            },

            "type": "required"
        })

        if attendees_input.strip():

            attendee_list = attendees_input.split(",")

            for attendee in attendee_list:

                attendee = attendee.strip()

                if not validate_email(access_token, attendee):

                    return f"""

                    <h1 style='
                        color:red;
                        text-align:center;
                        margin-top:100px;
                    '>

                        Invalid Attendee:
                        {attendee}

                    </h1>

                    """

                attendees.append({

                    "emailAddress": {
                        "address": attendee
                    },

                    "type": "required"
                })

        conn = sqlite3.connect("room.db")

        cursor = conn.cursor()

        cursor.execute("""

        SELECT *

        FROM meetings

        WHERE (

            start_time <= ?
            AND end_time >= ?

        )

        """, (

            end,
            start

        ))

        existing_booking = cursor.fetchone()

        conn.close()

        if existing_booking:

            return """

            <h1 style='
                color:red;
                text-align:center;
                margin-top:100px;
            '>

                Conference Room Already Reserved
                For Selected Time Slot

            </h1>

            """

        meeting_data = {

            "subject": subject,

            "body": {
                "contentType": "HTML",
                "content": "Conference Room Booking"
            },

            "start": {
                "dateTime": start,
                "timeZone": "Asia/Kolkata"
            },

            "end": {
                "dateTime": end,
                "timeZone": "Asia/Kolkata"
            },

            "location": {
                "displayName": "Conference Room 1"
            },

            "attendees": attendees,

            "isOnlineMeeting": True,

            "onlineMeetingProvider": "teamsForBusiness"
        }

        url = "https://graph.microsoft.com/v1.0/users/fevin.s@mahathiinfotech.com/calendar/events"

        response = requests.post(
            url,
            headers=headers,
            json=meeting_data
        )

        print(response.status_code)
        print(response.text)

        return """

        <h1 style='
            color:#4ade80;
            text-align:center;
            margin-top:100px;
            font-family:Arial;
        '>

            Meeting Created Successfully

        </h1>

        <p style='
            text-align:center;
            color:white;
            font-size:22px;
        '>

            Outlook and Teams Invites Sent

        </p>

        """

    return render_template("booking.html")

if __name__ == "__main__":

    sync_meetings()

    app.run(
        host="0.0.0.0",
        port=5000,
        debug=True
    )