from flask import Flask
import sqlite3
from datetime import datetime, timezone
from apscheduler.schedulers.background import BackgroundScheduler
from msal import ConfidentialClientApplication
import requests
import random

from config import *

app = Flask(__name__)

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"


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

        subject = event.get("subject", "No Subject")

        organizer = event.get(
            "organizer", {}
        ).get(
            "emailAddress", {}
        ).get(
            "name", "Unknown"
        )

        start_time = event.get(
            "start", {}
        ).get(
            "dateTime", ""
        )

        end_time = event.get(
            "end", {}
        ).get(
            "dateTime", ""
        )

        cursor.execute(
            "SELECT * FROM meetings WHERE event_id=?",
            (event_id,)
        )

        existing = cursor.fetchone()

        if not existing:

            pin_code = str(
                random.randint(100000, 999999)
            )

            cursor.execute("""
            INSERT INTO meetings (
                event_id,
                subject,
                organizer,
                start_time,
                end_time,
                pin_code,
                status
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                event_id,
                subject,
                organizer,
                start_time,
                end_time,
                pin_code,
                "RESERVED"
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
        start_time,
        end_time,
        pin_code
    FROM meetings
    ORDER BY start_time ASC
    """)

    meetings = cursor.fetchall()

    conn.close()

    now = datetime.now(timezone.utc)

    current_meeting = None
    next_meeting = None

    for meeting in meetings:

        subject = meeting[0]
        organizer = meeting[1]
        start_time = meeting[2]
        end_time = meeting[3]
        pin_code = meeting[4]

        start_dt = datetime.fromisoformat(
    start_time.replace("Z", "+00:00")
)

end_dt = datetime.fromisoformat(
    end_time.replace("Z", "+00:00")
)

        if start_dt <= now <= end_dt:

            current_meeting = {
                "subject": subject,
                "organizer": organizer,
                "start": start_dt.strftime("%I:%M %p"),
                "end": end_dt.strftime("%I:%M %p"),
                "pin": pin_code
            }

        elif start_dt > now and next_meeting is None:

            next_meeting = {
                "subject": subject,
                "start": start_dt.strftime("%I:%M %p")
            }

    current_time = datetime.now().strftime(
        "%d-%b-%Y %I:%M %p"
    )

    if current_meeting:

        next_html = ""

        if next_meeting:

            next_html = f"""
            <div class="card">
                <h2>Next Meeting</h2>

                <p>{next_meeting['subject']}</p>

                <p>Starts At:
                {next_meeting['start']}</p>
            </div>
            """

        return f"""
        <html>

        <head>

            <title>
                Smart Conference Room
            </title>

            <meta http-equiv="refresh" content="30">

            <style>

                body {{

                    background:
                    linear-gradient(
                        135deg,
                        #0f172a,
                        #111827,
                        #1e293b
                    );

                    color:white;

                    font-family:Arial;

                    text-align:center;

                    padding:30px;
                }}

                .card {{

                    background:rgba(
                        255,
                        255,
                        255,
                        0.08
                    );

                    padding:25px;

                    margin:20px auto;

                    border-radius:20px;

                    width:500px;

                    box-shadow:0 0 20px rgba(
                        0,
                        0,
                        0,
                        0.4
                    );
                }}

                .pin {{

                    font-size:60px;

                    color:#facc15;

                    font-weight:bold;
                }}

            </style>

        </head>

        <body>

            <h1>
                Conference Room 1
            </h1>

            <div class="card">

                <h2 style="color:#ef4444;">
                    🔴 RESERVED
                </h2>

                <h2>
                    {current_meeting['subject']}
                </h2>

                <p>
                    Organizer:
                    {current_meeting['organizer']}
                </p>

                <p>
                    {current_meeting['start']}
                    -
                    {current_meeting['end']}
                </p>

                <div class="pin">
                    {current_meeting['pin']}
                </div>

            </div>

            {next_html}

            <div class="card">

                <h3>
                    Current Time
                </h3>

                <p>
                    {current_time}
                </p>

            </div>

        </body>

        </html>
        """

    else:

        return f"""
        <html>

        <head>

            <meta http-equiv="refresh" content="30">

            <style>

                body {{

                    background:
                    linear-gradient(
                        135deg,
                        #052e16,
                        #14532d,
                        #166534
                    );

                    color:white;

                    font-family:Arial;

                    text-align:center;

                    padding-top:100px;
                }}

                .card {{

                    background:rgba(
                        255,
                        255,
                        255,
                        0.08
                    );

                    padding:25px;

                    margin:auto;

                    border-radius:20px;

                    width:500px;
                }}

            </style>

        </head>

        <body>

            <div class="card">

                <h1>
                    Conference Room 1
                </h1>

                <h1 style="color:#4ade80;">
                    🟢 AVAILABLE
                </h1>

                <h3>
                    Available Now
                </h3>

                <p>
                    Current Time:
                    {current_time}
                </p>

            </div>

        </body>

        </html>
        """


if __name__ == "__main__":

    sync_meetings()

    app.run(debug=True)