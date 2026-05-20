from msal import ConfidentialClientApplication
import requests
import sqlite3
import random

from config import *

AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

app = ConfidentialClientApplication(
    CLIENT_ID,
    authority=AUTHORITY,
    client_credential=CLIENT_SECRET
)

token_result = app.acquire_token_for_client(
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

        print(f"Added: {subject}")

conn.commit()

conn.close()

print("Meeting sync completed")