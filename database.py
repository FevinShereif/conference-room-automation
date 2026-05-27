import sqlite3
from config import DB_PATH

conn = sqlite3.connect(DB_PATH)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS meetings (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    event_id        TEXT UNIQUE,
    subject         TEXT,
    organizer       TEXT,
    organizer_email TEXT,
    start_time      TEXT,
    end_time        TEXT,
    pin_code        TEXT,
    status          TEXT DEFAULT 'RESERVED'
)
""")

conn.commit()
conn.close()
print(f"Database ready: {DB_PATH}")
