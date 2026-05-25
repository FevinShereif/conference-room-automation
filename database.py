import sqlite3

conn = sqlite3.connect("room.db")

cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS meetings (
    id INTEGER PRIMARY KEY AUTOINCREMENT,

    event_id TEXT UNIQUE,

    subject TEXT,

    organizer TEXT,

    organizer_email TEXT,

    start_time TEXT,

    end_time TEXT,

    pin_code TEXT,

    status TEXT
)
""")

conn.commit()

print("Database initialized successfully")

conn.close()