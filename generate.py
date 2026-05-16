import pandas as pd
from datetime import datetime
from uuid import uuid4

df = pd.read_excel("events.xlsx")

calendar = []
calendar.append("BEGIN:VCALENDAR")
calendar.append("VERSION:2.0")
calendar.append("PRODID:-//School Calendar//JP")

for _, row in df.iterrows():
    uid = str(uuid4())

    start = pd.to_datetime(row["start"])
    end = pd.to_datetime(row["end"])

    start_str = start.strftime("%Y%m%dT%H%M%S")
    end_str = end.strftime("%Y%m%dT%H%M%S")
    now_str = datetime.utcnow().strftime("%Y%m%dT%H%M%SZ")

    calendar.extend([
        "BEGIN:VEVENT",
        f"UID:{uid}",
        f"DTSTAMP:{now_str}",
        f"DTSTART;TZID=Asia/Tokyo:{start_str}",
        f"DTEND;TZID=Asia/Tokyo:{end_str}",
        f"SUMMARY:{row['title']}",
        f"DESCRIPTION:{row['description']}",
        "END:VEVENT"
    ])

calendar.append("END:VCALENDAR")

with open("calendar.ics", "w", encoding="utf-8") as f:
    f.write("\n".join(calendar))

print("calendar.ics generated!")
