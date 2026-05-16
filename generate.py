import pandas as pd
from datetime import datetime
from uuid import uuid4

excel_file = pd.ExcelFile("events.xlsx")

for sheet_name in excel_file.sheet_names:

    df = pd.read_excel("events.xlsx", sheet_name=sheet_name)

    # 列名の空白除去
    df.columns = df.columns.str.strip()

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//{sheet_name} Calendar//JP"
    ]

    for _, row in df.iterrows():

        try:
            start = pd.to_datetime(row["start"])
            end = pd.to_datetime(row["end"])

            title = str(row["title"]) if pd.notna(row["title"]) else "予定"

            description = ""
            if "description" in df.columns and pd.notna(row["description"]):
                description = str(row["description"])

            location = ""
            if "location" in df.columns and pd.notna(row["location"]):
                location = str(row["location"])

            lines.extend([
                "BEGIN:VEVENT",
                f"UID:{uuid4()}",
                f"DTSTAMP:{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}",
                f"DTSTART:{start.strftime('%Y%m%dT%H%M%S')}",
                f"DTEND:{end.strftime('%Y%m%dT%H%M%S')}",
                f"SUMMARY:{title}",
                f"DESCRIPTION:{description}",
                f"LOCATION:{location}",
                "END:VEVENT"
            ])

        except Exception as e:
            print("ERROR ROW:")
            print(row)
            print(e)

    lines.append("END:VCALENDAR")

    filename = f"{sheet_name}.ics"

    with open(filename, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"{filename} generated!")
