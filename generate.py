import pandas as pd
from datetime import datetime
from uuid import uuid4

# Excelファイル内の全シート取得
excel_file = pd.ExcelFile("events.xlsx")

for sheet_name in excel_file.sheet_names:

    # シート読み込み
    df = pd.read_excel("events.xlsx", sheet_name=sheet_name)

    lines = [
        "BEGIN:VCALENDAR",
        "VERSION:2.0",
        f"PRODID:-//{sheet_name} Calendar//JP"
    ]

    for _, row in df.iterrows():

        start = pd.to_datetime(row["start"])
        end = pd.to_datetime(row["end"])

        lines.extend([
            "BEGIN:VEVENT",
            f"UID:{uuid4()}",
            f"DTSTAMP:{datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')}",
            f"DTSTART:{start.strftime('%Y%m%dT%H%M%S')}",
            f"DTEND:{end.strftime('%Y%m%dT%H%M%S')}",
            f"SUMMARY:{row['title']}",
            f"DESCRIPTION:{row['description']}",
            "END:VEVENT"
        ])

    lines.append("END:VCALENDAR")

    # シート名でics生成
    filename = f"{sheet_name}.ics"

    with open(filename, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"{filename} generated!")
