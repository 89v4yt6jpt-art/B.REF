#!/usr/bin/env python3
"""Convert the source B.League referee workbook into the site's data.json."""
import json
from pathlib import Path
from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "source.xlsx"
OUTPUT = ROOT / "data.json"
LEAGUES = ["B.PREMIER", "B.ONE", "B.NEXT"]
SHEETS = ["B.PREMIER", "B.ONE", "B.NEXT"]

def clean(value):
    if value is None:
        return ""
    return str(value).strip()

def date_text(value):
    if value is None:
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%Y.%m.%d")
    s = clean(value).replace("/", ".").replace("-", ".")
    # Preserve existing YYYY.MM.DD format; normalize numeric date strings if possible.
    parts = s.split(".")
    if len(parts) == 3 and len(parts[0]) == 4:
        return f"{parts[0]}.{int(parts[1]):02d}.{int(parts[2]):02d}"
    return s

def build():
    if not SOURCE.exists():
        raise FileNotFoundError(f"Excel file not found: {SOURCE}")
    wb = load_workbook(SOURCE, data_only=True, read_only=True)
    if "担当試合数リスト" not in wb.sheetnames:
        raise ValueError("シート『担当試合数リスト』が見つかりません")
    ws = wb["担当試合数リスト"]
    referees = []
    # Summary sheet columns: number/name/affiliation, then league totals at F/J/N.
    total_cols = {"B.PREMIER": 6, "B.ONE": 10, "B.NEXT": 14}
    for row in ws.iter_rows(min_row=3, values_only=True):
        number, name = clean(row[0] if len(row) > 0 else None), clean(row[1] if len(row) > 1 else None)
        if not number or not name or not number.startswith("#"):
            continue
        totals = {}
        for league, col in total_cols.items():
            value = row[col] if len(row) > col else 0
            try: totals[league] = int(value or 0)
            except (TypeError, ValueError): totals[league] = 0
        totals["合計"] = sum(totals.values())
        referees.append({"number": number, "name": name,
                         "affiliation": clean(row[2] if len(row) > 2 else None),
                         "totals": totals})

    games = []
    for league, sheet_name in zip(LEAGUES, SHEETS):
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        current_date = ""
        for row in ws.iter_rows(min_row=2, values_only=True):
            # Sheet layout: B=date, C=home, D=separator, E=away, F=venue, G:I=CC/U1/U2
            raw_date = row[1] if len(row) > 1 else None
            if raw_date is not None and clean(raw_date):
                current_date = date_text(raw_date)
            home = clean(row[2] if len(row) > 2 else None)
            away = clean(row[4] if len(row) > 4 else None)
            venue = clean(row[5] if len(row) > 5 else None)
            if not home or not away:
                continue
            card = f"{home} vs {away}"
            for offset, role in enumerate(["CC", "U1", "U2"], start=6):
                referee = clean(row[offset] if len(row) > offset else None)
                if referee:
                    games.append({"league": league, "date": current_date, "home": home,
                                  "away": away, "card": card, "venue": venue,
                                  "role": role, "referee": referee})
    data = {"referees": referees, "games": games,
            "sourceNote": "source.xlsx から自動生成。担当試合数リストと各リーグシートを参照。"}
    OUTPUT.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Generated {OUTPUT.name}: {len(referees)} referees, {len(games)} referee-game assignments")
    wb.close()

if __name__ == "__main__":
    build()
