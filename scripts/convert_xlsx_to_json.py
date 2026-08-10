"""
Convert the daily store-performance Excel file into a per-month JSON snapshot,
and keep a manifest of all months so the dashboard can switch between them.

Run from the repo root: python scripts/convert_xlsx_to_json.py

Input : source/dashboard_data.xlsx     (overwrite this file every day with your update)
Output: data/<YYYY-MM>.json            (one snapshot per month, auto-generated)
        data/manifest.json             (list of available months, auto-generated)
Don't edit anything inside data/ by hand — it gets regenerated every run.
"""
import openpyxl
import json
import re
import os
import sys
from datetime import datetime
from pathlib import Path

try:
    from zoneinfo import ZoneInfo
    JAKARTA = ZoneInfo("Asia/Jakarta")
except Exception:
    JAKARTA = None

REPO_ROOT = Path(__file__).resolve().parent.parent
INPUT_PATH = REPO_ROOT / "source" / "dashboard_data.xlsx"
DATA_DIR = REPO_ROOT / "data"
MANIFEST_PATH = DATA_DIR / "manifest.json"

BULAN_ID = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]
MONTHS_EN = [
    "JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE",
    "JULY", "AUGUST", "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER",
]


def format_updated_at():
    now = datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()
    return f"{now.day} {BULAN_ID[now.month]} {now.year}, {now.strftime('%H:%M')} WIB"


def clean_area(s):
    m = re.findall(r"[A-Za-z][A-Za-z\s]*$", str(s))
    return m[-1].strip() if m else str(s)


def clean_store_name(s):
    return re.sub(r"^[A-Za-z0-9]+-", "", str(s)).strip()


def parse_period(title):
    title = str(title)
    if " ON " in title.upper():
        return title.upper().split(" ON ")[-1].strip()
    return title.strip()


def month_number(period_name):
    p = period_name.strip().upper()
    if p in MONTHS_EN:
        return MONTHS_EN.index(p) + 1
    # fallback: kalau nama bulan gak dikenali, pakai bulan berjalan
    now = datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()
    return now.month


def main():
    if not INPUT_PATH.exists():
        print(f"ERROR: {INPUT_PATH} tidak ditemukan. Upload file Excel ke source/dashboard_data.xlsx dulu.")
        sys.exit(1)

    DATA_DIR.mkdir(exist_ok=True)

    wb = openpyxl.load_workbook(INPUT_PATH, data_only=True)
    ws = wb.worksheets[0]  # selalu ambil sheet pertama

    # --- pacing info (baris 1-2), posisinya tetap sama tiap bulan ---
    title = ws.cell(1, 1).value or ""
    total_days = ws.cell(2, 4).value or 30   # D2 = TOTAL DATE
    to_date_day = ws.cell(2, 5).value or 0   # E2 = TO DATE
    time_gone_raw = ws.cell(2, 6).value      # F2 = TIME GONE (rasio 0-1)
    time_gone = time_gone_raw if time_gone_raw is not None else (
        to_date_day / total_days if total_days else 0
    )

    period = parse_period(title)
    m_num = month_number(period)
    year = (datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()).year
    period_key = f"{year}-{m_num:02d}"
    period_label = f"{BULAN_ID[m_num]} {year}"

    def v(row, col):
        val = ws.cell(row, col).value
        return val if val is not None else 0

    def vas_block(r):
        return {
            "acc": {"target": v(r, 55), "ach": v(r, 56), "pct": v(r, 57)},
            "qoala": {"target": v(r, 58), "ach": v(r, 59), "pct": v(r, 60)},
            "bca_insurance": {"target": v(r, 61), "ach": v(r, 62), "pct": v(r, 63)},
            "indosat": {"target": v(r, 64), "ach": v(r, 65), "pct": v(r, 66)},
            "telkomsel": {"target": v(r, 67), "ach": v(r, 68), "pct": v(r, 69)},
        }

    stores = []
    total_row = None

    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        if str(a).strip().upper() == "TOTAL":
            total_row = {
                "target_unit": v(r, 7), "target_value": v(r, 8),
                "ach_unit": v(r, 9), "ach_pct": v(r, 10), "gap_unit": v(r, 11),
                "mio3_unit": v(r, 13), "mio3_pct": v(r, 14),
                "iqoo_unit": v(r, 15), "iqoo_pct": v(r, 16),
                "ach_value": v(r, 18), "ach_value_pct": v(r, 19), "gap_value": v(r, 20),
                "growth_all_prev": v(r, 35), "growth_all_curr": v(r, 36),
                "growth_all_gap": v(r, 37), "growth_all_pct": v(r, 38),
                "growth_value_prev": v(r, 47), "growth_value_curr": v(r, 48),
                "growth_value_gap": v(r, 49), "growth_value_pct": v(r, 50),
                "vas": vas_block(r),
            }
            continue
        if not isinstance(a, (int, float)):
            continue  # skip baris kosong/footer lain

        stores.append({
            "no": v(r, 1),
            "area": clean_area(v(r, 2)),
            "area_raw": v(r, 2),
            "store_id": v(r, 3),
            "store_name": clean_store_name(v(r, 4)),
            "manager": v(r, 5),
            "jml_pc": v(r, 6),
            "target_unit": v(r, 7),
            "target_value": v(r, 8),
            "ach_unit": v(r, 9),
            "ach_pct": v(r, 10),
            "gap_unit": v(r, 11),
            "mio3_unit": v(r, 13),
            "mio3_pct": v(r, 14),
            "iqoo_unit": v(r, 15),
            "iqoo_pct": v(r, 16),
            "ach_value": v(r, 18),
            "ach_value_pct": v(r, 19),
            "gap_value": v(r, 20),
            "growth_all_prev": v(r, 35),
            "growth_all_curr": v(r, 36),
            "growth_all_gap": v(r, 37),
            "growth_all_pct": v(r, 38),
            "growth_value_prev": v(r, 47),
            "growth_value_curr": v(r, 48),
            "growth_value_gap": v(r, 49),
            "growth_value_pct": v(r, 50),
            "vas": vas_block(r),
        })

    data = {
        "period": period,
        "period_key": period_key,
        "period_label": period_label,
        "to_date_day": to_date_day,
        "total_days": total_days,
        "time_gone": time_gone,
        "updated_by": os.environ.get("UPDATED_BY", "Michael Fumar"),
        "updated_at": format_updated_at(),
        "stores": stores,
        "total": total_row,
    }

    out_path = DATA_DIR / f"{period_key}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    # --- update manifest (daftar semua bulan yang tersedia) ---
    if MANIFEST_PATH.exists():
        manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    else:
        manifest = {"periods": []}

    existing = {p["key"]: p for p in manifest["periods"]}
    existing[period_key] = {
        "key": period_key,
        "label": period_label,
        "file": f"data/{period_key}.json",
    }
    manifest["periods"] = sorted(existing.values(), key=lambda p: p["key"])

    with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    print(f"OK: {len(stores)} toko diproses -> {out_path}")
    print(f"Manifest updated -> {MANIFEST_PATH} ({len(manifest['periods'])} bulan tersedia)")


if __name__ == "__main__":
    main()
