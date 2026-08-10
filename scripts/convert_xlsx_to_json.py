"""
Convert the daily Excel file into one JSON snapshot per month for the dashboard.
Reads two sheets from the same file:
  - Sheet1 (first sheet)      -> achievement/VAS/growth per store
  - "SALES VALUE PER TYPE"    -> sales qty & value per product type, vs last month (optional sheet)

Run from the repo root: python scripts/convert_xlsx_to_json.py

Input : source/dashboard_data.xlsx   (overwrite this file every day/update with your data)
Output: data/<YYYY-MM>.json          (one snapshot per month, auto-generated)
        data/manifest.json           (list of available months, auto-generated)
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
    now = datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()
    return now.month


# ---------------------------------------------------------------------------
# Sheet 1: achievement / VAS / growth per store
# ---------------------------------------------------------------------------
def parse_store_performance(wb):
    ws = wb.worksheets[0]

    total_days = ws.cell(2, 4).value or 30
    to_date_day = ws.cell(2, 5).value or 0
    time_gone_raw = ws.cell(2, 6).value
    time_gone = time_gone_raw if time_gone_raw is not None else (
        to_date_day / total_days if total_days else 0
    )

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
            continue

        stores.append({
            "no": v(r, 1), "area": clean_area(v(r, 2)), "area_raw": v(r, 2),
            "store_id": v(r, 3), "store_name": clean_store_name(v(r, 4)), "manager": v(r, 5),
            "jml_pc": v(r, 6), "target_unit": v(r, 7), "target_value": v(r, 8),
            "ach_unit": v(r, 9), "ach_pct": v(r, 10), "gap_unit": v(r, 11),
            "mio3_unit": v(r, 13), "mio3_pct": v(r, 14), "iqoo_unit": v(r, 15), "iqoo_pct": v(r, 16),
            "ach_value": v(r, 18), "ach_value_pct": v(r, 19), "gap_value": v(r, 20),
            "growth_all_prev": v(r, 35), "growth_all_curr": v(r, 36),
            "growth_all_gap": v(r, 37), "growth_all_pct": v(r, 38),
            "growth_value_prev": v(r, 47), "growth_value_curr": v(r, 48),
            "growth_value_gap": v(r, 49), "growth_value_pct": v(r, 50),
            "vas": vas_block(r),
        })

    return stores, total_row, to_date_day, total_days, time_gone


# ---------------------------------------------------------------------------
# Sheet 2: "SALES VALUE PER TYPE" — optional, only present from Agustus onward
# ---------------------------------------------------------------------------
def find_sales_type_sheet(wb):
    for name in wb.sheetnames:
        if "SALES" in name.upper() and "TYPE" in name.upper():
            return wb[name]
    return None


def parse_sales_per_type(wb):
    ws = find_sales_type_sheet(wb)
    if ws is None:
        return None

    groups = []
    for c in range(6, ws.max_column + 1):
        val = ws.cell(3, c).value
        if val is not None:
            groups.append((c, str(val).strip()))

    total_group = next((c for c, name in groups if name.upper() == "TOTAL"), None)
    type_groups = [(c, name) for c, name in groups if name.upper() != "TOTAL"]
    types = [name for _, name in type_groups]

    def v(row, col):
        val = ws.cell(row, col).value
        return val if val is not None else 0

    def block(row, start_col):
        return {
            "qty_curr": v(row, start_col), "qty_prev": v(row, start_col + 1), "qty_delta": v(row, start_col + 2),
            "value_curr": v(row, start_col + 3), "value_prev": v(row, start_col + 4), "value_delta": v(row, start_col + 5),
        }

    by_store = {}
    network_total = None
    type_totals = {}

    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        if str(a).strip().upper() == "TOTAL":
            for c, name in type_groups:
                type_totals[name] = block(r, c)
            if total_group:
                network_total = block(r, total_group)
            continue
        if not isinstance(a, (int, float)):
            continue

        store_id = v(r, 3)
        by_type = {name: block(r, c) for c, name in type_groups}
        store_total = block(r, total_group) if total_group else None
        by_store[store_id] = {"by_type": by_type, "total": store_total}

    return {"types": types, "type_totals": type_totals, "network_total": network_total, "by_store": by_store}


def main():
    if not INPUT_PATH.exists():
        print(f"ERROR: {INPUT_PATH} tidak ditemukan. Upload file Excel ke source/dashboard_data.xlsx dulu.")
        sys.exit(1)

    DATA_DIR.mkdir(exist_ok=True)

    wb = openpyxl.load_workbook(INPUT_PATH, data_only=True)
    ws_main = wb.worksheets[0]
    title = ws_main.cell(1, 1).value or ""
    period = parse_period(title)
    m_num = month_number(period)
    year = (datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()).year
    period_key = f"{year}-{m_num:02d}"
    period_label = f"{BULAN_ID[m_num]} {year}"

    stores, total_row, to_date_day, total_days, time_gone = parse_store_performance(wb)
    sales_by_type = parse_sales_per_type(wb)

    # gabungkan sales_by_type ke masing-masing store (kalau sheet-nya ada)
    if sales_by_type:
        for s in stores:
            s["sales_by_type"] = sales_by_type["by_store"].get(s["store_id"])

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
        "sales_types": sales_by_type["types"] if sales_by_type else None,
        "sales_type_totals": sales_by_type["type_totals"] if sales_by_type else None,
        "sales_network_total": sales_by_type["network_total"] if sales_by_type else None,
    }

    out_path = DATA_DIR / f"{period_key}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    if MANIFEST_PATH.exists():
        manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    else:
        manifest = {"periods": []}
    existing = {p["key"]: p for p in manifest["periods"]}
    existing[period_key] = {"key": period_key, "label": period_label, "file": f"data/{period_key}.json"}
    manifest["periods"] = sorted(existing.values(), key=lambda p: p["key"])
    with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    sales_note = f", {len(sales_by_type['types'])} tipe produk" if sales_by_type else " (tanpa data sales-per-type)"
    print(f"OK: {len(stores)} toko{sales_note} -> {out_path}")
    print(f"Manifest updated -> {MANIFEST_PATH} ({len(manifest['periods'])} bulan tersedia)")


if __name__ == "__main__":
    main()
