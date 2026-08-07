"""
Convert the "SALES VALUE PER TYPE" sheet inside source/dashboard_data.xlsx
into a per-month JSON snapshot (network totals by product model + per-store
breakdown), and keep a manifest so the dashboard can switch between months.
Mirrors scripts/convert_xlsx_to_json.py and convert_stock_xlsx_to_json.py.

Run from the repo root: python scripts/convert_sales_type_xlsx_to_json.py

Input : source/dashboard_data.xlsx, sheet "SALES VALUE PER TYPE"
        (same file as the Performance sheet — no separate upload needed)
Output: data/sales-type-<YYYY-MM>.json   (one snapshot per month, auto-generated)
        data/sales-type-manifest.json    (list of available months, auto-generated)
Don't edit anything inside data/ by hand — it gets regenerated every run.

The sheet has a multi-row header (product model -> Qty/Value) with model
columns sometimes repeated (e.g. two "Y11d" columns) — all columns sharing
the same (whitespace-normalized) model name are summed together.
"""
import openpyxl
import json
import os
import re
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
SHEET_NAME = "SALES VALUE PER TYPE"
DATA_DIR = REPO_ROOT / "data"
MANIFEST_PATH = DATA_DIR / "sales-type-manifest.json"

BULAN_ID = [
    "", "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]


def format_updated_at():
    now = datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()
    return f"{now.day} {BULAN_ID[now.month]} {now.year}, {now.strftime('%H:%M')} WIB"


def clean_area(s):
    m = re.findall(r"[A-Za-z][A-Za-z\s]*$", str(s))
    return m[-1].strip() if m else str(s)


def clean_store_name(s):
    return re.sub(r"^[A-Za-z0-9]+-", "", str(s)).strip()


def norm_model(s):
    return re.sub(r"\s+", " ", str(s)).strip()


def find_layout(ws):
    header_row = None
    for r in range(1, min(ws.max_row, 8) + 1):
        v = ws.cell(r, 1).value
        if v and str(v).strip().upper() == "NO":
            header_row = r
            break
    if header_row is None:
        raise ValueError(f"Header row 'NO' tidak ditemukan di sheet {ws.title}")

    metric_row = header_row + 1
    data_start = header_row + 2

    # "TOTAL" header is a merged cell spanning 2 columns (Qty, Value) — only
    # the first column carries the text, so locate it and check the next
    # column too rather than requiring "TOTAL" literally in both.
    col_total_qty = col_total_value = None
    for c in range(1, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v and str(v).strip().upper() == "TOTAL":
            for cc in (c, c + 1):
                metric = str(ws.cell(metric_row, cc).value or "").strip().upper()
                if metric == "QTY":
                    col_total_qty = cc
                elif metric == "VALUE":
                    col_total_value = cc
    if not (col_total_qty and col_total_value):
        raise ValueError(f"Kolom TOTAL Qty/Value tidak ditemukan di sheet {ws.title}")

    return header_row, metric_row, data_start, col_total_qty, col_total_value


def build_column_maps(ws, header_row, metric_row):
    col_model = {}
    current_model = None
    for c in range(6, ws.max_column + 1):
        v = ws.cell(header_row, c).value
        if v:
            vs = str(v).strip().upper()
            current_model = None if vs == "TOTAL" else norm_model(v)
        col_model[c] = current_model

    col_metric = {}
    for c in range(6, ws.max_column + 1):
        v = ws.cell(metric_row, c).value
        col_metric[c] = str(v).strip().upper() if v else None

    return col_model, col_metric


def main():
    if not INPUT_PATH.exists():
        print(f"SKIP: {INPUT_PATH} tidak ditemukan.")
        sys.exit(0)

    wb = openpyxl.load_workbook(INPUT_PATH, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        print(f"SKIP: sheet '{SHEET_NAME}' belum ada di {INPUT_PATH.name} — belum ada data sales-per-type untuk diconvert.")
        sys.exit(0)

    ws = wb[SHEET_NAME]
    header_row, metric_row, data_start, col_total_qty, col_total_value = find_layout(ws)
    col_model, col_metric = build_column_maps(ws, header_row, metric_row)

    DATA_DIR.mkdir(exist_ok=True)

    stores = []
    network_totals = {}  # model -> {qty, value}

    for r in range(data_start, ws.max_row + 1):
        store_name_raw = ws.cell(r, 4).value
        if not store_name_raw or "total" in str(ws.cell(r, 1).value or "").lower():
            continue
        if not isinstance(ws.cell(r, 1).value, (int, float)):
            continue

        by_model_acc = {}
        for c in range(6, ws.max_column + 1):
            model = col_model.get(c)
            metric = col_metric.get(c)
            if not model or not metric:
                continue
            val = ws.cell(r, c).value or 0
            acc = by_model_acc.setdefault(model, {"qty": 0, "value": 0})
            if metric == "QTY":
                acc["qty"] += val
            elif metric == "VALUE":
                acc["value"] += val

        by_type = []
        for model, acc in by_model_acc.items():
            if not acc["qty"] and not acc["value"]:
                continue
            by_type.append({"model": model, "qty": acc["qty"], "value": acc["value"]})
            net = network_totals.setdefault(model, {"qty": 0, "value": 0})
            net["qty"] += acc["qty"]
            net["value"] += acc["value"]
        by_type.sort(key=lambda t: -t["value"])

        stores.append({
            "store_id": ws.cell(r, 3).value,
            "store_name": clean_store_name(store_name_raw),
            "area": clean_area(ws.cell(r, 2).value),
            "headstore": ws.cell(r, 5).value,
            "total_qty": ws.cell(r, col_total_qty).value or 0,
            "total_value": ws.cell(r, col_total_value).value or 0,
            "by_type": by_type,
        })

    stores.sort(key=lambda s: s["store_name"])

    totals_by_type = [
        {"model": m, "qty": v["qty"], "value": v["value"]}
        for m, v in network_totals.items()
    ]
    totals_by_type.sort(key=lambda t: -t["value"])

    # period tag: reuse Sheet1's title if present ("DATA SALES ON <MONTH>"),
    # otherwise fall back to the current run month
    period_key = None
    period_label = None
    if "Sheet1" in wb.sheetnames:
        title = str(wb["Sheet1"].cell(1, 1).value or "")
        months_en = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"]
        if " ON " in title.upper():
            mname = title.upper().split(" ON ")[-1].strip()
            if mname in months_en:
                now = datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()
                m_num = months_en.index(mname) + 1
                period_key = f"{now.year}-{m_num:02d}"
                period_label = f"{BULAN_ID[m_num]} {now.year}"
    if not period_key:
        now = datetime.now(JAKARTA) if JAKARTA else datetime.utcnow()
        period_key = f"{now.year}-{now.month:02d}"
        period_label = f"{BULAN_ID[now.month]} {now.year}"

    payload = {
        "period_key": period_key,
        "period_label": period_label,
        "updated_by": os.environ.get("UPDATED_BY", "Michael Fumar"),
        "updated_at": format_updated_at(),
        "totals_by_type": totals_by_type,
        "stores": stores,
    }

    out_path = DATA_DIR / f"sales-type-{period_key}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, ensure_ascii=False)

    if MANIFEST_PATH.exists():
        manifest = json.loads(MANIFEST_PATH.read_text(encoding="utf-8"))
    else:
        manifest = {"periods": []}

    existing = {p["key"]: p for p in manifest["periods"]}
    existing[period_key] = {
        "key": period_key,
        "label": period_label,
        "file": f"data/sales-type-{period_key}.json",
    }
    manifest["periods"] = sorted(existing.values(), key=lambda p: p["key"])

    with open(MANIFEST_PATH, "w", encoding="utf-8") as f:
        json.dump(manifest, f, indent=2, ensure_ascii=False)

    print(f"OK: {len(stores)} toko, {len(totals_by_type)} tipe produk -> {out_path}")
    print(f"Manifest updated -> {MANIFEST_PATH} ({len(manifest['periods'])} bulan tersedia)")


if __name__ == "__main__":
    main()
