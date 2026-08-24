"""
Helper validasi ringan yang dipakai oleh semua script convert_*.py.

Tujuannya: kalau struktur Excel tiba-tiba berubah (kolom geser, sheet
berganti nama/hilang, header berubah teks) -- ketahuan lewat pesan JELAS
di log GitHub Actions (tab Actions -> run terakhir -> step "convert"),
bukan diam-diam menghasilkan angka yang salah tanpa ada yang sadar.

Cara pakai singkat:
    from validate import Validator
    v = Validator("Store Performance")
    v.expect_header(ws, 3, 1, "NO")
    v.check(len(stores) >= 30, f"Jumlah toko cuma {len(stores)}, biasanya ~43")
    v.report()
"""


class Validator:
    def __init__(self, name):
        self.name = name
        self.issues = []  # list of (level, message)

    def check(self, condition, message, level="WARN"):
        """Catat masalah kalau condition False. level: 'WARN' atau 'FAIL'."""
        if not condition:
            self.issues.append((level, message))
        return condition

    def expect_header(self, ws, row, col, expected, label=None):
        """Bandingkan teks di satu sel header dengan yang diharapkan
        (case-insensitive, whitespace-trimmed)."""
        actual = ws.cell(row, col).value
        actual_str = str(actual).strip().upper() if actual is not None else "(kosong)"
        expected_str = str(expected).strip().upper()
        ok = actual_str == expected_str
        loc = label or f"baris {row} kolom {col}"
        self.check(
            ok,
            f"Header di {loc} harusnya '{expected}', tapi ketemu '{actual_str}' — "
            f"kemungkinan struktur sheet berubah, cek ulang hasilnya.",
        )
        return ok

    def report(self):
        if not self.issues:
            print(f"[VALIDASI {self.name}] OK — semua pengecekan lolos.")
            return
        fails = [m for lvl, m in self.issues if lvl == "FAIL"]
        warns = [m for lvl, m in self.issues if lvl == "WARN"]
        print(f"[VALIDASI {self.name}] {len(fails)} FAIL, {len(warns)} WARNING ditemukan:")
        for lvl, m in self.issues:
            print(f"  [{lvl}] {m}")
        if fails:
            print(f"[VALIDASI {self.name}] ADA FAIL — data tetap diproses supaya dashboard gak kosong, "
                  f"tapi WAJIB dicek manual, kemungkinan besar ada angka yang salah.")
