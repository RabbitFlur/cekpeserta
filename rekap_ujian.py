"""Rekap jumlah peserta ujian UT dari file "Daftar Ujian" (.txt) ke file Excel (.xlsx).

Dipakai oleh app.py (Streamlit) dan bisa juga dijalankan lewat command line:

    python rekap_ujian.py daftar-hasil.txt -o rekap.xlsx
"""
from __future__ import annotations

import argparse
import io
import re
import sys
from dataclasses import dataclass, field
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# --------------------------------------------------------------------------
# Model data
# --------------------------------------------------------------------------


@dataclass
class Ruang:
    kode: str                      # mis. "21111.1"
    label: str                     # mis. "1"  -> ditulis "R.1"
    counts: list[int] = field(default_factory=list)      # peserta per jam (dipakai di Excel)
    n_nama: list[int] = field(default_factory=list)      # hitungan dari baris NAMA MHS
    n_nim: list[int] = field(default_factory=list)       # hitungan dari baris NIM (pembanding)
    nomor: list[int] = field(default_factory=list)       # nomor urut baris yang terbaca
    dobel: list[int] = field(default_factory=list)       # nomor urut yang muncul lebih dari sekali
    tanda: dict[int, list[str]] = field(default_factory=dict)  # jam (0-based) -> pesan; -1 = tingkat ruang


@dataclass
class Tabel:
    lokasi: str                    # nama lokasi, mis. "SMAN 1 CIKARANG UTARA"
    hari: int
    tanggal: str
    upbjj: str                     # mis. "21 / Jakarta"
    tempat: str                    # wilayah ujian, mis. "321874 / CIKARANG UTARA"
    n_jam: int = 0
    ruang: dict[str, Ruang] = field(default_factory=dict)

    def daftar_ruang(self) -> list[Ruang]:
        return list(self.ruang.values())


@dataclass
class Hasil:
    tabel: list[Tabel] = field(default_factory=list)
    peringatan: list[str] = field(default_factory=list)

    @property
    def hari(self) -> list[int]:
        return sorted({t.hari for t in self.tabel})


# --------------------------------------------------------------------------
# Membaca file sumber
# --------------------------------------------------------------------------

RE_HARI = re.compile(r"HARI KE-(\d+)\s*\(\s*TGL\s*:\s*(.*?)\s*\)")
RE_UPBJJ = re.compile(r"UT DAERAH\s*:\s*(.+?)\s*$")
RE_KODE = re.compile(r"KODE RUANG\s*:\s*(\S+)")
RE_WILAYAH = re.compile(r"WILAYAH UJIAN\s*:\s*(.+?)\s*$")
RE_RUANG = re.compile(r"^\s*RUANG\s*:\s*(\S+)")
RE_LOKASI = re.compile(r"LOKASI UJIAN\s*:\s*(.+?)\s*$")
RE_JAM = re.compile(r"JAM KE-[A-Z]+")
RE_BARIS = re.compile(r"^\|\s*(\d+)\s*\|")
RE_BARIS_NIM = re.compile(r"^\|\s+\|")
RE_HAL = re.compile(r"^\s*Halaman\s*:\s*(\d+)")


def baca_teks(data: bytes) -> str:
    for enc in ("utf-8", "cp1252"):
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("latin-1", errors="replace")


def _pisah(raw: str) -> str:
    """'21/Jakarta' -> '21 / Jakarta'"""
    if "/" in raw:
        a, b = raw.split("/", 1)
        return f"{a.strip()} / {b.strip()}"
    return raw.strip()


def _sel(line: str, n_jam: int) -> list[str] | None:
    """Pecah satu baris tabel menjadi isi n_jam kolom jam. None jika bentuknya tidak cocok."""
    bagian = line.split("|")
    # '' | NO | sel1 | ... | selN | (sisa)
    if len(bagian) != n_jam + 3:
        return None
    return [b.strip() for b in bagian[2:-1]]


def _rentang(angka: list[int]) -> str:
    angka = sorted(angka)
    out, awal, prev = [], angka[0], angka[0]
    for a in angka[1:] + [None]:
        if a is not None and a == prev + 1:
            prev = a
            continue
        out.append(str(awal) if awal == prev else f"{awal}–{prev}")
        if a is not None:
            awal = prev = a
    return ", ".join(out)


def parse_daftar(teks: str) -> Hasil:
    hasil = Hasil()
    peta: dict[tuple[str, int], Tabel] = {}
    halaman: list[tuple[int, tuple[str, int] | None, str | None]] = []   # (no. halaman, kunci tabel, kunci ruang)
    ctx: dict = {}

    def reset():
        ctx.update(hari=None, tanggal="", upbjj="", kode=None, wilayah="", ruang=None,
                   lokasi_raw=None, n_jam=None)

    reset()
    baris = teks.splitlines()
    for i, line in enumerate(baris):
        no_baris = i + 1

        if "DAFTAR UJIAN" in line:
            reset()
            continue

        m = RE_BARIS.match(line)
        if not m:
            if (mm := RE_HARI.search(line)):
                ctx["hari"], ctx["tanggal"] = int(mm.group(1)), mm.group(2)
            if (mm := RE_UPBJJ.search(line)):
                ctx["upbjj"] = _pisah(mm.group(1))
            if (mm := RE_KODE.search(line)):
                ctx["kode"] = mm.group(1)
            if (mm := RE_WILAYAH.search(line)):
                ctx["wilayah"] = _pisah(mm.group(1))
            if (mm := RE_RUANG.search(line)):
                ctx["ruang"] = mm.group(1)
            if (mm := RE_LOKASI.search(line)):
                ctx["lokasi_raw"] = mm.group(1)
            if "JAM KE-" in line:
                ctx["n_jam"] = len(RE_JAM.findall(line))
            if (mm := RE_HAL.match(line)):
                kt = (ctx["lokasi_raw"], ctx["hari"]) if ctx["lokasi_raw"] and ctx["hari"] else None
                halaman.append((int(mm.group(1)), kt, ctx["kode"] or ctx["ruang"]))
            continue

        # ---- baris peserta (2 baris: NAMA lalu NIM) ----
        no = int(m.group(1))
        if None in (ctx["hari"], ctx["lokasi_raw"], ctx["ruang"], ctx["n_jam"]):
            hasil.peringatan.append(f"Baris {no_baris}: header halaman (hari/lokasi/ruang/jam) tidak terbaca, baris diabaikan.")
            continue

        n = ctx["n_jam"]
        nxt = baris[i + 1] if i + 1 < len(baris) else ""
        nama_sel = _sel(line, n)
        nim_sel = _sel(nxt, n) if RE_BARIS_NIM.match(nxt) else None
        if nama_sel is None or nim_sel is None:
            hasil.peringatan.append(
                f"Baris {no_baris} (hari {ctx['hari']}, ruang {ctx['ruang']}, no. {no}): bentuk baris tidak sesuai, TIDAK dihitung. Periksa manual.")
            continue

        raw = ctx["lokasi_raw"]
        nama_lokasi = raw.split("/", 1)[1].strip() if "/" in raw else raw.strip()
        kunci = (raw, ctx["hari"])
        t = peta.get(kunci)
        if t is None:
            t = Tabel(lokasi=nama_lokasi, hari=ctx["hari"], tanggal=ctx["tanggal"],
                      upbjj=ctx["upbjj"], tempat=ctx["wilayah"], n_jam=n)
            peta[kunci] = t
            hasil.tabel.append(t)

        if n > t.n_jam:
            hasil.peringatan.append(f"{t.lokasi} H{t.hari}: jumlah jam bertambah dari {t.n_jam} menjadi {n} (baris {no_baris}); kolom disesuaikan.")
            for r in t.ruang.values():
                for lst in (r.counts, r.n_nama, r.n_nim):
                    lst.extend([0] * (n - t.n_jam))
            t.n_jam = n
        elif n < t.n_jam:
            hasil.peringatan.append(f"{t.lokasi} H{t.hari}: halaman di baris {no_baris} hanya punya {n} jam (tabel punya {t.n_jam}); sisanya diisi 0.")

        kunci_ruang = ctx["kode"] or ctx["ruang"]
        r = t.ruang.get(kunci_ruang)
        if r is None:
            r = Ruang(kode=kunci_ruang, label=ctx["ruang"],
                      counts=[0] * t.n_jam, n_nama=[0] * t.n_jam, n_nim=[0] * t.n_jam)
            t.ruang[kunci_ruang] = r

        if no in r.nomor:
            r.dobel.append(no)      # dilaporkan sekali per ruang di bawah
            continue
        r.nomor.append(no)

        for j in range(n):
            nama, nim = nama_sel[j], nim_sel[j]
            if nama:
                r.n_nama[j] += 1
            if nim:
                r.n_nim[j] += 1
            if nama or nim:
                r.counts[j] += 1
            if bool(nama) != bool(nim):
                pesan = f"No. {no}: {'NIM' if nama else 'nama'} kosong (nama='{nama}', NIM='{nim}'). Tetap dihitung 1 peserta, mohon dicek."
                r.tanda.setdefault(j, []).append(pesan)
                hasil.peringatan.append(f"{t.lokasi} H{t.hari} R.{r.label} Jam {j + 1}: {pesan}")

    # ---- nomor halaman di file sumber harus berurutan (mendeteksi halaman hilang) ----
    for (n_sebelum, kt, kr), (n_ini, _, _) in zip(halaman, halaman[1:]):
        if n_ini > n_sebelum + 1:
            hilang = _rentang(list(range(n_sebelum + 1, n_ini)))
            pesan = f"Halaman {hilang} tidak ada di file (loncat dari halaman {n_sebelum} ke {n_ini}). Peserta pada halaman itu tidak terhitung."
            hasil.peringatan.append(pesan)
            if kt in peta and kr in peta[kt].ruang:
                peta[kt].ruang[kr].tanda.setdefault(-1, []).append(pesan)
        elif n_ini < n_sebelum and n_ini != 1:
            hasil.peringatan.append(f"Urutan halaman mundur dari {n_sebelum} ke {n_ini}; periksa apakah file sumber utuh.")

    # ---- pemeriksaan setelah semua terbaca ----
    for t in hasil.tabel:
        for r in t.ruang.values():
            if r.dobel:
                pesan = f"Nomor urut {_rentang(r.dobel)} muncul lebih dari sekali di file (halaman ganda?). Baris kedua diabaikan, tidak dihitung dua kali."
                r.tanda.setdefault(-1, []).append(pesan)
                hasil.peringatan.append(f"{t.lokasi} H{t.hari} R.{r.label}: {pesan}")
            if r.nomor:
                hilang = sorted(set(range(1, max(r.nomor) + 1)) - set(r.nomor))
                if hilang:
                    pesan = f"Nomor urut {_rentang(hilang)} tidak ada di file (kemungkinan halaman hilang)."
                    r.tanda.setdefault(-1, []).append(pesan)
                    hasil.peringatan.append(f"{t.lokasi} H{t.hari} R.{r.label}: {pesan}")
        # urutkan ruang secara numerik bila semua label angka
        if all(r.label.isdigit() for r in t.ruang.values()):
            t.ruang = dict(sorted(t.ruang.items(), key=lambda kv: (int(kv[1].label), kv[0])))
            angka = [int(r.label) for r in t.ruang.values()]
            if angka:
                hilang = sorted(set(range(min(angka), max(angka) + 1)) - set(angka))
                if hilang:
                    hasil.peringatan.append(
                        f"{t.lokasi} H{t.hari}: nomor ruang {_rentang(hilang)} tidak ada di file sumber, tidak dimasukkan ke tabel.")
    return hasil


# --------------------------------------------------------------------------
# Validasi & ringkasan
# --------------------------------------------------------------------------


def total_tabel(t: Tabel) -> tuple[list[int], list[int], int]:
    """(total per ruang, total per jam, grand total) — sama seperti rumus SUM di Excel."""
    rs = t.daftar_ruang()
    per_ruang = [sum(r.counts) for r in rs]
    per_jam = [sum(r.counts[j] for r in rs) for j in range(t.n_jam)]
    return per_ruang, per_jam, sum(per_ruang)


def validasi(hasil: Hasil) -> list[dict]:
    out = []
    for t in hasil.tabel:
        per_ruang, per_jam, grand = total_tabel(t)
        rs = t.daftar_ruang()
        dari_nama = sum(sum(r.n_nama) for r in rs)
        dari_nim = sum(sum(r.n_nim) for r in rs)
        sama_total = sum(per_ruang) == sum(per_jam) == grand
        sama_sumber = dari_nama == dari_nim == grand
        out.append(dict(
            lokasi=t.lokasi, hari=t.hari,
            total_ruang=sum(per_ruang), total_jam=sum(per_jam), grand_total=grand,
            hitung_nama=dari_nama, hitung_nim=dari_nim,
            ok=sama_total and sama_sumber,
        ))
    return out


def ringkasan(hasil: Hasil) -> list[dict]:
    out = []
    for t in sorted(hasil.tabel, key=lambda x: (x.hari, x.lokasi)):
        _, _, grand = total_tabel(t)
        out.append({"Lokasi": t.lokasi, "Hari": f"H{t.hari}", "Tanggal": t.tanggal,
                    "Jumlah ruang": len(t.ruang), "Jumlah jam": t.n_jam, "Total peserta-jam": grand})
    return out


def tabel_ke_baris(t: Tabel) -> tuple[list[str], list[list]]:
    """Untuk pratinjau: header + baris (termasuk baris Total Jam)."""
    head = ["Ruang"] + [f"Jam {j + 1}" for j in range(t.n_jam)] + ["Total Ruang"]
    per_ruang, per_jam, grand = total_tabel(t)
    rows = [[f"R.{r.label}", *r.counts, per_ruang[k]] for k, r in enumerate(t.daftar_ruang())]
    rows.append(["Total Jam", *per_jam, grand])
    return head, rows


# --------------------------------------------------------------------------
# Menulis Excel
# --------------------------------------------------------------------------

FONT = "Calibri"
F_NORM = Font(name=FONT, size=11)
F_BOLD = Font(name=FONT, size=11, bold=True)
F_HEAD = Font(name=FONT, size=11, bold=True, color="FFFFFF")
FILL_HEAD = PatternFill("solid", start_color="7A7A7A", end_color="7A7A7A")
FILL_WARN = PatternFill("solid", start_color="FFFF00", end_color="FFFF00")
THIN = Side(style="thin", color="000000")
DOT = Side(style="dotted", color="000000")
TENGAH = Alignment(horizontal="center", vertical="center")
CATATAN = ("Jumlah peserta = jumlah mahasiswa yang terdaftar pada kolom jam tersebut di file Daftar Ujian. "
           "Total Ruang dan Total Jam adalah rumus SUM.")


def _tulis_tabel(ws, r0: int, t: Tabel) -> int:
    """Tulis satu tabel mulai baris r0. Mengembalikan baris pertama setelah tabel."""
    n = t.n_jam
    kol_total = n + 2
    L = get_column_letter

    # blok info (3 baris), mengikuti contoh screenshot
    for k, (lab, val) in enumerate([("UPBJJ - UT", t.upbjj), ("Tempat Ujian", t.tempat), ("Lokasi Ujian", t.lokasi)]):
        ws.cell(r0 + k, 1, lab).font = F_NORM
        ws.cell(r0 + k, 2, f": {val}").font = F_NORM

    r_ban = r0 + 4          # satu baris kosong setelah blok info
    r_jam, r_unit, r_data = r_ban + 1, r_ban + 2, r_ban + 3
    rs = t.daftar_ruang()
    r_akhir = r_data + len(rs) - 1
    r_total = r_akhir + 1

    # judul (banner) + sel "Ruang"
    ws.merge_cells(start_row=r_ban, start_column=2, end_row=r_ban, end_column=kol_total)
    ws.merge_cells(start_row=r_ban, start_column=1, end_row=r_unit, end_column=1)
    for r in range(r_ban, r_unit + 1):
        for c in range(1, kol_total + 1):
            if c == 1 or r == r_ban:
                ws.cell(r, c).fill = FILL_HEAD
    ban = ws.cell(r_ban, 2, f"{t.lokasi} - H{t.hari}")
    ban.font, ban.alignment = F_HEAD, TENGAH
    ban.comment = Comment(CATATAN, "Rekap")
    ruang_h = ws.cell(r_ban, 1, "Ruang")
    ruang_h.font, ruang_h.alignment = F_HEAD, TENGAH

    # header jam + satuan
    judul = [f"Jam {j + 1}" for j in range(n)] + ["Total Ruang"]
    for k, teks in enumerate(judul):
        c = ws.cell(r_jam, 2 + k, teks)
        c.font, c.alignment, c.border = F_BOLD, TENGAH, Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        u = ws.cell(r_unit, 2 + k, "Peserta")
        u.font, u.alignment, u.border = F_NORM, TENGAH, Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    # baris data
    for k, r in enumerate(rs):
        row = r_data + k
        bawah = THIN if row == r_akhir else DOT
        a = ws.cell(row, 1, f"R.{r.label}")
        a.font, a.alignment, a.border = F_NORM, TENGAH, Border(bottom=bawah)
        for j in range(n):
            c = ws.cell(row, 2 + j, r.counts[j])
            c.font, c.alignment = F_NORM, TENGAH
            c.border = Border(left=THIN, bottom=bawah)
        tot = ws.cell(row, kol_total, f"=SUM({L(2)}{row}:{L(n + 1)}{row})")
        tot.font, tot.alignment = F_NORM, TENGAH
        tot.border = Border(left=THIN, right=THIN, bottom=bawah)

        # tandai sel bermasalah
        for j, pesan in r.tanda.items():
            sel = a if j == -1 else ws.cell(row, 2 + j)
            sel.fill = FILL_WARN
            sel.comment = Comment("\n".join(pesan), "Rekap")

    # baris Total Jam
    a = ws.cell(r_total, 1, "Total Jam")
    a.font, a.alignment, a.border = F_NORM, TENGAH, Border(top=THIN, bottom=THIN)
    for j in range(n):
        col = L(2 + j)
        c = ws.cell(r_total, 2 + j, f"=SUM({col}{r_data}:{col}{r_akhir})")
        c.font, c.alignment, c.border = F_NORM, TENGAH, Border(left=THIN, top=THIN, bottom=THIN)
    g = ws.cell(r_total, kol_total, f"=SUM({L(kol_total)}{r_data}:{L(kol_total)}{r_akhir})")
    g.font, g.alignment = F_NORM, TENGAH
    g.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

    return r_total + 2      # satu baris kosong sebelum tabel berikutnya


def buat_workbook(hasil: Hasil) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    wb.calculation.fullCalcOnLoad = True      # Excel menghitung ulang semua rumus saat file dibuka
    for hari in hasil.hari:                       # hanya hari yang ada di file sumber
        ws = wb.create_sheet(f"H{hari}")
        tabel_hari = [t for t in hasil.tabel if t.hari == hari and t.ruang]
        r = 1
        for t in tabel_hari:
            r = _tulis_tabel(ws, r, t)
        maks = max((t.n_jam for t in tabel_hari), default=5)
        ws.column_dimensions["A"].width = 22
        for j in range(maks):
            ws.column_dimensions[get_column_letter(2 + j)].width = 9
        ws.column_dimensions[get_column_letter(maks + 2)].width = 13
        ws.page_setup.orientation = "portrait"
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    return wb


def nama_berkas(hasil: Hasil) -> str:
    """Nama file hasil = nama lokasi ujian (bila ada beberapa lokasi: lokasi pertama + jumlah lainnya)."""
    lokasi = list(dict.fromkeys(t.lokasi for t in hasil.tabel))
    if not lokasi:
        dasar = "rekap_peserta_ujian"
    elif len(lokasi) == 1:
        dasar = lokasi[0]
    else:
        dasar = f"{lokasi[0]} dan {len(lokasi) - 1} lokasi lain"
    dasar = re.sub(r'[\\/:*?"<>|\x00-\x1f]', " ", dasar)      # karakter terlarang di nama file Windows
    dasar = re.sub(r"\s+", " ", dasar).strip(" .")[:120]
    return (dasar or "rekap_peserta_ujian") + ".xlsx"


def buat_xlsx_bytes(hasil: Hasil) -> bytes:
    buf = io.BytesIO()
    buat_workbook(hasil).save(buf)
    return buf.getvalue()


# --------------------------------------------------------------------------
# Command line
# --------------------------------------------------------------------------


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(description="Rekap jumlah peserta ujian UT ke Excel.")
    ap.add_argument("berkas", help="file Daftar Ujian (.txt)")
    ap.add_argument("-o", "--output", help="nama file .xlsx hasil (default: nama lokasi ujian, di folder yang sama dengan file sumber)")
    a = ap.parse_args(argv)

    sumber = Path(a.berkas)
    hasil = parse_daftar(baca_teks(sumber.read_bytes()))
    if not hasil.tabel:
        print("Tidak ada data peserta yang terbaca. Pastikan ini file Daftar Ujian yang benar.", file=sys.stderr)
        return 1

    print("Ringkasan temuan:")
    for r in ringkasan(hasil):
        print(f"  {r['Lokasi']} | {r['Hari']} ({r['Tanggal']}) | {r['Jumlah ruang']} ruang | "
              f"{r['Jumlah jam']} jam | total {r['Total peserta-jam']}")
    print("Validasi:")
    for v in validasi(hasil):
        status = "OK" if v["ok"] else "SELISIH"
        print(f"  {v['lokasi']} H{v['hari']}: Total Ruang={v['total_ruang']} Total Jam={v['total_jam']} "
              f"Grand={v['grand_total']} (nama={v['hitung_nama']}, NIM={v['hitung_nim']}) -> {status}")
    if hasil.peringatan:
        print(f"Peringatan ({len(hasil.peringatan)}):")
        for p in hasil.peringatan:
            print("  -", p)

    out = Path(a.output) if a.output else sumber.with_name(nama_berkas(hasil))
    out.write_bytes(buat_xlsx_bytes(hasil))
    print(f"Tersimpan: {out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
