"""Fungsi serverless Vercel: terima file Daftar Ujian (.txt), balas JSON berisi ringkasan, validasi,
pratinjau, dan file Excel (base64).

POST /api/rekap   body = isi file .txt (boleh di-gzip oleh browser, dikenali dari byte awal 1f 8b)
"""
import base64
import json
import os
import sys
import traceback
import zlib
from http.server import BaseHTTPRequestHandler

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from rekap_ujian import (baca_teks, buat_xlsx_bytes, nama_berkas, parse_daftar,  # noqa: E402
                         ringkasan, tabel_ke_baris, validasi)

BATAS_UPLOAD = 4_500_000          # batas body request fungsi Vercel (~4,5 MB)
BATAS_TEKS = 60 * 1024 * 1024     # batas ukuran teks setelah dekompresi (cegah gzip bomb)


def proses(data: bytes) -> tuple[int, dict]:
    if data[:2] == b"\x1f\x8b":                       # gzip dari browser
        try:
            d = zlib.decompressobj(16 + zlib.MAX_WBITS)
            data = d.decompress(data, BATAS_TEKS + 1)
        except zlib.error:
            return 400, {"error": "File terkompresi tidak bisa dibaca. Coba unggah ulang."}
    if len(data) > BATAS_TEKS:
        return 413, {"error": "File terlalu besar untuk diproses."}

    hasil = parse_daftar(baca_teks(data))
    if not hasil.tabel:
        return 422, {"error": "Tidak ada data peserta yang terbaca. Pastikan file yang diunggah adalah "
                              "Daftar Ujian dengan format yang sama seperti biasanya."}

    pratinjau = []
    for t in hasil.tabel:
        head, rows = tabel_ke_baris(t)
        pratinjau.append({"lokasi": t.lokasi, "hari": t.hari, "tanggal": t.tanggal, "head": head, "rows": rows})

    return 200, {
        "namaBerkas": nama_berkas(hasil),
        "ringkasan": ringkasan(hasil),
        "peringatan": hasil.peringatan[:200],
        "jumlahPeringatan": len(hasil.peringatan),
        "validasi": validasi(hasil),
        "pratinjau": pratinjau,
        "xlsx": base64.b64encode(buat_xlsx_bytes(hasil)).decode("ascii"),
    }


class handler(BaseHTTPRequestHandler):
    def _kirim(self, kode: int, isi: dict):
        badan = json.dumps(isi, ensure_ascii=False).encode("utf-8")
        self.send_response(kode)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(badan)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(badan)

    def do_POST(self):
        try:
            n = int(self.headers.get("content-length") or 0)
            if n <= 0:
                return self._kirim(400, {"error": "File kosong."})
            if n > BATAS_UPLOAD:
                return self._kirim(413, {"error": "File terlalu besar (batas Vercel sekitar 4,5 MB)."})
            kode, isi = proses(self.rfile.read(n))
            self._kirim(kode, isi)
        except Exception:                                   # jangan bocorkan detail ke pengguna
            traceback.print_exc()
            self._kirim(500, {"error": "Terjadi kesalahan saat memproses file."})

    def do_GET(self):
        self._kirim(405, {"error": "Gunakan metode POST."})
