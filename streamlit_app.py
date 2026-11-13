"""Aplikasi Streamlit: rekap jumlah peserta ujian UT dari file Daftar Ujian (.txt) ke Excel."""
import pandas as pd
import streamlit as st

from rekap_ujian import (baca_teks, buat_xlsx_bytes, nama_berkas, parse_daftar, ringkasan,
                         tabel_ke_baris, validasi)

st.set_page_config(page_title="Rekap Peserta Ujian", page_icon="📊", layout="wide")

st.title("Rekap Jumlah Peserta Ujian")
st.caption("Unggah file Daftar Ujian (.txt). Aplikasi menghitung jumlah peserta per ruang per jam, "
           "lalu membuat file Excel (.xlsx) yang siap diunduh.")

berkas = st.file_uploader("Pilih file Daftar Ujian", type=["txt"])
if berkas is None:
    st.info("Belum ada file. Silakan unggah file Daftar Ujian (.txt) di atas.")
    st.stop()

with st.spinner("Membaca file..."):
    hasil = parse_daftar(baca_teks(berkas.getvalue()))

if not hasil.tabel:
    st.error("Tidak ada data peserta yang terbaca. Pastikan file yang diunggah adalah Daftar Ujian "
             "dengan format yang sama seperti biasanya.")
    st.stop()

# ---- 1. Ringkasan temuan ----
st.subheader("1. Ringkasan temuan")
st.dataframe(pd.DataFrame(ringkasan(hasil)), hide_index=True)

# ---- 2. Peringatan ----
if hasil.peringatan:
    st.warning(f"Ada {len(hasil.peringatan)} peringatan. Sel terkait ditandai kuning di file Excel.")
    with st.expander("Lihat peringatan", expanded=True):
        for p in hasil.peringatan[:100]:
            st.write("• " + p)
        if len(hasil.peringatan) > 100:
            st.write(f"... dan {len(hasil.peringatan) - 100} peringatan lainnya.")
else:
    st.success("Tidak ada peringatan: nomor urut dan nomor halaman berurutan, semua sel terisi lengkap.")

# ---- 3. Validasi ----
st.subheader("2. Validasi")
for v in validasi(hasil):
    teks = (f"**{v['lokasi']} – H{v['hari']}**: jumlah Total Ruang = {v['total_ruang']:,}, "
            f"jumlah Total Jam = {v['total_jam']:,}, grand total = {v['grand_total']:,}. "
            f"Hitungan nama = {v['hitung_nama']:,}, hitungan NIM = {v['hitung_nim']:,}.")
    (st.success if v["ok"] else st.error)(("✅ Cocok. " if v["ok"] else "❌ Ada selisih. ") + teks)

# ---- 4. Pratinjau ----
st.subheader("3. Pratinjau")
hari_list = hasil.hari
tabs = st.tabs([f"H{h}" for h in hari_list])
for tab, h in zip(tabs, hari_list):
    with tab:
        for t in [x for x in hasil.tabel if x.hari == h]:
            st.markdown(f"**{t.lokasi} - H{t.hari}**  ({t.tanggal})")
            head, rows = tabel_ke_baris(t)
            st.dataframe(pd.DataFrame(rows, columns=head), hide_index=True)

# ---- 5. Unduh ----
st.subheader("4. Unduh hasil")
st.download_button(
    "⬇️ Unduh rekap (.xlsx)",
    data=buat_xlsx_bytes(hasil),
    file_name=nama_berkas(hasil),
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",
)
