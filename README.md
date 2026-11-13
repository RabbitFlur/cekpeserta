# Rekap Jumlah Peserta Ujian

Membaca file **Daftar Ujian (.txt)** dan membuat file Excel berisi jumlah peserta per ruang per jam
(sheet `H1`, `H2`, satu tabel per lokasi ujian). **Nama file hasil = nama lokasi ujian**
(contoh: `SMAN 1 CIKARANG UTARA.xlsx`; bila satu file memuat beberapa lokasi: `<lokasi pertama> dan N lokasi lain.xlsx`).

## A. Deploy ke Vercel (versi web)

Struktur: `public/index.html` (halaman) + `api/rekap.py` (fungsi Python) + `rekap_ujian.py` (logika).

**Lewat GitHub (paling mudah)**
1. Unggah isi folder ini ke repositori GitHub.
2. Di vercel.com: **Add New → Project** → pilih repositori → **Deploy**. Tidak ada pengaturan yang perlu diubah.

**Lewat command line**
```
npm i -g vercel
vercel          # pertama kali: ikuti pertanyaan, jawab default
vercel --prod
```

Catatan:
- Batas ukuran request di Vercel sekitar 4,5 MB. Browser mengompres file (gzip) sebelum dikirim, jadi file teks
  jauh lebih besar dari itu biasanya tetap muat. Jika tetap terlalu besar, halaman menampilkan pesan.
- File diproses di memori dan tidak disimpan di server.
- Alamat Vercel bisa dibuka siapa saja yang tahu linknya. Karena file berisi nama dan NIM mahasiswa, aktifkan
  proteksi di **Settings → Deployment Protection** (Vercel Authentication atau Password Protection).

## B. Jalankan di komputer sendiri

Streamlit (dengan antarmuka web lokal):
```
pip install -r requirements-streamlit.txt
streamlit run streamlit_app.py
```

Command line:
```
pip install -r requirements.txt
python rekap_ujian.py daftar-hasil.txt
```

## Aturan hitung

- Jumlah peserta per jam = jumlah mahasiswa yang tercantum pada kolom jam itu (satu sel nama = satu peserta).
- Ruang tanpa peserta pada jam tertentu diisi 0.
- Jumlah kolom jam mengikuti header file (tidak harus 5). Sheet hanya dibuat untuk hari yang ada di file.
- Total Ruang, Total Jam, dan grand total berupa rumus `=SUM(...)`.

## Peringatan (sel ditandai kuning di Excel)

- nama ada tetapi NIM kosong, atau sebaliknya (tetap dihitung 1 peserta);
- nomor urut peserta loncat atau ganda dalam satu ruang;
- nomor halaman di file loncat (halaman hilang);
- baris yang bentuknya tidak sesuai (tidak dihitung).
