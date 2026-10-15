# Daftar Isi Berkas Aktif - Input Surat Masuk

Skrip untuk memasukkan data surat masuk bulanan ke dalam `Daftar Isi Berkas Aktif`,
dibagi per sheet berdasarkan kode klasifikasi surat.

## Folder sumber

Semua berkas surat masuk bulanan diletakkan di folder **`Daftar Berkas Per Bulan`**:

```
Daftar Berkas Per Bulan/
├── Surat Masuk Januari 2026.xls
├── Surat Masuk Februari 2026.xls
├── Surat Masuk Maret 2026.xls
├── surat masuk April 2026.xls
├── Surat masuk Mei 2026.xls
├── surat masuk Juni 2026.xls
├── Surat Masuk Juli 2026.xls
└── Surat Masuk Agustus 2026.xls
```

Bulan dan tahun diambil dari **nama berkas**, jadi penulisan kapitalisasi bebas
(`Surat Masuk` maupun `surat masuk`). Kalau nanti dapat September, cukup taruh
`Surat Masuk September 2026.xls` di folder lalu jalankan ulang.

## Cara pakai

```bash
python build_daftar_isi.py --folder "Daftar Berkas Per Bulan"
```

Semua berkas di folder dibaca, diurutkan kronologis (Januari lebih dulu),
digabung, duplikatnya dibuang, lalu ditulis ke:

```
Daftar Isi Berkas Aktif <Bulan Terakhir> <Tahun>.xlsx
```

Mis. dengan isi folder Januari–Agustus, hasilnya
`Daftar Isi Berkas Aktif Agustus 2026.xlsx`. Nama bisa dipaksa lewat `--keluaran`.

| Argumen | Fungsi |
| --- | --- |
| `--folder` | Folder berisi berkas surat masuk bulanan |
| `--keluaran` | Nama file hasil (opsional) |
| `--template` | File acuan, bawaan `template.xlsx` |

File sumber **harus ditutup** di Excel sebelum dijalankan. Butuh `openpyxl` saja.

### Mode lama (satu berkas / tambah per bulan)

Masih tersedia kalau perlu menambah satu berkas saja:

```bash
python build_daftar_isi.py --sumber "surat masuk Oktober 2026.xls" \
                           --dasar "Daftar Isi Berkas Aktif Juni 2026.xlsx" \
                           --keluaran "Daftar Isi Berkas Aktif Oktober 2026.xlsx" \
                           --bulan Oktober
```

Tanpa argumen, skrip mereproduksi Daftar Isi dari satu sumber bawaan.

## Duplikasi

Dua baris dengan uraian persis sama dianggap duplikat, dan yang lebih dulu
(bulan yang lebih awal) dipertahankan. Baris yang dibuang dicetak di layar beserta
jumlah per berkas asal:

```
Duplikat dibuang 254 baris (uraian persis sama, disimpan yang lebih dulu):
  - [Surat Masuk Januari 2026.xls] PW 0202-Ij/1621
  - [Surat Masuk Februari 2026.xls] PR0201/B/Gd/2026/014
  ...
  asal baris yang dibuang: Agustus=11, Februari=30, Januari=14, Juli=78, ...
```

Kuncinya uraian lengkap (`Surat {Dari} Nomor : {Nomor} {Isi Ringkas}`), bukan
nomor surat saja. Alasannya ada surat yang nomornya sama tapi isinya berbeda,
mis. `UND-059/KSP/D.4/06/2026` yang tercatat dua kali dengan uraian sedikit
berbeda. Kunci uraian persis tidak salah buang data.

Untuk melihat baris apa saja yang dibuang tanpa memotong 15 baris pertama, ubah
`buang[:15]` di fungsi `tampilkan_duplikat()`.

## Struktur file sumber

Header **tidak selalu di baris 1** — sebagian berkas punya judul
"Laporan Surat Masuk" di atasnya (header di baris 4). Skrip mencari sendiri baris
yang memuat `No.` dan `Nomor dan Tanggal Surat`, lalu data dibaca dari bawahnya.

| Kolom | Header | Dipakai untuk |
| --- | --- | --- |
| A | `No.` | urutan asal (pemecah urutan terakhir) |
| B | `Tanggal Penerimaan Surat` | **urutan** (tanggal masuk) |
| C | `Nomor dan Tanggal Surat` | kode, nomor, dan **tanggal surat** |
| D | `Sifat Surat` | kolom H |
| E | `Isi Ringkas` | uraian |
| F | `Dari` | uraian |

Kolom C berisi `NOMOR - TANGGAL`, contoh `UM0102/B/Sj/2026/190 - 2026-05-31 01:05:00`.
Nomor dipakai untuk kode dan uraian, bagian setelah `" - "` dipakai untuk kolom TANGGAL.

## Pemetaan data

| Target | Sumber | Isi |
| --- | --- | --- |
| A `NO. BERKAS` | - | kosong |
| B `NO ITEM BERKAS` | - | kosong |
| C `KODE KLASIFIKASI` | `Nomor dan Tanggal Surat` | segmen sebelum `/`, titik dan spasi diabaikan, huruf besar |
| D `URAIAN INFORMASI` | `Dari` + `Nomor` + `Isi Ringkas` | `Surat {Dari} Nomor : {Nomor} {Isi Ringkas}` |
| E `TANGGAL` | `Nomor dan Tanggal Surat` | tanggal surat, format `dd-mm-yyyy` |
| F `JUMLAH` | konstanta | `1 Berkas` |
| G `KETERANGAN` | konstanta | `Asli` |
| H `KET. KLASIFIKASI KEAMANAN DAN AKSES ARSIP` | `Sifat Surat` | `Biasa` / `Terbatas` / `Rahasia` |

**Normalisasi kode** - titik dan spasi tidak dianggap, jadi beberapa penulisan dianggap sama:

| Nomor surat | Kode |
| --- | --- |
| `UM.01.01/B/Sj/2026/172` | `UM0101` |
| `UM0101/B/Sj/2026/172` | `UM0101` |
| `UM 0102/B/Sj/2026/190` | `UM0102` |

Surat tanpa nomor (`-`) tetap masuk dengan kode `-` di sheet `LAIN-LAIN`.

## Pembagian sheet

Nama sheet = 2 huruf pertama segmen sebelum `/`, kalau berupa huruf.
Contoh `UM1010` -> sheet `UM`. Selain itu masuk sheet `LAIN-LAIN`.

```
UM0102/B/Sj/...        -> sheet UM
KP0602/B/Ge/...        -> sheet KP
400.1.12/433/...       -> sheet LAIN-LAIN   (tidak diawali 2 huruf)
UND-059/KSP/...        -> sheet UN          (UND-059 diawali "UN")
PPE.PP.01.05-2012      -> sheet PP
```

Karena aturannya bersifat harfiah, ada sheet yang muncul dari awalan non-klasifikasi:
`PP` (dari `PPE.`), `AH` (dari `AHU-`), `UN` (dari `UND-`), `AP` (dari `APUD/`),
`BW`, `CM`, `IW`, `LB`, dll. Bisa dipindah ke `LAIN-LAIN` dengan menambah
pengecualian di kode.

**Urutan sheet**: abjad kode, tapi `LAIN-LAIN` sengaja ditaruh **paling bawah**
sebelum sheet rekap.

## Isi tiap sheet

```
baris 1-4    judul berkas (merge A1:H1 ... A4:H4), diambil dari sheet Template
baris 6      TAHUN        : 2026
baris 7      BULAN        : Januari s/d Agustus
baris 9      header kolom A-H
baris 10     baris kuning, berisi judul klasifikasi, mis. "UMUM (UM)"
baris 11+    data surat
```

Label bulan disusun otomatis dari nama berkas yang ada di folder.

**Baris kuning sebagai pemisah** - setiap kode dipisahkan 1 baris kosong berwarna
kuning (tinggi 13,5).

**Urutan** - kode menaik, lalu tanggal masuk menaik, lalu urutan asal.
Kolom TANGGAL menampilkan *tanggal surat*, jadi tidak selalu menaik.

## Kolom bantu (disembunyikan)

Kolom J, K, L disembunyikan dan tidak ikut tercetak.

| Kolom | Isi | Dipakai untuk |
| --- | --- | --- |
| J | `NOMOR SURAT` | pelacakan duplikat |
| K | `TANGGAL MASUK` | kunci urutan |
| L | `URUT` | pemecah urutan terakhir |

Daftar yang belum punya kolom ini akan dicari kembali tanggal masuk dan nomornya
dari berkas `*.xls` yang ada, dicocokkan lewat uraian.

## Sheet "Daftar Isi Surat"

Sheet rekap berisi **nilai biasa, bukan rumus**. Isinya ditulis langsung oleh
script sehingga sama persis dengan gabungan semua sheet kode.

Sebelumnya dipakai rumus `=IF('UM'!$C11="","",'UM'!A11)` di setiap sel. Untuk
6 276 surat itu berarti 50 208 rumus: file jadi berat dan lambat dibuka, padahal
isinya tidak pernah berubah di luar script. Sekarang cukup nilai.

Konsekuensinya: mengedit langsung salah satu sheet kode **tidak** otomatis
mengubah sheet rekap. Kalau ada perubahan data, jalankan ulang script.

## Hasil saat ini

Folder Januari–Agustus 2026: 6 530 baris sumber, 254 duplikat dibuang,
**6 276 surat** di 41 sheet kode, 0 rumus.

| Sheet | Surat | Sheet | Surat | Sheet | Surat |
| --- | --- | --- | --- | --- | --- |
| KP | 1358 | HM | 115 | PP | 4 |
| UM | 1238 | OR | 52 | BM | 3 |
| PR | 839 | PB | 55 | BW | 3 |
| PW | 522 | RU | 28 | LB | 3 |
| SM | 466 | BK | 26 | IW | 2 |
| LAIN-LAIN | 458 | UN | 22 | KD | 2 |
| HK | 273 | HL | 18 | NO | 2 |
| PA | 239 | PL | 14 | PH | 2 |
| PS | 229 | PE | 7 | AH | 1 |
| KU | 203 | SA | 6 | AP | 1 |
| CK | 40 | ND | 36 | BU | 1 |
| | | | | CM | 1 |
| | | | | KM | 1 |
| | | | | LE | 1 |
| | | | | LP | 1 |
| | | | | NA | 1 |
| | | | | SD | 1 |
| | | | | SE | 1 |
| | | | | SL | 1 |

## Judul sheet

`JUDUL_SHEET` di skrip berisi judul yang sudah ada di file contoh. Kode yang tidak
terdaftar memakai nama kodenya sendiri. Tinggal isi dict-nya kalau sudah ada
judul resminya.

## Catatan

- File `.xls` ini sebenarnya berformat xlsx (OOXML) walaupun ekstensinya `.xls`.
  Skrip membacanya lewat `io.BytesIO` supaya openpyxl tidak menolak berdasarkan
  ekstensi. Kalau nanti dapat `.xls` sungguhan (BIFF), konversi dulu ke `.xlsx`.
- Kolom `NO AGENDA` (I) di file contoh dibiarkan kosong.
- Semua sheet dibangun ulang dari nol setiap kali dijalankan. Isi baris tetap sama
  karena nilainya disalin apa adanya dari berkas sumber.
- Penamaan berkas di folder tidak seragam besar-kecilnya; itu tidak masalah.
"# daftar_berkas_aktif" 
