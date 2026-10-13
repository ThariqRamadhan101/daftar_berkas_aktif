"""
Memasukkan data "surat masuk" ke dalam Daftar Isi Berkas Aktif.

Tanpa argumen, mereproduksi Daftar Isi dari satu sumber:

    python build_daftar_isi.py

Menambahkan bulan berikutnya, memakai daftar bulan lalu sebagai basis:

    python build_daftar_isi.py --sumber "surat masuk Oktober 2026.xls" \
                               --dasar "Daftar Isi Berkas Aktif Juni 2026.xlsx" \
                               --keluaran "Daftar Isi Berkas Aktif Oktober 2026.xlsx" \
                               --bulan Oktober

Surat dibagi per sheet berdasarkan 2 huruf pertama kode surat
(mis. UM1010 -> sheet "UM"), sisanya masuk sheet "LAIN-LAIN".

Baris yang uraiannya identik dianggap duplikat dan hanya disimpan satu,
yaitu yang muncul lebih dulu (dari daftar bulan sebelumnya).
"""
import argparse
import glob
import io
import os
import re
import sys
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ============================== KONFIGURASI ==============================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SUMBER = os.path.join(BASE_DIR, "surat masuk Juni 2026.xls")
TEMPLATE = os.path.join(BASE_DIR, "Daftar Isi Berkas Aktif.xlsx")
OUTPUT = os.path.join(BASE_DIR, "Daftar Isi Berkas Aktif Juni 2026.xlsx")
FOLDER_SUMBER = os.path.join(BASE_DIR, "Daftar Berkas Per Bulan")

NAMA_BULAN = "Juni"
TAHUN = 2026
SHEET_LAIN = "LAIN-LAIN"
NAMA_SHEET_REKAP = "Daftar Isi Surat"
SHEET_BUKAN_DATA = {"Template", NAMA_SHEET_REKAP}

# Kolom bantu, disembunyikan. Dipakai supaya bulan berikutnya tidak perlu
# menebak lagi tanggal masuk dan nomor surat dari uraian.
KOL_BANTU_NOMOR = 10      # J
KOL_BANTU_MASUK = 11      # K
KOL_BANTU_URUT = 12       # L
KOLOM_BANTU = (
    (KOL_BANTU_NOMOR, "NOMOR SURAT"),
    (KOL_BANTU_MASUK, "TANGGAL MASUK"),
    (KOL_BANTU_URUT, "URUT"),
)

BULAN_URUT = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember",
]

# Lebar kolom (dari file contoh)
LEBAR_KOLOM = {
    "A": 14.0,
    "B": 15.4,
    "C": 21.0,
    "D": 57.4,
    "E": 14.8,
    "F": 12.0,
    "G": 21.0,
    "H": 22.8,
}

# Jumlah baris yang dijangkau rumus pada sheet rekap (dalam 1 sheet)
BATAS_BARIS_REKAP = 20000

# Judul sheet yang sudah ada di file contoh. Kode yang tidak terdaftar
# akan memakai nama kodenya sendiri, edit dict ini bila perlu.
JUDUL_SHEET = {
    "HK": "HUKUM (HK)",
    "UM": "UMUM (UM)",
    "SM": "PENGEMBANGAN SUMBER DAYA MANUSIA ( SM )",
    "BK": "BINA KONSTRUKSI (BK)",
    "RU": "PERUMAHAN (RU)",
    "SA": "SUMBER DAYA AIR (SA)",
    "PB": "PENGADAAN BARANG DAN JASA",
    "PS": "PENGELOLAAN ASET BARANG MILIK NEGARA ( PS )",
    "PR": "PERENCANAAN (PR)",
    "PA": "PENGELOLA DATA ( PA )",
    "OR": "ORGANISASI DAN TATA LAKSANA (OR)",
    "KU": "KEUANGAN (KU)",
    "KP": "KEPEGAWAIAN ( KP )",
    "HM": "HUBUNGAN MASYARAKAT (HM)",
}

# ============================== KOLOM SUMBER =============================

KOL_NO = "No."
KOL_TGL_MASUK = "Tanggal Penerimaan Surat"
KOL_NOMOR_TGL = "Nomor dan Tanggal Surat"
KOL_SIFAT = "Sifat Surat"
KOL_ISI = "Isi Ringkas"
KOL_DARI = "Dari"

PEMBATAS = " - "  # pemisah antara nomor surat dan tanggal surat

# ============================== STYLE ==============================

FONT_JUDUL = Font(name="Arial", size=14, bold=True)
FONT_ISI = Font(name="Arial", size=11)
FONT_ISI_BOLD = Font(name="Arial", size=11, bold=True)
KUNING = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
GARIS = Side(style="thin", color="FF000000")
BORDIR = Border(left=GARIS, right=GARIS, top=GARIS, bottom=GARIS)
RATA_TENGAH = Alignment(horizontal="center", vertical="center", wrap_text=True)
RATA_KIRI = Alignment(horizontal="left", vertical="center", wrap_text=True)
FORMAT_TANGGAL = "dd-mm-yyyy"

BARIS_JUDUL_MERGE = 1  # baris 1-4 = judul berkas, baris 5 kosong
BARIS_TAHUN = 6
BARIS_BULAN = 7
BARIS_HEADER = 9
BARIS_JUDUL_SHEET = 10
BARIS_DATA = 11
TINGGI_PEMBATAS = 13.5


# ============================== HELPER ==============================


def bersihkan(nilai):
    """Ubah sel kosong / "-" menjadi string kosong."""
    if nilai is None:
        return ""
    teks = str(nilai).strip()
    if teks in ("-", "--"):
        return ""
    return teks


def pecah_nomor_tanggal(nilai):
    """Pisahkan "NOMOR - TANGGAL" menjadi (nomor, tanggal surat)."""
    teks = bersihkan(nilai)
    if PEMBATAS in teks:
        nomor, _, tanggal = teks.partition(PEMBATAS)
        return nomor.strip(), tanggal.strip()
    return teks, ""


def ambil_kode(nomor):
    """
    Kode = segmen sebelum "/", titik dan spasi tidak dianggap.
    Contoh: "UM.01.01/B/Sj/..." -> "UM0101"; "UM 0102/..." -> "UM0102".
    """
    segmen = nomor.split("/")[0] if "/" in nomor else nomor
    kode = re.sub(r"[.\s]", "", segmen).upper()
    return kode or "TANPA KODE"


def ambil_nama_sheet(nomor):
    """
    Nama sheet = 2 huruf pertama kode bila berupa huruf,
    selain itu masuk sheet LAIN-LAIN.
    """
    segmen = nomor.split("/")[0] if "/" in nomor else nomor
    cocok = re.match(r"^([A-Za-z]{2})", segmen.strip())
    if not cocok:
        return SHEET_LAIN
    return cocok.group(1).upper()


def buat_uraian(dari, nomor, isi):
    """Uraian yang ditulis di kolom D, sekaligus kunci duplikat."""
    return "Surat %s Nomor : %s %s" % (dari, nomor, isi)


def ke_tanggal(nilai):
    """Ubah teks atau datetime menjadi datetime, untuk pengurutan."""
    if isinstance(nilai, datetime):
        return nilai
    teks = bersihkan(nilai)
    if not teks:
        return datetime.min
    for pola in (
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%dT%H:%M:%S",
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
    ):
        try:
            return datetime.strptime(teks[:19], pola)
        except ValueError:
            continue
    return datetime.min


# ============================== BACA SUMBER ==============================


def baca_sumber(path):
    """
    File .xls ini sebenarnya berformat xlsx (OOXML) walaupun ekstensinya .xls.
    openpyxl menolak berdasarkan ekstensi, jadi isinya dibaca lewat BytesIO.
    """
    if not os.path.exists(path):
        sys.exit("File sumber tidak ditemukan: %s" % path)
    with open(path, "rb") as berkas:
        data = berkas.read()
    try:
        buku = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        sys.exit(
            "Gagal membaca %s. Kalau file ini .xls sungguhan (bukan xlsx), "
            "konversi dulu ke .xlsx." % path
        )

    lembar = buku.active
    # Header tidak selalu di baris 1. Sebagian berkas punya judul
    # "Laporan Surat Masuk" di atasnya, jadi baris header dicari lewat
    # keberadaan kolom "No." dan "Nomor dan Tanggal Surat".
    baris_header = None
    indeks = {}
    for baris in range(1, 13):
        kolom = [bersihkan(lembar.cell(row=baris, column=c).value) for c in range(1, 14)]
        if KOL_NO in kolom and KOL_NOMOR_TGL in kolom:
            baris_header = baris
            indeks = {nama: pos for pos, nama in enumerate(kolom) if nama}
            break
    if baris_header is None:
        sys.exit(
            "Baris header (No. + Nomor dan Tanggal Surat) tidak ditemukan di %s" % path
        )

    kurang = [
        nama
        for nama in (
            KOL_NO,
            KOL_TGL_MASUK,
            KOL_NOMOR_TGL,
            KOL_SIFAT,
            KOL_ISI,
            KOL_DARI,
        )
        if nama not in indeks
    ]
    if kurang:
        sys.exit("Kolom sumber tidak lengkap, tidak ditemukan: %s" % ", ".join(kurang))

    def ambil(nama, nomor_baris):
        return lembar.cell(row=nomor_baris, column=indeks[nama] + 1).value

    daftar = []
    for nomor_baris in range(baris_header + 1, lembar.max_row + 1):
        if ambil(KOL_NO, nomor_baris) is None:
            continue
        nomor, tanggal_surat = pecah_nomor_tanggal(ambil(KOL_NOMOR_TGL, nomor_baris))
        if not nomor:
            continue
        dari = bersihkan(ambil(KOL_DARI, nomor_baris))
        isi = bersihkan(ambil(KOL_ISI, nomor_baris))
        daftar.append(
            {
                "kode": ambil_kode(nomor),
                "sheet": ambil_nama_sheet(nomor),
                "tanggal": ke_tanggal(tanggal_surat),
                "tanggal_masuk": ke_tanggal(ambil(KOL_TGL_MASUK, nomor_baris)),
                "sifat": bersihkan(ambil(KOL_SIFAT, nomor_baris)),
                "isi": isi,
                "dari": dari,
                "nomor": nomor,
                "no_sumber": bersihkan(ambil(KOL_NO, nomor_baris)),
                "uraian": buat_uraian(dari, nomor, isi),
                "asal": os.path.basename(path),
            }
        )
    buku.close()
    return daftar


# ============================ BACA DAFTAR SEBELUMNYA ============================

HURUF_KOLOM = {i: openpyxl.utils.get_column_letter(i) for i in range(1, 13)}


def buat_indeks_pemulihan(daftar_sumber):
    """
    Daftar Isi lama hanya menyimpan kode, uraian, tanggal surat, dan sifat.
    Tanggal masuk (kunci urutan) dan nomor surat tidak tersimpan, jadi dicari
    kembali dari file sumber memakai uraian yang sama.

    Satu uraian bisa cocok dengan beberapa baris sumber (surat dobel), maka
    kandidat diurutkan dan dikonsumsi berurutan (FIFO) mengikuti urutan baris
    di daftar lama.
    """
    indeks = {}
    for surat in daftar_sumber:
        indeks.setdefault(surat["uraian"], []).append(surat)
    for antrean in indeks.values():
        antrean.sort(key=lambda s: (s["tanggal_masuk"], s["no_sumber"], s["asal"]))
    return indeks


def baca_dasar(path, indeks_pemulihan):
    """Ambil daftar surat dari Daftar Isi bulan sebelumnya."""
    if not os.path.exists(path):
        sys.exit("File dasar tidak ditemukan: %s" % path)
    buku = openpyxl.load_workbook(path)
    daftar = []
    gagal = []
    for nama_sheet in buku.sheetnames:
        if nama_sheet in SHEET_BUKAN_DATA:
            continue
        lembar = buku[nama_sheet]
        for baris in range(BARIS_DATA, lembar.max_row + 1):
            # "-" sah untuk surat tanpa nomor, jadi tidak lewat bersihkan()
            kode = str(lembar.cell(row=baris, column=3).value or "").strip()
            if not kode:
                continue
            uraian = bersihkan(lembar.cell(row=baris, column=4).value)
            surat = {
                "kode": kode,
                "sheet": nama_sheet,
                "tanggal": ke_tanggal(lembar.cell(row=baris, column=5).value),
                "sifat": bersihkan(lembar.cell(row=baris, column=8).value),
                "uraian": uraian,
            }
            # kolom bantu sudah ada kalau daftar ini dibuat versi terbaru
            nomor = bersihkan(lembar.cell(row=baris, column=KOL_BANTU_NOMOR).value)
            masuk = lembar.cell(row=baris, column=KOL_BANTU_MASUK).value
            if nomor and isinstance(masuk, datetime):
                surat["nomor"] = nomor
                surat["tanggal_masuk"] = masuk
                surat["no_sumber"] = ""
            else:
                antrean = indeks_pemulihan.get(uraian)
                if antrean:
                    asli = antrean.pop(0)
                    surat["nomor"] = asli["nomor"]
                    surat["tanggal_masuk"] = asli["tanggal_masuk"]
                    surat["no_sumber"] = asli["no_sumber"]
                else:
                    # tidak ketemu di sumber: pakai tanggal surat, tetap tersimpan
                    gagal.append("%s baris %d (%s)" % (nama_sheet, baris, kode))
                    surat["nomor"] = ""
                    surat["tanggal_masuk"] = surat["tanggal"]
                    surat["no_sumber"] = ""
            daftar.append(surat)
    buku.close()

    if gagal:
        print("Perhatian: %d baris tidak ketemu di file sumber" % len(gagal))
        for ketik in gagal[:10]:
            print("  - %s" % ketik)
        if len(gagal) > 10:
            print("  ... dan %d baris lain" % (len(gagal) - 10))
        print("  Urutannya memakai tanggal surat, cek ulang bila perlu.")
    return daftar


def cek_urutan_dasar(daftar):
    """
    Daftar lama harus sudah terurut kode lalu tanggal masuk. Kalau tidak,
    urutan barisnya berubah setelah dibangun ulang.
    """
    sebelum = [(s["sheet"], s["kode"]) for s in daftar]
    sesudah = [
        (s["sheet"], s["kode"])
        for s in sorted(daftar, key=lambda s: (s["sheet"], s["kode"], s["tanggal_masuk"], s["no_sumber"]))
    ]
    if sebelum == sesudah:
        return True
    for indeks, (a, b) in enumerate(zip(sebelum, sesudah)):
        if a != b:
            print("Perhatian: urutan baris %s akan berubah di baris %d" % (a[0], indeks + 1))
            print("  sekarang %s /%s -> sesudah %s /%s" % (a[0], a[1], b[0], b[1]))
            break
    return False


def hapus_duplikat(daftar):
    """
    Uraian persis sama = duplikat. Yang lebih dulu dipertahankan, supaya isi
    bulan sebelumnya tidak ikut berubah.
    """
    terkunci = set()
    unik = []
    buang = []
    for surat in daftar:
        if surat["uraian"] in terkunci:
            buang.append(surat)
        else:
            terkunci.add(surat["uraian"])
            unik.append(surat)
    return unik, buang


def susun_label_bulan(bulan_terkumpul):
    """"Juni" + "Oktober" -> "Juni s/d Oktober"."""
    urut = [b for b in BULAN_URUT if b in bulan_terkumpul]
    urut += [b for b in bulan_terkumpul if b not in BULAN_URUT]
    if not urut:
        return ""
    if len(urut) == 1:
        return urut[0]
    return "%s s/d %s" % (urut[0], urut[-1])


def baca_bulan_dasar(path):
    """Ambil label bulan dari daftar sebelumnya, mis. "Juni"."""
    if not path or not os.path.exists(path):
        return []
    buku = openpyxl.load_workbook(path)
    bulan = []
    for nama_sheet in buku.sheetnames:
        if nama_sheet in SHEET_BUKAN_DATA:
            continue
        sel = buku[nama_sheet].cell(row=BARIS_BULAN, column=1).value
        if sel and ":" in str(sel):
            bulan += [b.strip() for b in str(sel).split(":")[1].split("s/d")]
        break
    buku.close()
    bulan = [b for b in bulan if b and b != "Januari - Desember"]
    # "Januari s/d Desember" pada Template bukan bulan berjalan
    if bulan == ["Januari", "Desember"]:
        bulan = []
    return bulan


# ============================== BACA TEMPLATE ==============================


def baca_template(path):
    """Ambil blok judul + baris header dari sheet Template."""
    if not os.path.exists(path):
        sys.exit("File template tidak ditemukan: %s" % path)
    buku = openpyxl.load_workbook(path)
    if "Template" not in buku.sheetnames:
        sys.exit("Sheet 'Template' tidak ada di %s" % path)
    lembar = buku["Template"]
    judul = []
    for baris in range(1, BARIS_HEADER):
        nilai = lembar.cell(row=baris, column=1).value
        if nilai not in (None, ""):
            judul.append((baris, str(nilai)))
    header = [str(lembar.cell(row=BARIS_HEADER, column=c).value) for c in range(1, 9)]
    lebar = {}
    for indeks_kolom in range(1, 9):
        huruf = openpyxl.utils.get_column_letter(indeks_kolom)
        lebar[huruf] = lembar.column_dimensions[huruf].width
    buku.close()
    return judul, header, lebar


# ============================== TULIS SHEET ==============================


def siapkan_lembar(lembar, judul, header, lebar, bulan, lebar_pakai, sembunyikan_bantu=False):
    """Tulis blok judul berkas + baris header, sama persis seperti Template."""
    for huruf, lebar_kolom in lebar_pakai.items():
        lembar.column_dimensions[huruf].width = lebar_kolom or lebar.get(huruf)
    if sembunyikan_bantu:
        for indeks_kolom, _ in KOLOM_BANTU:
            lembar.column_dimensions[HURUF_KOLOM[indeks_kolom]].hidden = True

    for baris, teks in judul:
        if baris in (BARIS_TAHUN, BARIS_BULAN):
            continue
        lembar.merge_cells(
            start_row=baris, start_column=1, end_row=baris, end_column=8
        )
        sel = lembar.cell(row=baris, column=1, value=teks)
        sel.font = FONT_JUDUL if baris == BARIS_JUDUL_MERGE else FONT_ISI_BOLD
        sel.alignment = Alignment(horizontal="center")

    lembar.merge_cells(start_row=BARIS_TAHUN, start_column=1, end_row=BARIS_TAHUN, end_column=3)
    sel = lembar.cell(row=BARIS_TAHUN, column=1, value="TAHUN        : %d" % TAHUN)
    sel.font = FONT_ISI_BOLD
    sel.alignment = RATA_KIRI

    lembar.merge_cells(start_row=BARIS_BULAN, start_column=1, end_row=BARIS_BULAN, end_column=4)
    sel = lembar.cell(row=BARIS_BULAN, column=1, value="BULAN        : %s" % bulan)
    sel.font = FONT_ISI_BOLD
    sel.alignment = RATA_KIRI

    for indeks_kolom, teks in enumerate(header, start=1):
        sel = lembar.cell(row=BARIS_HEADER, column=indeks_kolom, value=teks)
        sel.font = FONT_ISI_BOLD
        sel.alignment = RATA_TENGAH
        sel.border = BORDIR

    lembar.page_setup.orientation = "landscape"
    lembar.page_setup.paperSize = lembar.PAPERSIZE_A4
    lembar.page_setup.fitToWidth = 1
    lembar.page_setup.fitToHeight = 0
    lembar.sheet_properties.pageSetUpPr.fitToPage = True
    lembar.print_title_rows = "%d:%d" % (BARIS_JUDUL_MERGE, BARIS_HEADER)


def tulis_baris_pembatas(lembar, baris):
    """Baris kosong berwarna kuning sebagai pemisah antar kode."""
    for indeks_kolom in range(1, 9):
        sel = lembar.cell(row=baris, column=indeks_kolom)
        sel.fill = KUNING
        sel.border = BORDIR
    lembar.row_dimensions[baris].height = TINGGI_PEMBATAS


def tulis_judul_sheet(lembar, baris, teks):
    """Baris kuning pertama berisi judul klasifikasi."""
    for indeks_kolom in range(1, 9):
        sel = lembar.cell(row=baris, column=indeks_kolom)
        sel.fill = KUNING
        sel.border = BORDIR
    sel = lembar.cell(row=baris, column=1, value=teks)
    sel.font = FONT_ISI_BOLD
    sel.alignment = RATA_KIRI
    sel.border = BORDIR
    lembar.row_dimensions[baris].height = TINGGI_PEMBATAS


def tulis_data_surat(lembar, baris, surat, urutan):
    """
    Isi A..H: A/B kosong, C kode, D uraian, E tanggal, F/G konstanta, H sifat.
    Kolom bantu J/K/L disembunyikan, isinya tidak ikut tercetak.
    """
    isi_baris = (
        None, None, surat["kode"], surat["uraian"], surat["tanggal"],
        "1 Berkas", "Asli", surat["sifat"],
    )
    for indeks_kolom, nilai in enumerate(isi_baris, start=1):
        sel = lembar.cell(row=baris, column=indeks_kolom, value=nilai)
        sel.font = FONT_ISI
        sel.alignment = RATA_KIRI if indeks_kolom == 4 else RATA_TENGAH
        sel.border = BORDIR
        if indeks_kolom == 5:
            sel.number_format = FORMAT_TANGGAL

    lembar.cell(row=baris, column=KOL_BANTU_NOMOR, value=surat["nomor"])
    sel = lembar.cell(row=baris, column=KOL_BANTU_MASUK, value=surat["tanggal_masuk"])
    sel.number_format = "dd-mm-yyyy hh:mm:ss"
    lembar.cell(row=baris, column=KOL_BANTU_URUT, value=urutan)


# ============================== REKAP ==============================


def urut_sheet(nama):
    """Abjad sesuai kode surat, tapi LAIN-LAIN paling bawah."""
    if nama == SHEET_LAIN:
        return (1, "")
    return (0, nama)


def tulis_rekap(lembar, daftar_surat, bulan, judul, header, lebar, lebar_pakai):
    """
    Sheet rekap berisi nilai biasa, bukan rumus. Semua isinya ditulis ulang
    dari data tiap kali skrip dijalankan, jadi selalu sama dengan sheet kode.

    Sebelumnya dipakai rumus =IF(sheet!A11,"",...) di tiap sel. File jadi berat
    (8 ribuan rumus) dan lambat dibuka, padahal isinya tidak pernah berubah
    di luar skrip.
    """
    siapkan_lembar(lembar, judul, header, lebar, bulan, lebar_pakai)

    baris_tujuan = BARIS_JUDUL_SHEET
    for surat in daftar_surat:
        isi_baris = (
            None, None, surat["kode"], surat["uraian"], surat["tanggal"],
            "1 Berkas", "Asli", surat["sifat"],
        )
        for indeks_kolom, nilai in enumerate(isi_baris, start=1):
            sel = lembar.cell(row=baris_tujuan, column=indeks_kolom, value=nilai)
            sel.font = FONT_ISI
            sel.alignment = RATA_KIRI if indeks_kolom == 4 else RATA_TENGAH
            sel.border = BORDIR
            if indeks_kolom == 5:
                sel.number_format = FORMAT_TANGGAL
        baris_tujuan += 1

    lembar.freeze_panes = "A%d" % BARIS_DATA


# ============================== UTAMA ==============================


def bulan_dari_nama(path):
    """Ambil (bulan, tahun) dari nama berkas, mis. 'Surat Masuk Agustus 2026.xls'."""
    nama = os.path.basename(path)
    bulan = ""
    for kandidat in BULAN_URUT:
        if re.search(r"\b%s\b" % kandidat, nama, re.IGNORECASE):
            bulan = kandidat
            break
    tahun = ""
    cocok = re.search(r"(20\d{2})", nama)
    if cocok:
        tahun = cocok.group(1)
    return bulan, tahun


def baca_folder(folder):
    """
    Baca semua berkas surat masuk di satu folder. Bulan diambil dari nama
    berkas, lalu diurutkan kronologis (Januari lebih dulu) supaya berkas bulan
    yang lebih awal dipertahankan saat deduplikasi.
    """
    if not os.path.isdir(folder):
        sys.exit("Folder tidak ditemukan: %s" % folder)
    berkas = []
    for pola in ("*.xls", "*.xlsx", "*.xlsm"):
        berkas += glob.glob(os.path.join(folder, pola))
    berkas = sorted(
        {os.path.abspath(p) for p in berkas if not os.path.basename(p).startswith("~$")}
    )
    hasil = []
    for path in berkas:
        bulan, tahun = bulan_dari_nama(path)
        if not bulan:
            print("  dilewati (nama berkas tanpa nama bulan): %s" % os.path.basename(path))
            continue
        isi = baca_sumber(path)
        for surat in isi:
            surat["bulan"] = bulan
            surat["tahun"] = tahun
        hasil.append((BULAN_URUT.index(bulan), path, isi))
    hasil.sort(key=lambda baris: baris[0])
    return hasil


def tampilkan_duplikat(buang):
    if not buang:
        print("\nTidak ada duplikat.")
        return
    print(
        "\nDuplikat dibuang %d baris (uraian persis sama, disimpan yang lebih dulu):"
        % len(buang)
    )
    ringkas = {}
    for surat in buang:
        kunci = surat.get("asal", "daftar lalu")
        ringkas[kunci] = ringkas.get(kunci, 0) + 1
    for surat in buang[:15]:
        print(
            "  - [%s] %s"
            % (
                surat.get("asal", "daftar lalu"),
                bersihkan(surat["nomor"]) or surat["uraian"][:60],
            )
        )
    if len(buang) > 15:
        print("  ... dan %d baris lain" % (len(buang) - 15))
    print(
        "  asal baris yang dibuang: %s"
        % ", ".join("%s=%d" % (k, v) for k, v in sorted(ringkas.items()))
    )


def tulis_buku(daftar, label_bulan, template, keluaran):
    """Bangun workbook Daftar Isi dari daftar surat yang sudah bersih."""
    # Urut berdasarkan kode, lalu tanggal masuk
    daftar.sort(key=lambda s: (s["sheet"], s["kode"], s["tanggal_masuk"], s["no_sumber"]))

    per_sheet = {}
    for surat in daftar:
        per_sheet.setdefault(surat["sheet"], []).append(surat)

    judul, header, lebar = baca_template(template)

    buku = openpyxl.Workbook()
    lembar_template = buku.active
    lembar_template.title = "Template"
    siapkan_lembar(lembar_template, judul, header, lebar, "Januari - Desember", LEBAR_KOLOM)
    for indeks_kolom in range(1, 9):
        sel = lembar_template.cell(row=BARIS_JUDUL_SHEET, column=indeks_kolom)
        sel.fill = KUNING
        sel.border = BORDIR

    rekap_surat = []
    urutan = 0
    for nama_sheet in sorted(per_sheet, key=urut_sheet):
        lembar = buku.create_sheet(title=nama_sheet[:31])
        siapkan_lembar(
            lembar, judul, header, lebar, label_bulan, LEBAR_KOLOM, sembunyikan_bantu=True
        )
        tulis_judul_sheet(
            lembar, BARIS_JUDUL_SHEET, JUDUL_SHEET.get(nama_sheet, nama_sheet)
        )

        baris = BARIS_DATA
        kode_sebelumnya = None
        for surat in per_sheet[nama_sheet]:
            if kode_sebelumnya is not None and surat["kode"] != kode_sebelumnya:
                tulis_baris_pembatas(lembar, baris)
                baris += 1
            urutan += 1
            tulis_data_surat(lembar, baris, surat, urutan)
            rekap_surat.append(surat)
            kode_sebelumnya = surat["kode"]
            baris += 1

        lembar.freeze_panes = "A%d" % BARIS_DATA
        print(
            "  %-12s %4d surat, %2d kode, baris %d-%d"
            % (
                nama_sheet,
                len(per_sheet[nama_sheet]),
                len({s["kode"] for s in per_sheet[nama_sheet]}),
                BARIS_DATA,
                baris - 1,
            )
        )

    rekap = buku.create_sheet(title=NAMA_SHEET_REKAP)
    tulis_rekap(rekap, rekap_surat, label_bulan, judul, header, lebar, LEBAR_KOLOM)

    buku.save(keluaran)
    print("\nTotal surat : %d" % len(daftar))
    print("Sheet kode  : %d" % len(per_sheet))
    print("Baris rekap : %d (nilai biasa, tanpa rumus)" % len(rekap_surat))
    print("Label bulan : %s" % label_bulan)
    print("Tersimpan   : %s" % keluaran)


def parse_args():
    parser = argparse.ArgumentParser(
        description="Bangun Daftar Isi Berkas Aktif dari data surat masuk."
    )
    parser.add_argument(
        "--folder",
        metavar="FOLDER",
        help="folder berisi berkas surat masuk bulanan, mis. 'Daftar Berkas Per Bulan'",
    )
    parser.add_argument(
        "--sumber",
        action="append",
        metavar="BERKAS",
        help="file surat masuk yang ditambahkan, boleh diulang",
    )
    parser.add_argument(
        "--dasar",
        metavar="BERKAS",
        help="Daftar Isi bulan sebelumnya, isinya ikut dipertahankan",
    )
    parser.add_argument("--keluaran", metavar="BERKAS", help="nama file hasil")
    parser.add_argument("--bulan", help="nama bulan untuk baris BULAN")
    parser.add_argument("--template", metavar="BERKAS", help="file Template")
    return parser.parse_args()


def kumpulkan_semua_sumber(path_terpilih):
    """
    Semua "surat masuk *.xls" di folder dipakai sebagai sarana pemulihan
    tanggal masuk/nomor surat untuk baris dari daftar sebelumnya.
    """
    nama_semua = set(path_terpilih)
    for pola in ("surat masuk *.xls", "surat masuk *.xlsx"):
        for path in glob.glob(os.path.join(BASE_DIR, pola)):
            nama_semua.add(os.path.normcase(os.path.abspath(path)))
    return nama_semua


def main():
    args = parse_args()
    template = os.path.abspath(os.path.join(BASE_DIR, args.template or TEMPLATE))

    if args.folder:
        folder = os.path.abspath(os.path.join(BASE_DIR, args.folder))
        berkas = baca_folder(folder)
        if not berkas:
            sys.exit("Tidak ada berkas surat masuk di folder: %s" % folder)

        print("Folder %s" % folder)
        gabungan = []
        bulan_terkumpul = []
        tahun_terakhir = str(TAHUN)
        for _, path, isi in berkas:
            bulan, tahun = bulan_dari_nama(path)
            print("  %-34s %4d surat" % (os.path.basename(path), len(isi)))
            gabungan += isi
            if bulan not in bulan_terkumpul:
                bulan_terkumpul.append(bulan)
            if tahun:
                tahun_terakhir = tahun

        label_bulan = susun_label_bulan(bulan_terkumpul)
        if not gabungan:
            sys.exit("Tidak ada data surat yang bisa diproses.")
        daftar, buang = hapus_duplikat(gabungan)
        tampilkan_duplikat(buang)

        if args.keluaran:
            keluaran = os.path.abspath(os.path.join(BASE_DIR, args.keluaran))
        else:
            keluaran = os.path.join(
                BASE_DIR,
                "Daftar Isi Berkas Aktif %s %s.xlsx" % (bulan_terkumpul[-1], tahun_terakhir),
            )
        tulis_buku(daftar, label_bulan, template, keluaran)
        return

    # ---- mode lama: satu/lebih sumber + daftar bulan sebelumnya ----
    sumber = [os.path.abspath(os.path.join(BASE_DIR, p)) for p in (args.sumber or [SUMBER])]
    dasar = os.path.abspath(os.path.join(BASE_DIR, args.dasar)) if args.dasar else None
    bulan_baru = args.bulan or NAMA_BULAN

    if dasar and not args.keluaran:
        sys.exit("--dasar harus dipakai bersama --keluaran supaya file lama tidak tertimpa.")
    keluaran = os.path.abspath(os.path.join(BASE_DIR, args.keluaran or OUTPUT))
    if dasar and os.path.normcase(keluaran) == os.path.normcase(dasar):
        sys.exit("--keluaran tidak boleh sama dengan --dasar.")

    baru = []
    for path in sumber:
        isi = baca_sumber(path)
        print("Sumber %-30s %4d surat" % (os.path.basename(path), len(isi)))
        baru += isi

    semua_sumber = []
    seen = set()
    for path in sorted(kumpulkan_semua_sumber(sumber)):
        kunci = os.path.normcase(path)
        if kunci in seen or not os.path.exists(path):
            continue
        seen.add(kunci)
        semua_sumber += baca_sumber(path)
    indeks_pemulihan = buat_indeks_pemulihan(semua_sumber)

    gabungan = []
    if dasar:
        lama = baca_dasar(dasar, indeks_pemulihan)
        cek_urutan_dasar(lama)
        print("Dasar %-30s %4d surat" % (os.path.basename(dasar), len(lama)))
        gabungan += lama
    gabungan += baru
    if not gabungan:
        sys.exit("Tidak ada data surat yang bisa diproses.")

    daftar, buang = hapus_duplikat(gabungan)
    tampilkan_duplikat(buang)

    bulan_terkumpul = baca_bulan_dasar(dasar)
    bulan_terkumpul.append(bulan_baru)
    label_bulan = susun_label_bulan(bulan_terkumpul)
    tulis_buku(daftar, label_bulan, template, keluaran)


if __name__ == "__main__":
    main()
