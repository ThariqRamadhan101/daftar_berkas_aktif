# AGENTS.md

Single-script repo: `build_daftar_isi.py` builds `Daftar Isi Berkas Aktif` from monthly `surat masuk` files. Only dependency is `openpyxl`.

## Commands

- Preferred (full rebuild, chronological): `python build_daftar_isi.py --folder "Daftar Berkas Per Bulan"`
- Output name defaults to `Daftar Isi Berkas Aktif <LastMonth> <Year>.xlsx`; override with `--keluaran`.
- Legacy single-file append mode only: `python build_daftar_isi.py --sumber "<new>.xls" --dasar "<old Daftar Isi>.xlsx" --keluaran "<new Daftar Isi>.xlsx" --bulan <Bulan>`
- No tests, lint, or build. Verify by running the script and checking stdout counts (source rows, duplicates removed, per-sheet totals).

## Gotchas

- Input `*.xls` files are actually OOXML/xlsx despite the extension; script reads them via `io.BytesIO` so openpyxl ignores the extension (`build_daftar_isi.py:201`). Do not rename/convert unless a file is true BIFF (script exits with conversion message).
- Close source files in Excel before running (Windows lock / `~$` temp files are skipped).
- Month/year come from the filename (`bulan_dari_nama`, `build_daftar_isi.py:587`); files without a month name are skipped. Capitalization is free.
- `*.xls` / `*.xlsx` are gitignored, except tracked `template.xlsx`. Edit the `.py` and `README.md`, not the workbooks.
- Dedup key is the full uraian string (`Surat {Dari} Nomor : {Nomor} {Isi Ringkas}`), not the letter number — same number with different content is kept.
- Every run rebuilds all sheets from zero; the `Daftar Isi Surat` recap holds plain values, not formulas, so manual edits to output are lost on rerun.
- Sort order is sheet, kode, tanggal masuk, no asal (`tulis_buku`, `build_daftar_isi.py:662`); col E shows tanggal surat so it is not always ascending.
- Hidden helper cols J/K/L (nomor, tanggal masuk, urut) drive ordering and dedup recovery; keep them hidden, do not delete.
- Sheet name = first 2 letters of segment before `/` (`ambil_nama_sheet`), else `LAIN-LAIN`. Literal rule creates odd sheets (`PP` from `PPE.`, `UN` from `UND-`); move to `LAIN-LAIN` only by adding exceptions in code.
