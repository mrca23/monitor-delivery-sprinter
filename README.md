# Report Sprinter Delivery

Tool untuk membuat report per sprinter dari export JMS **Monitor Delivery (Refine)(Detail)** (.xlsx).

**Pakai online:** https://mrca23.github.io/monitor-delivery-sprinter/

File diolah langsung di browser (pakai [ExcelJS](https://github.com/exceljs/exceljs)). Tidak ada data yang dikirim ke server.

## Isi report
Sheet **Report Sprinter** (per DP & sprinter):

| Kolom | Definisi |
|---|---|
| Total Delivery | Jumlah AWB yang di-scan delivery oleh sprinter |
| Total TTD | AWB TTD oleh penerima (Waktu TTD ada & TTD Retur = 0). TTD Retur = 1 tidak dihitung (TTD retur di pengirim) |
| Total Scan Bermasalah | AWB yang punya Waktu Paket Bermasalah |
| Total Scan Inventori | AWB yang punya Waktu Scan Inventory |
| Total AWB Tidak Scan Inventori | AWB tidak TTD (retur / belum TTD) tanpa scan inventori |
| Keterangan | Kosong, untuk diisi manual |

Sheet lain: **AWB Tidak Scan Inventori**, **AWB Gagal Tanpa Bermasalah**, **Definisi**.

## Kolom wajib di file input
`No. Waybill`, `DP Delivery`, `Waktu Delivery`, `Kode Sprinter`, `Sprinter Delivery`, `Waktu Scan Inventory`, `Waktu Paket Bermasalah`, `Waktu TTD`, `TTD Retur`

## Versi Python (offline)
```bash
pip install pandas openpyxl
python python/report_sprinter.py "Monitor Delivery(Refine)(Detail).xlsx"
```
Hasil disimpan di folder `output/` (dibuat otomatis di sebelah folder `python/`).

## Struktur
- `index.html` - halaman web
- `report.js` - logika perhitungan & pembuatan Excel (browser dan Node)
- `python/report_sprinter.py` - versi Python dengan hasil yang sama

**Jangan commit file data (.xlsx)** - berisi data pelanggan. Sudah diblokir di `.gitignore`.
