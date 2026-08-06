# Dashboard AMORE Nasional

Dashboard live untuk monitoring AMORE Nasional — data ditarik langsung dari Google Spreadsheet **`[LIVE] Dashboard AMORE Nasional - For MD`** setiap 5 menit.

## URL Live

Setelah GitHub Pages aktif, dashboard tersedia di:

**https://pqmahm-sketch.github.io/dashboard_amore/**

## Cara Kerja

- `index.html` — dashboard UI (semua chart, tabel, filter).
- `data-loader.js` — mengambil data dari Google Sheets via endpoint `gviz/tq?tqx=out:csv` (tanpa API key, tanpa server backend).
- Refresh otomatis setiap **5 menit**, atau klik tombol **↻ Refresh** untuk update manual.

## Setup Sekali (Wajib)

Supaya dashboard dapat membaca sheet, spreadsheet harus di-share sebagai *view-only untuk siapapun yang punya link*.

1. Buka spreadsheet **`[LIVE] Dashboard AMORE Nasional - For MD`** di Google Sheets.
2. Klik **Share** (kanan atas).
3. Di bagian **General access**, ubah dari *Restricted* → **Anyone with the link**.
4. Pastikan role adalah **Viewer** (bukan Editor/Commenter).
5. Klik **Done**.

Sheet tetap read-only untuk publik — hanya user yang sudah dishare eksplisit yang bisa edit.

## Struktur Sheet yang Diperlukan

Data-loader mengharapkan tab-tab berikut ada di workbook (nama harus persis sama):

| Tab | Isi |
|---|---|
| `Dashboard` | Ringkasan bulan berjalan (Overall + 5 sub-proses) per MD, Region, Nasional |
| `Data Mentah` | Raw laporan (Timestamp, MD, No. Rangka, 5 proses, Status, No Claim, No AHASS, Region) |
| `Rekap Result All MD` | Trend historis per bulan (Jan..Dec) per MD, Region, Nasional untuk 6 proses |

Kalau salah satu tab di-rename, dashboard akan menampilkan pesan error yang menjelaskan tab mana yang tidak ditemukan.

## Deployment via GitHub Pages

Repo ini sudah termasuk workflow di `.github/workflows/deploy.yml`. Aktifkan sekali:

1. Buka **Settings → Pages** di repo GitHub ini.
2. Di **Source**, pilih **GitHub Actions**.
3. Push ke branch `main` — deploy otomatis berjalan; hasil workflow bisa dilihat di tab **Actions**.

Setiap update ke `main` akan otomatis redeploy dashboard.

## Struktur File

```
.
├── index.html            # Dashboard UI
├── data-loader.js        # Live-fetch dari Google Sheets
├── README.md             # Dokumen ini
└── .github/workflows/
    └── deploy.yml        # GitHub Pages auto-deploy
```

## Troubleshooting

- **"Gagal fetch sheet" / dashboard tidak muncul**: pastikan spreadsheet sudah di-share `Anyone with the link can view` (lihat Setup di atas).
- **Data lama muncul**: klik **↻ Refresh**. Kalau masih lama, cek "Last Update" di sheet — apa memang formula-nya sudah recalculate.
- **MD tidak muncul di dashboard**: pastikan MD tersebut ada di sheet `Dashboard`. Jika baru ditambahkan dan belum ada di `Data Mentah`, region-nya akan diambil dari fallback map di `data-loader.js` (hanya berisi 29 MD default).
- **Nama tab berubah**: edit konstanta `SHEETS` di bagian atas `data-loader.js`.

## Ubah Interval Auto-Refresh

Edit baris di `index.html`:

```js
setInterval(() => loadAndRender(true), 5 * 60 * 1000); // 5 menit
```

Ubah angka `5` menjadi menit yang diinginkan (mis. `10 * 60 * 1000` = 10 menit).
