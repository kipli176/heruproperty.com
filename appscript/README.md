# Backend Google Apps Script — Heru Property

Backend ini melayani dua request dari frontend:

- `POST /exec`: menerima form titip jual dan foto.
- `GET /exec?list=1`: mengirim listing berstatus `Tayang` tanpa data privat.

## Cara memasang

1. Buat Google Spreadsheet baru.
2. Dari Spreadsheet, buka **Extensions → Apps Script**.
3. Ganti isi `Code.gs` dengan isi file `appscript/Code.gs` dari repository ini.
4. Di **Project Settings**, ubah timezone menjadi `Asia/Jakarta`.
5. Jalankan fungsi `setup()` dari editor dan izinkan akses Spreadsheet serta Drive.
6. Kembali ke Spreadsheet. Tab **Titip Jual** dan folder Drive untuk foto akan dibuat otomatis.
7. Pilih **Deploy → New deployment → Web app**.
8. Atur:
   - **Execute as:** Me
   - **Who has access:** Anyone
9. Klik **Deploy**, lalu salin URL yang berakhiran `/exec` ke variabel `ENDPOINT` di `config.js`.

Jika script dibuat terpisah dan bukan dari menu Extensions pada Spreadsheet, jalankan:

```javascript
configure('ID_SPREADSHEET', 'ID_FOLDER_DRIVE_OPSIONAL');
```

ID folder boleh dikosongkan. Script akan membuat folder foto sendiri.

## Alur moderasi

Submission baru masuk dengan status `Baru`. Agen memeriksa data di tab **Titip Jual**, melengkapi atau memperbaiki data bila perlu, lalu mengubah kolom **Status** menjadi `Tayang`.

Hanya baris dengan status persis `Tayang` dan koordinat yang valid yang dikirim ke website. Data berikut tidak pernah dikirim oleh endpoint publik:

- Nomor HP penjual
- Harga net
- Keterangan internal
- Link postingan Facebook
- Informasi sertifikat atas nama dan PBB

Foto disimpan privat ketika form diterima. Saat listing pertama kali diminta melalui endpoint publik dalam status `Tayang`, akses foto diubah menjadi **Anyone with the link — Viewer** agar dapat tampil di website.

Untuk menarik listing dari website, ubah status menjadi `Terjual`, `Diproses`, `Ditolak`, atau nilai selain `Tayang`.

## Memperbarui deployment

Setelah mengubah kode Apps Script, buat versi deployment baru melalui **Deploy → Manage deployments → Edit → New version**. URL `/exec` tetap sama, sehingga `config.js` tidak perlu diubah.

## Catatan Google Workspace

Beberapa akun organisasi melarang berbagi file Drive kepada publik. Jika foto tidak tampil, administrator Workspace perlu mengizinkan **Anyone with the link**, atau backend foto perlu dipindahkan ke layanan penyimpanan publik lain.
