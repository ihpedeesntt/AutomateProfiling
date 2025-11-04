# Automate Matchapro Entry
Program pendukung untuk aplikasi Profiling BPS NTT. Menggunakan hasil eksport Profiling sebagai bahan untuk automate entry Matchapro. Dikembangkan menggunakan library PySide6

## Instalasi
1. Salin repositori menggunakan git clone:

```https://github.com/ihpedeesntt/AutomateProfiling.git ```

2. Pastikan di PC lokal sudah terinstall uv. Tata cara instalasi UV dapat diakses di tautan berikut:

```https://docs.astral.sh/uv/getting-started/installation/#installation-methods```

3. Rename file .env.example menjadi .env dan isikan username dan password anda di file .env tersebut

4. Siapkan file excel hasil ekspor dari aplikasi profiling. Pastikan field Idsbr, Keberadaan usaha, Catatan, Sumber Profiling sudah terisi. Note : Perhatikan kolom Idsbr duplikat jika Keberadaan usaha berkode 9, Jika tidak terisi maka bisa isi manual di file excel atau row tersebut dihapus dan entry secara manual di Matchapro. Sementara aplikasi ini belum mengakomodasi untuk direktori yang berkode Aktif Pindah (kode 7) atau Salah Kode Wilayah (kode 11). Harap dicek file excel kolom  Is new dan pastikan berkode 0.

5. Pastikan PC sudah terkoneksi VPN

6. Buka terminal pada direktori repositori yang sudah diclone dan run command untuk instalasi browser agent. Command ini hanya dilakukan sekali:

``` uv run playwright install ```

7. Run command ini setiap mau menjalankan program
``` uv run main.py ```




## Lisensi
This project depends on third-party components:

PySide6 (Qt for Python)
Copyright © The Qt Company Ltd.
Licensed under LGPL v3 (or commercial).
Source & license: https://doc.qt.io/qtforpython/licenses.html

Qt libraries (bundled with PySide6)
Various modules under LGPL v3 / GPL / other Qt-provided licenses.
See Qt’s license summary and “About Qt” menu for module-specific licensing.
