<p align="center">
  <img src="/img/demo1.png" alt="SmartExcel Banner" width="900"/>
</p>

<h1 align="center">📊 SMART EXCEL — ASISTEN RUMUS EXCEL</h1>

<p align="center">
  <em>Asisten Pintar untuk Otomatisasi Rumus di Microsoft Excel 💡</em>
</p>

<p align="center">
  <a href="LICENSE"><img src="https://img.shields.io/github/license/Sneijderlino/SmartExcel-Asisten-Rumus-Excel?style=for-the-badge&color=2ecc71" alt="License"></a>
  <a href="https://www.python.org/"><img src="https://img.shields.io/badge/Python-3.8+-3776AB.svg?style=for-the-badge&logo=python" alt="Python"></a>
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20macOS-blue?style=for-the-badge" alt="Supported OS">
  <img src="https://img.shields.io/github/stars/Sneijderlino/SmartExcel-Asisten-Rumus-Excel?style=for-the-badge&color=yellow" alt="Stars">
  <img src="https://img.shields.io/github/forks/Sneijderlino/SmartExcel-Asisten-Rumus-Excel?style=for-the-badge&color=orange" alt="Forks">
</p>

---

## 🧠 Apa Itu SmartExcel?

_SmartExcel_ adalah aplikasi berbasis _Python + Tkinter_ yang berfungsi sebagai asisten cerdas untuk membantu pengguna dalam membuat, memperbaiki, dan mengotomatisasi _rumus-rumus Excel_ tanpa harus mengetik manual.  
Dirancang khusus untuk pekerja kantoran, guru, mahasiswa, dan analis data yang ingin bekerja _lebih cepat, akurat, dan efisien_.

---

## ⚡ Fitur Unggulan

| 🚀 Fitur                           | Deskripsi                                                                                                                |
| :--------------------------------- | :----------------------------------------------------------------------------------------------------------------------- |
| 💬 _Asisten Rumus Otomatis_        | Cukup tulis deskripsi tugas (misal: “Cari total nilai kolom B”) → SmartExcel akan menghasilkan rumusnya secara otomatis. |
| 🧩 _Deteksi & Koreksi Rumus Error_ | Sistem mendeteksi kesalahan logika pada formula Excel dan memberikan solusi koreksi cepat.                               |
| 📈 _Rekomendasi Fungsi Excel_      | Memberikan saran fungsi Excel terbaik berdasarkan konteks input pengguna.                                                |
| 🧮 _Mode Panduan Interaktif_       | Tersedia mode “belajar rumus” untuk membantu pemula memahami cara kerja setiap fungsi.                                   |
| 🔄 _Ekspor & Impor Template_       | Simpan template rumus atau struktur kerja favoritmu ke file .json untuk digunakan kembali.                               |
| 🎨 _UI Modern dan Responsif_       | Dibangun dengan Tkinter yang diberi sentuhan tema modern dan smooth transition.                                          |

---

## 🧰 Dependensi & Persyaratan

### 📦 Library Utama

| Library     | Fungsi                                |   Wajib    |
| :---------- | :------------------------------------ | :--------: |
| _pandas_    | Manipulasi dan ekspor data Excel      |     ✅     |
| _openpyxl_  | Engine Excel untuk Pandas             |     ✅     |
| _tkinter_   | GUI utama aplikasi                    |     ✅     |
| _pyperclip_ | Copy rumus ke clipboard secara instan |     ✅     |
| _Pillow_    | Untuk tampilan ikon dan gambar        | ⚙ Opsional |

---

## 🧑‍💻 Instalasi & Menjalankan Aplikasi

### 1️⃣ Clone Repositori

```bash
git clone https://github.com/Sneijderlino/SmartExcel-Asisten-Rumus-Excel.git
cd SmartExcel-Asisten-Rumus-Excel
```

---

### 🧩 Instalasi di Visual Studio Code (VS Code)

```bash
Langkah-langkah:

1. Buka VS Code
2. Klik File → Open Folder → pilih folder SmartExcel-Asisten-Rumus-Excel
3. Pastikan Python sudah terpasang di sistem
4. Di VS Code, buka terminal (Ctrl + `)
5. Jalankan:
python -m venv venv
venv\Scripts\activate   # Windows
# atau
source venv/bin/activate  # Linux/Mac
pip install -r requirements.txt
python SmartExcel.py
```

---

### 🐧 Instalasi di Kali Linux

Pastikan Python 3 dan pip sudah tersedia:

```bash
sudo apt update
sudo apt install python3 python3-pip python3-tk -y
```

Jalankan:

```bash
git clone https://github.com/Sneijderlino/SmartExcel-Asisten-Rumus-Excel.git
cd SmartExcel-Asisten-Rumus-Excel
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python3 SmartExcel.py
```

---

### 📦 Download Versi Rilis (Tag Release)

Kamu tidak ingin repot setup manual?
Cukup download versi rilis siap pakai di GitHub!
🔽 Langkah:

```bash
1. Kunjungi halaman 📦 Releases SmartExcel
2. Pilih versi terbaru (misal v1.0.0)
3. Unduh file berikut di bagian Assets:
             > SmartExcel-v1.0.0.zip → Source code
             > SmartExcel-v1.0.0-setup.exe → Installer Windows
4. Ekstrak atau jalankan langsung file .exe
```

---

🖼 Preview Aplikasi

<p align="center">
  <img src="/img/demo1.png" alt="SmartExcel Preview" width="750"/>
  <br>
  <em>Rumus Otomatis Dibuat</em>
</p>

---

### ✅ Tips:

- Gunakan Windows 10/11 atau Linux dengan Python 3.8+

- Jika diperlukan, install dependensi dari requirements.txt:
- pip install -r requirements.txt

<p align="center">
  <img src="https://img.shields.io/badge/Made%20with-Python-blue?style=for-the-badge&logo=python" alt="Python Badge"/>
  <img src="https://img.shields.io/badge/Status-Active-success?style=for-the-badge" alt="Status Active"/>
  <img src="https://img.shields.io/github/stars/Sneijderlino/youtube-downloader-pro?style=for-the-badge" alt="GitHub Stars"/>
  <img src="https://img.shields.io/github/forks/Sneijderlino/youtube-downloader-pro?style=for-the-badge" alt="GitHub Forks"/>
</p>

---

<h3 align="center">📜 Lisensi</h3>

<p align="center">
  Proyek ini dilisensikan di bawah <a href="LICENSE">MIT License</a>.<br>
  Bebas digunakan, dimodifikasi, dan dibagikan selama mencantumkan kredit.
</p>

---

<h3 align="center">💬 Dukungan & Kontribusi</h3>

<p align="center">
  💡 Temukan bug atau ingin menambahkan fitur baru?<br>
  Silakan buka <a href="https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA/issues">Issues</a> atau buat <a href="https://github.com/Sneijderlino/Aplikasi-Laporan-eKINERJA/pulls">Pull Request</a>.<br><br>
  ⭐ Jangan lupa beri bintang jika proyek ini bermanfaat!
</p>

---

<p align="center">
  Dibuat dengan ❤ oleh <a href="https://www.tiktok.com/@sneijderlino_official?is_from_webapp=1&sender_device=pc">Sneijderlino</a><br>
  <em>“Code. Create. Conquer.”</em>
</p>

---
# SmartExcel-Asisten-Rumus-Excel
