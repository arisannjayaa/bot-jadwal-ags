🤖 WhatsApp Bot Jadwal Event (Agspro)
Bot WhatsApp otomatis untuk menarik data jadwal loading dan daftar alat langsung dari file Excel (.xlsx) di Google Drive. Memudahkan kru lapangan untuk mengecek tugas harian melalui perintah teks sederhana.

🚀 Fitur Utama
Menu Otomatis: Cukup ketik 1, 2, atau 3 untuk cek jadwal.

Integrasi Google Drive: Membaca file Excel asli tanpa perlu konversi manual.

Auto-Polling (Ronda): Mengecek revisi admin setiap 15 menit dan mengirim notifikasi otomatis jika ada perubahan.

Smart Calendar: Otomatis mendeteksi tab bulan (JANUARI - DESEMBER).

🛠️ Setup Google Cloud Console
Sebelum menjalankan kode, Anda wajib menyiapkan akses API:

Buka Google Cloud Console.

Buat Project Baru (misal: Bot-Jadwal-Agspro).

Enable API: Cari dan aktifkan Google Drive API.

OAuth Consent Screen: * Pilih User Type: External.

Isi App Name dan Email.

Penting: Tambahkan email Anda di bagian Test Users.

Credentials:

Klik Create Credentials -> OAuth Client ID.

Application Type: Desktop App.

Download file JSON yang dihasilkan, ubah namanya menjadi credentials.json, dan masukkan ke folder project.

💻 Instalasi Lokal (Windows/Laragon)
Persyaratan: Pastikan sudah terinstall Node.js (versi 18+).

Clone/Copy Project: Masukkan semua file ke folder C:\laragon\www\project-wa.

Instal Library:

Bash
npm install
Jalankan:

Bash
node index.js
Otorisasi: * Scan QR Code yang muncul dengan WhatsApp Anda.

Klik link Google yang muncul di terminal untuk izin akses Drive.

Wajib: Centang izin "View files in Google Drive" saat login.

Copy kode dari browser dan paste ke terminal.

🏠 Setup di Proxmox (Ubuntu Server VM)
Jika Anda ingin bot menyala 24 jam di rumah menggunakan Proxmox:

Buat VM/LXC dengan OS Ubuntu 22.04.

Update & Install Node.js:

Bash
sudo apt update
curl -fsSL https://deb.nodesource.com/setup_18.x | sudo bash -
sudo apt install -y nodejs
Install Library Chrome (Penting untuk WA):
Karena VPS/Proxmox tidak punya GUI, kita butuh library pendukung untuk menjalankan browser internal:

Bash
sudo apt install -y libgbm-dev wget unzip fontconfig locales gconf-service libasound2 libatk1.0-0 libc6 libcairo2 libcups2 libdbus-1-3 libexpat1 libfontconfig1 libgcc1 libgconf-2-4 libgdk-pixbuf2.0-0 libglib2.0-0 libgtk-3-0 libnspr4 libpango-1.0-0 libpangocairo-1.0-0 libstdc++6 libx11-6 libx11-xcb1 libxcb1 libxcomposite1 libxcursor1 libxdamage1 libxext6 libxfixes3 libxi6 libxrandr2 libxrender1 libxss1 libxtst6 ca-certificates fonts-liberation libappindicator1 libnss3 lsb-release xdg-utils
Gunakan PM2 (Agar Bot Auto-Restart):

Bash
sudo npm install -map2 -g
pm2 start index.js --name bot-agspro
pm2 save
pm2 startup
📋 Cara Penggunaan
Kirim pesan berikut ke nomor WhatsApp Bot:

1 : Jadwal hari ini.

2 : Jadwal besok.

3 : Rekap semua jadwal bulan ini.

halo / menu : Menampilkan menu utama.

⚠️ Troubleshooting
Error 403 Forbidden: Hapus file token.json, restart bot, dan login ulang dengan mencentang izin Drive.

Tanggal Error (Angka): Pastikan fungsi formatTanggalExcel sudah aktif di index.js.

QR Code Tidak Muncul: Pastikan terminal memiliki ukuran font yang cukup atau cek koneksi internet.

Project Manager: Wayan Arisanjaya

Developer: Gemini AI Collaboration

Year: 2026