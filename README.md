# Portal Koperasi - Sistem Pengajuan & Monitoring Pinjaman

Full-Stack Node.js (Express.js) & Google Sheets Database Architecture dengan Antarmuka Vanilla JS + Tailwind CSS SPA.

## 🚀 Panduan Instalasi & Menjalankan

1. **Install Dependencies:**
   `ash
   npm install
   `

2. **Konfigurasi Environment (.env):**
   Salin .env.example menjadi .env lalu lengkapi spreadsheet ID dan konfigurasi SMTP Anda:
   `env
   PORT=3000
   SPREADSHEET_ID=1abc..._spreadsheet_id_anda
   GOOGLE_APPLICATION_CREDENTIALS=./credentials.json

   SMTP_HOST=smtp.gmail.com
   SMTP_PORT=587
   SMTP_SECURE=false
   SMTP_USER=koperasi.official.system@gmail.com
   SMTP_PASS=xxxx-xxxx-xxxx-xxxx
   EMAIL_FROM=Portal Koperasi <koperasi.official.system@gmail.com>
   `

3. **Jalankan Aplikasi:**
   `ash
   npm start
   # atau untuk development mode:
   npm run dev
   `
   Buka browser di: http://localhost:3000