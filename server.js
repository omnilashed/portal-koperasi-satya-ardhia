/**
 * PORTAL KOPERASI BACKEND SERVER
 * Tech Stack: Node.js, Express.js, Google Sheets API (googleapis), Nodemailer, CORS, Dotenv
 */

const express = require('express');
const cors = require('cors');
const dotenv = require('dotenv');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;
const DEFAULT_SPREADSHEET_ID = '1mFlMWzoKkEeBxidakZSZz_yMZgK3qzk8nbPMSGIujJ0';
const SPREADSHEET_ID = (process.env.SPREADSHEET_ID && process.env.SPREADSHEET_ID.trim()) ? process.env.SPREADSHEET_ID.trim() : DEFAULT_SPREADSHEET_ID;
const CREDENTIALS_PATH = process.env.GOOGLE_APPLICATION_CREDENTIALS || path.join(__dirname, 'credentials.json');

// Middleware
app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Serve static frontend files from /public or root
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// Google Sheets Auth Setup
let sheetsInstance = null;

async function getSheetsClient() {
  if (sheetsInstance) return sheetsInstance;

  let resolvedPath = CREDENTIALS_PATH;
  if (!fs.existsSync(resolvedPath)) {
    resolvedPath = path.join(__dirname, 'credentials.json');
  }

  // Attempt 1: Try GOOGLE_CREDENTIALS_JSON if provided in Vercel Environment Variables
  if (process.env.GOOGLE_CREDENTIALS_JSON) {
    try {
      const rawEnv = process.env.GOOGLE_CREDENTIALS_JSON.trim();
      const credsObj = typeof rawEnv === 'string' ? JSON.parse(rawEnv) : rawEnv;

      if (credsObj && credsObj.private_key) {
        credsObj.private_key = credsObj.private_key.replace(/\\n/g, '\n');
      }

      const auth = new google.auth.GoogleAuth({
        credentials: credsObj,
        scopes: ['https://www.googleapis.com/auth/spreadsheets']
      });

      const authClient = await auth.getClient();
      sheetsInstance = google.sheets({ version: 'v4', auth: authClient });
      return sheetsInstance;
    } catch (envError) {
      console.warn('[Google Auth]: GOOGLE_CREDENTIALS_JSON gagal, mencoba file credentials.json...', envError.message);
    }
  }

  // Attempt 2: Use keyFile credentials.json directly
  try {
    if (!fs.existsSync(resolvedPath)) {
      throw new Error('File credentials.json tidak ditemukan di: ' + resolvedPath);
    }

    const auth = new google.auth.GoogleAuth({
      keyFile: resolvedPath,
      scopes: ['https://www.googleapis.com/auth/spreadsheets']
    });

    const authClient = await auth.getClient();
    sheetsInstance = google.sheets({ version: 'v4', auth: authClient });
    return sheetsInstance;
  } catch (fileError) {
    console.error('[Google Sheets Auth Error]:', fileError.message);
    throw fileError;
  }
}

// Nodemailer Transporter Setup
function createEmailTransporter() {
  return nodemailer.createTransport({
    host: process.env.SMTP_HOST || 'smtp.gmail.com',
    port: parseInt(process.env.SMTP_PORT || '587', 10),
    secure: process.env.SMTP_SECURE === 'true',
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });
}

// Helper: Format tanggal dd/MM/yyyy HH:mm
function getCurrentTimestamp() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, '0');
  const d = pad(now.getDate());
  const m = pad(now.getMonth() + 1);
  const y = now.getFullYear();
  const hr = pad(now.getHours());
  const min = pad(now.getMinutes());
  return d + '/' + m + '/' + y + ' ' + hr + ':' + min;
}

// ==========================================
// 1. ENDPOINTS: AUTENTIKASI & AKUN
// ==========================================

/**
 * POST /api/login
 * Body: { username, password }
 */
app.post('/api/login', async (req, res) => {
  try {
    const { username, password } = req.body;
    if (!username || !password) {
      return res.status(400).json({ success: false, message: 'Username dan Password wajib diisi.' });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A2:H'
    });

    const rows = response.data.values || [];
    const normalizedUsername = String(username).trim().toLowerCase();

    // Cari user berdasarkan username
    const userRow = rows.find(r => (r[0] || '').toLowerCase() === normalizedUsername);

    if (!userRow) {
      return res.status(401).json({ success: false, message: 'Username atau Password salah.' });
    }

    const [uUsername, uPassword, uRole, uNamaLengkap, uNik, uEmail, , uAccountStatus] = userRow;

    if (String(uPassword) !== String(password)) {
      return res.status(401).json({ success: false, message: 'Username atau Password salah.' });
    }

    const status = (uAccountStatus || 'Pending').trim();

    if (status === 'Pending') {
      return res.status(403).json({
        success: false,
        message: 'Akun Anda sedang menunggu persetujuan/validasi dari Admin.'
      });
    }
    if (status === 'Deactive') {
      return res.status(403).json({
        success: false,
        message: 'Akun Anda telah dinonaktifkan. Silakan hubungi Admin.'
      });
    }
    if (status === 'Rejected') {
      return res.status(403).json({
        success: false,
        message: 'Pendaftaran akun Anda ditolak oleh Admin.'
      });
    }
    if (status !== 'Active') {
      return res.status(403).json({
        success: false,
        message: 'Status akun tidak valid (' + status + '). Hubungi Admin.'
      });
    }

    return res.json({
      success: true,
      message: 'Login berhasil.',
      user: {
        username: uUsername,
        role: uRole || 'User',
        namaLengkap: uNamaLengkap || '',
        nik: uNik || '',
        email: uEmail || '',
        accountStatus: status
      }
    });
  } catch (error) {
    console.error('Error /api/login:', error);
    return res.status(500).json({ success: false, message: 'Terjadi kesalahan server: ' + error.message });
  }
});

/**
 * POST /api/register
 * Body: { username, password, role, namaLengkap, nik, email }
 */
app.post('/api/register', async (req, res) => {
  try {
    const { username, password, namaLengkap, nik, email, role } = req.body;

    if (!username || !password || !namaLengkap || !nik || !email) {
      return res.status(400).json({ success: false, message: 'Semua field formulir wajib diisi.' });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A2:H'
    });

    const rows = response.data.values || [];
    const normalizedUsername = String(username).trim().toLowerCase();
    const normalizedEmail = String(email).trim().toLowerCase();

    // Validasi duplikasi username atau email
    const duplicate = rows.find(r => 
      (r[0] || '').toLowerCase() === normalizedUsername || 
      (r[5] || '').toLowerCase() === normalizedEmail
    );

    if (duplicate) {
      const isUserDup = (duplicate[0] || '').toLowerCase() === normalizedUsername;
      return res.status(400).json({
        success: false,
        message: isUserDup 
          ? 'Username sudah terdaftar. Gunakan username lain.' 
          : 'Email sudah terdaftar. Gunakan email lain.'
      });
    }

    // Role default adalah 'User' jika tidak ditentukan atau didaftarkan lewat form umum
    const finalRole = (role && ['Admin', 'Financial', 'Pengurus', 'User'].includes(role)) ? role : 'User';
    const newRow = [
      normalizedUsername,
      password,
      finalRole,
      namaLengkap,
      nik,
      normalizedEmail,
      '',         // ResetToken kosong
      'Pending'   // Default status
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A:H',
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [newRow]
      }
    });

    return res.json({
      success: true,
      message: 'Registrasi berhasil! Akun Anda berstatus Pending dan menunggu verifikasi Admin.'
    });
  } catch (error) {
    console.error('Error /api/register:', error);
    return res.status(500).json({ success: false, message: 'Gagal melakukan registrasi: ' + error.message });
  }
});

/**
 * POST /api/auth/forgot-password
 * Body: { email }
 */
app.post('/api/auth/forgot-password', async (req, res) => {
  try {
    const { email } = req.body;
    if (!email) {
      return res.status(400).json({ success: false, message: 'Alamat email wajib diisi.' });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A2:H'
    });

    const rows = response.data.values || [];
    const normalizedEmail = String(email).trim().toLowerCase();

    let targetRowIndex = -1;
    let targetUser = null;

    for (let i = 0; i < rows.length; i++) {
      if ((rows[i][5] || '').toLowerCase() === normalizedEmail) {
        targetRowIndex = i + 2; // Mengingat A2 adalah row 2
        targetUser = rows[i];
        break;
      }
    }

    if (targetRowIndex === -1 || !targetUser) {
      return res.status(404).json({
        success: false,
        message: 'Email tidak ditemukan dalam database anggota.'
      });
    }

    // Generate 6 digit numeric OTP
    const otp = Math.floor(100000 + Math.random() * 900000).toString();

    // Update kolom ResetToken (Kolom G)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!G' + targetRowIndex,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[otp]]
      }
    });

    // Kirim email via Nodemailer jika SMTP terkonfigurasi
    try {
      const transporter = createEmailTransporter();
      await transporter.sendMail({
        from: process.env.EMAIL_FROM || '"Portal Koperasi" <noreply@koperasi.local>',
        to: normalizedEmail,
        subject: 'Kode Verifikasi Reset Password - Portal Koperasi',
        html: '<div style="font-family: Arial, sans-serif; padding: 20px; color: #333; max-width: 600px; margin: auto; border: 1px solid #e2e8f0; border-radius: 8px;"><h2 style="color: #0284c7; text-align: center;">Portal Koperasi</h2><p>Halo <strong>' + (targetUser[3] || 'Anggota') + '</strong>,</p><p>Gunakan kode OTP 6-digit berikut untuk mengatur ulang kata sandi Anda:</p><div style="text-align: center; margin: 24px 0;"><span style="font-size: 32px; font-weight: bold; letter-spacing: 6px; background-color: #f0f9ff; color: #0369a1; padding: 10px 24px; border-radius: 6px; border: 1px dashed #0284c7;">' + otp + '</span></div><p style="font-size: 13px; color: #64748b;">Jika tidak melakukan permintaan ini, abaikan email ini.</p></div>'
      });
      console.log('[OTP Sent] Email sent to ' + normalizedEmail + ' with OTP: ' + otp);
    } catch (mailError) {
      console.warn('[Mail Warning] Gagal kirim email via SMTP, namun OTP tersimpan:', mailError.message);
      return res.json({
        success: true,
        message: 'Kode OTP reset password telah digenerate. (Silakan periksa email Anda).',
        devOtp: process.env.NODE_ENV === 'development' ? otp : undefined
      });
    }

    return res.json({
      success: true,
      message: 'Kode OTP telah dikirimkan ke email Anda. Silakan cek inbox/spam.'
    });
  } catch (error) {
    console.error('Error /api/auth/forgot-password:', error);
    return res.status(500).json({ success: false, message: 'Gagal memproses permintaan reset: ' + error.message });
  }
});

/**
 * POST /api/auth/reset-password
 * Body: { email, otp, newPassword }
 */
app.post('/api/auth/reset-password', async (req, res) => {
  try {
    const { email, otp, newPassword } = req.body;
    if (!email || !otp || !newPassword) {
      return res.status(400).json({ success: false, message: 'Email, OTP, dan Password baru wajib diisi.' });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A2:H'
    });

    const rows = response.data.values || [];
    const normalizedEmail = String(email).trim().toLowerCase();
    const cleanOtp = String(otp).trim();

    let targetRowIndex = -1;
    let targetUser = null;

    for (let i = 0; i < rows.length; i++) {
      if ((rows[i][5] || '').toLowerCase() === normalizedEmail) {
        targetRowIndex = i + 2;
        targetUser = rows[i];
        break;
      }
    }

    if (targetRowIndex === -1 || !targetUser) {
      return res.status(404).json({ success: false, message: 'Email tidak terdaftar.' });
    }

    const savedOtp = String(targetUser[6] || '').trim();
    if (!savedOtp || savedOtp !== cleanOtp) {
      return res.status(400).json({ success: false, message: 'Kode OTP salah atau telah kadaluarsa.' });
    }

    // Update Password (Kolom B) dan Kosongkan ResetToken (Kolom G)
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: [
          {
            range: 'Users!B' + targetRowIndex,
            values: [[newPassword]]
          },
          {
            range: 'Users!G' + targetRowIndex,
            values: [['']]
          }
        ]
      }
    });

    return res.json({
      success: true,
      message: 'Password berhasil diubah. Silakan login kembali dengan password baru Anda.'
    });
  } catch (error) {
    console.error('Error /api/auth/reset-password:', error);
    return res.status(500).json({ success: false, message: 'Gagal mengatur ulang password: ' + error.message });
  }
});

// ==========================================
// 2. ENDPOINTS: MANAJEMEN USER (ADMIN ONLY)
// ==========================================

/**
 * GET /api/users
 * Query: ?role=Admin (atau header autentikasi)
 */
app.get('/api/users', async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A2:H'
    });

    const rows = response.data.values || [];
    const users = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      username: r[0] || '',
      role: r[2] || 'User',
      namaLengkap: r[3] || '',
      nik: r[4] || '',
      email: r[5] || '',
      accountStatus: r[7] || 'Pending'
    }));

    return res.json({ success: true, data: users });
  } catch (error) {
    console.error('Error /api/users:', error);
    return res.status(500).json({ success: false, message: 'Gagal memuat data pengguna: ' + error.message });
  }
});

/**
 * PUT /api/users/status
 * Body: { username, status }  --> status: 'Active' | 'Rejected' | 'Deactive' | 'Pending'
 */
app.put('/api/users/status', async (req, res) => {
  try {
    const { username, status } = req.body;
    if (!username || !status) {
      return res.status(400).json({ success: false, message: 'Username dan status baru wajib disediakan.' });
    }

    const validStatuses = ['Active', 'Pending', 'Deactive', 'Rejected'];
    if (!validStatuses.includes(status)) {
      return res.status(400).json({ success: false, message: 'Status tidak valid. Pilihan: ' + validStatuses.join(', ') });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!A2:H'
    });

    const rows = response.data.values || [];
    const normalizedUsername = String(username).trim().toLowerCase();

    let targetRowIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      if ((rows[i][0] || '').toLowerCase() === normalizedUsername) {
        targetRowIndex = i + 2;
        break;
      }
    }

    if (targetRowIndex === -1) {
      return res.status(404).json({ success: false, message: 'User tidak ditemukan.' });
    }

    // Update Kolom H (AccountStatus)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!H' + targetRowIndex,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[status]]
      }
    });

    return res.json({
      success: true,
      message: 'Status user ' + username + ' berhasil diubah menjadi ' + status + '.'
    });
  } catch (error) {
    console.error('Error /api/users/status:', error);
    return res.status(500).json({ success: false, message: 'Gagal memperbarui status user: ' + error.message });
  }
});

// ==========================================
// 3. ENDPOINTS: MODUL PENGAJUAN PINJAMAN/BIAYA
// ==========================================

/**
 * GET /api/pengajuan
 * Query params: role, nik, nama
 */
app.get('/api/pengajuan', async (req, res) => {
  try {
    const { role, nik, nama } = req.query;

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A2:J"
    });

    const rows = response.data.values || [];
    let list = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      timestamp: r[0] || '',
      tipeFormulir: r[1] || '',
      namaKaryawan: r[2] || '',
      nikUnitKerja: r[3] || '',
      noRekening: r[4] || '',
      atasNamaRekening: r[5] || '',
      totalDitransfer: r[6] || '0',
      cicilan: r[7] || '-',
      keterangan: r[8] || '',
      status: r[9] || 'Created'
    }));

    // Jika role adalah User, filter data milik sendiri
    if (role === 'User') {
      const userNik = (nik || '').trim().toLowerCase();
      const userNama = (nama || '').trim().toLowerCase();

      list = list.filter(item => {
        const itemNik = (item.nikUnitKerja || '').toLowerCase();
        const itemNama = (item.namaKaryawan || '').toLowerCase();
        return (userNik && itemNik.includes(userNik)) || (userNama && itemNama.includes(userNama));
      });
    }

    // Helper: Konversi berbagai format timestamp string menjadi Unix Epoch Timestamp
    function parseDateToTimestamp(dateStr) {
      if (!dateStr) return 0;
      const str = String(dateStr).trim();
      
      // Format 1: dd/MM/yyyy HH:mm:ss atau dd/MM/yyyy HH:mm
      const dmyMatch = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
      if (dmyMatch) {
        const [, day, month, year, hour = '0', minute = '0', second = '0'] = dmyMatch;
        return new Date(parseInt(year, 10), parseInt(month, 10) - 1, parseInt(day, 10), parseInt(hour, 10), parseInt(minute, 10), parseInt(second, 10)).getTime();
      }

      // Format 2: yyyy-MM-dd HH:mm:ss atau ISO string
      const ymdMatch = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2}):(\d{1,2})(?::(\d{1,2}))?)?/);
      if (ymdMatch) {
        const [, year, month, day, hour = '0', minute = '0', second = '0'] = ymdMatch;
        return new Date(parseInt(year, 10), parseInt(month, 10) - 1, parseInt(day, 10), parseInt(hour, 10), parseInt(minute, 10), parseInt(second, 10)).getTime();
      }

      const parsed = Date.parse(str);
      return isNaN(parsed) ? 0 : parsed;
    }

    // Urutkan berdasarkan waktu aktual yang paling mendekati waktu sekarang (terbaru / newest timestamp di atas)
    list.sort((a, b) => {
      const timeA = parseDateToTimestamp(a.timestamp);
      const timeB = parseDateToTimestamp(b.timestamp);
      if (timeA !== timeB) {
        return timeB - timeA; // Descending: waktu paling baru di atas
      }
      return b.rowIndex - a.rowIndex; // Fallback: baris terbawah spreadsheet di atas
    });

    return res.json({
      success: true,
      data: list
    });
  } catch (error) {
    console.error('Error /api/pengajuan:', error);
    return res.status(500).json({ success: false, message: 'Gagal mengambil data pengajuan: ' + error.message });
  }
});

/**
 * POST /api/pengajuan
 * Body: { tipeFormulir, namaKaryawan, nikUnitKerja, noRekening, atasNamaRekening, totalDitransfer, cicilan, keterangan }
 */
app.post('/api/pengajuan', async (req, res) => {
  try {
    const {
      tipeFormulir,
      namaKaryawan,
      nikUnitKerja,
      noRekening,
      atasNamaRekening,
      totalDitransfer,
      cicilan,
      keterangan
    } = req.body;

    if (!tipeFormulir || !namaKaryawan || !nikUnitKerja || !noRekening || !atasNamaRekening || !totalDitransfer) {
      return res.status(400).json({ success: false, message: 'Lengkapi semua field wajib pada formulir pengajuan.' });
    }

    const timestamp = getCurrentTimestamp();
    const defaultStatus = 'Created';

    const newRow = [
      timestamp,
      tipeFormulir,
      namaKaryawan,
      nikUnitKerja,
      String(noRekening),
      atasNamaRekening,
      totalDitransfer,
      cicilan || '-',
      keterangan || '-',
      defaultStatus
    ];

    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A:J",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [newRow]
      }
    });

    return res.json({
      success: true,
      message: 'Pengajuan pinjaman/biaya berhasil disimpan dengan status Created.',
      data: {
        timestamp,
        tipeFormulir,
        namaKaryawan,
        status: defaultStatus
      }
    });
  } catch (error) {
    console.error('Error /api/pengajuan:', error);
    return res.status(500).json({ success: false, message: 'Gagal menyimpan data pengajuan: ' + error.message });
  }
});

/**
 * PUT /api/pengajuan/status
 * Body: { rowIndex, newStatus, userRole }
 */
app.put('/api/pengajuan/status', async (req, res) => {
  try {
    const { rowIndex, newStatus, userRole } = req.body;

    if (!rowIndex || !newStatus) {
      return res.status(400).json({ success: false, message: 'Index baris (rowIndex) dan status baru wajib disediakan.' });
    }

    const validStatuses = ['Created', 'On Process', 'Done', 'Rejected'];
    if (!validStatuses.includes(newStatus)) {
      return res.status(400).json({ success: false, message: 'Status tidak valid. Pilihan: ' + validStatuses.join(', ') });
    }

    if (userRole === 'User') {
      return res.status(403).json({ success: false, message: 'Anggota (User) tidak memiliki hak untuk mengubah status pengajuan.' });
    }

    if (userRole === 'Pengurus') {
      if (!['On Process', 'Rejected'].includes(newStatus)) {
        return res.status(403).json({
          success: false,
          message: 'Role Pengurus hanya memiliki otoritas untuk Persetujuan (On Process) atau Penolakan (Rejected).'
        });
      }
    }

    if (userRole === 'Financial') {
      if (!['Done', 'On Process'].includes(newStatus)) {
        return res.status(403).json({
          success: false,
          message: 'Role Financial berwenang memvalidasi transfer selesai (Done) atau memproses status pengajuan.'
        });
      }
    }

    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!J" + rowIndex,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[newStatus]]
      }
    });

    return res.json({
      success: true,
      message: 'Status pengajuan baris ke-' + rowIndex + ' berhasil diperbarui menjadi ' + newStatus + '.'
    });
  } catch (error) {
    console.error('Error /api/pengajuan/status:', error);
    return res.status(500).json({ success: false, message: 'Gagal memperbarui status pengajuan: ' + error.message });
  }
});

// Fallback route to serve frontend index.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Export Express app for Vercel Serverless Functions
module.exports = app;

// Start local server if not running on Vercel
if (process.env.VERCEL !== '1') {
  app.listen(PORT, () => {
    console.log('===================================================');
    console.log(' Portal Koperasi Server running on port ' + PORT);
    console.log(' URL: http://localhost:' + PORT);
    console.log(' Spreadsheet ID: ' + (SPREADSHEET_ID ? SPREADSHEET_ID : 'BELUM DI-SET'));
    console.log(' Credentials: ' + CREDENTIALS_PATH);
    console.log('===================================================');
  });
}