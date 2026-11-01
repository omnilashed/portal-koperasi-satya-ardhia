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
      let rawEnv = process.env.GOOGLE_CREDENTIALS_JSON.trim();
      let credsObj;
      try {
        credsObj = JSON.parse(rawEnv);
      } catch (e1) {
        try {
          credsObj = JSON.parse(rawEnv.replace(/\r?\n/g, '\\n'));
        } catch (e2) {
          throw new Error('Format GOOGLE_CREDENTIALS_JSON tidak valid: ' + e1.message);
        }
      }

      if (typeof credsObj === 'string') {
        credsObj = JSON.parse(credsObj);
      }

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
      console.error('[Google Auth Error (GOOGLE_CREDENTIALS_JSON)]:', envError.message);
      if (!fs.existsSync(resolvedPath)) {
        throw envError;
      }
    }
  }

  // Attempt 2: Use local credentials.json directly (for local development)
  try {
    if (!fs.existsSync(resolvedPath)) {
      throw new Error('Kredensial Google belum disiapkan. Di Vercel, pastikan Environment Variable GOOGLE_CREDENTIALS_JSON sudah diisi.');
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

// Helper: Tanggal dan Waktu Realtime Zona Jakarta / Bangkok (WIB / UTC+7)
function getJakartaDateTime(d = new Date()) {
  const formatter = new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Jakarta',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  const parts = formatter.formatToParts(d);
  const partMap = {};
  for (const part of parts) {
    partMap[part.type] = part.value;
  }
  const year = partMap.year;
  const month = partMap.month;
  const day = partMap.day;
  const hour = parseInt(partMap.hour, 10);
  const minute = parseInt(partMap.minute, 10);
  const second = parseInt(partMap.second, 10);
  const pad = (n) => String(n).padStart(2, '0');

  const dateStr = `${year}-${month}-${day}`;
  const timeStr = `${pad(hour)}:${pad(minute)}:${pad(second)}`;
  const timestampStr = `${day}/${month}/${year} ${pad(hour)}:${pad(minute)}`;
  const totalMinutes = hour * 60 + minute;

  return {
    year,
    month,
    day,
    hour,
    minute,
    second,
    totalMinutes,
    dateStr,
    timeStr,
    timestampStr
  };
}

// Helper: Format tanggal dd/MM/yyyy HH:mm (WIB / UTC+7)
function getCurrentTimestamp() {
  return getJakartaDateTime().timestampStr;
}

// Helper: Normalisasi & Pengecekan Hierarki Role
function normalizeRole(r) {
  const s = String(r || '').trim().toLowerCase();
  if (s.includes('ketua')) return 'ketua koperasi';
  if (s.includes('admin')) return 'admin';
  if (s.includes('financial') || s.includes('keuangan')) return 'financial';
  if (s.includes('pengurus')) return 'pengurus';
  if (s.includes('executive car') || s.includes('executivecar') || s.includes('executive')) return 'executive car';
  if (s.includes('air-minum') || s.includes('air minum') || s.includes('air_minum') || s.includes('airminum') || s.includes('kurir air')) return 'air-minum-sender';
  if (s.includes('pegawai') || s.includes('karyawan')) return 'pegawai';
  return 'user';
}
function isRoleAdmin(r) { return normalizeRole(r) === 'admin'; }
function isRoleKetua(r) { return normalizeRole(r) === 'ketua koperasi'; }
function isRolePengurus(r) { return normalizeRole(r) === 'pengurus'; }
function isRoleFinancial(r) { return normalizeRole(r) === 'financial'; }
function isRolePegawai(r) { return normalizeRole(r) === 'pegawai'; }
function isRoleExecutiveCar(r) { return normalizeRole(r) === 'executive car'; }
function isRoleAirMinumSender(r) { return normalizeRole(r) === 'air-minum-sender'; }
function isRoleAnggota(r) { return normalizeRole(r) === 'user'; }

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
// 2. ENDPOINTS: MANAJEMEN USER (ADMIN & KETUA KOPERASI)
// ==========================================

/**
 * GET /api/users
 * Query: ?role=Admin (atau Ketua Koperasi)
 */
app.get('/api/users', async (req, res) => {
  try {
    const { role } = req.query;
    if (role && !isRoleAdmin(role) && !isRoleKetua(role)) {
      return res.status(403).json({ success: false, message: 'Akses ditolak: Hanya Admin dan Ketua Koperasi yang dapat mengakses data pengguna.' });
    }

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
 * Body: { username, status, requestingRole }  --> status: 'Active' | 'Rejected' | 'Deactive' | 'Pending'
 */
app.put('/api/users/status', async (req, res) => {
  try {
    const { username, status, requestingRole } = req.body;
    if (!username || !status) {
      return res.status(400).json({ success: false, message: 'Username dan status baru wajib disediakan.' });
    }

    if (requestingRole && !isRoleAdmin(requestingRole) && !isRoleKetua(requestingRole)) {
      return res.status(403).json({ success: false, message: 'Akses ditolak: Hanya Admin dan Ketua Koperasi yang dapat mengubah status akun.' });
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

/**
 * PUT /api/users/role
 * Body: { username, newRole, requestingRole }
 * Khusus: Admin & Ketua Koperasi untuk mengubah peran user (Pegawai, Financial, Pengurus, Ketua Koperasi, Admin, User)
 */
app.put('/api/users/role', async (req, res) => {
  try {
    const { username, newRole, requestingRole } = req.body;
    if (!username || !newRole) {
      return res.status(400).json({ success: false, message: 'Username dan Role baru wajib disediakan.' });
    }

    if (requestingRole && !isRoleAdmin(requestingRole) && !isRoleKetua(requestingRole)) {
      return res.status(403).json({ success: false, message: 'Akses ditolak: Hanya Admin dan Ketua Koperasi yang dapat mengubah role akun.' });
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

    // Update Kolom C (Role)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Users!C' + targetRowIndex,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[newRole]]
      }
    });

    return res.json({
      success: true,
      message: 'Role user ' + username + ' berhasil diperbarui menjadi ' + newRole + '.'
    });
  } catch (error) {
    console.error('Error /api/users/role:', error);
    return res.status(500).json({ success: false, message: 'Gagal memperbarui role user: ' + error.message });
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

    // Jika role adalah User/Anggota atau Pegawai, filter data milik sendiri secara akurat
    if (isRoleAnggota(role) || isRolePegawai(role)) {
      const userNik = (nik || '').trim().toLowerCase();
      const userNama = (nama || '').trim().toLowerCase();

      // Ekstrak angka unik jika NIK mengandung nomor identitas (misal: "20246808")
      const nikNumberMatch = userNik.match(/\d{4,}/);
      const nikOnly = nikNumberMatch ? nikNumberMatch[0] : '';

      list = list.filter(item => {
        const itemNik = (item.nikUnitKerja || '').toLowerCase();
        const itemNama = (item.namaKaryawan || '').toLowerCase();

        // 1. Cek kesesuaian Nama Lengkap (Karyawan)
        const isNameMatch = Boolean(
          userNama && (
            itemNama === userNama ||
            itemNama.includes(userNama) ||
            userNama.includes(itemNama)
          )
        );

        // 2. Cek kesesuaian NIK angka unik (jika ada nomor minimal 4 digit)
        let isNikMatch = false;
        if (nikOnly) {
          isNikMatch = itemNik.includes(nikOnly);
        } else if (userNik && !['karyawan koperasi', 'koperasi', 'unit', 'staff', 'avsec', 'arff', '-'].includes(userNik)) {
          isNikMatch = itemNik === userNik || itemNik.startsWith(userNik + ' ') || itemNik.endsWith(' ' + userNik);
        }

        // Jika Nama cocok, sertakan data
        if (isNameMatch) return true;

        // Jika hanya NIK yang cocok tanpa nama, hanya izinkan jika NIK berupa nomor identitas unik (bukan teks nama unit)
        if (isNikMatch && nikOnly) {
          return true;
        }

        return false;
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
 * Body: { tipeFormulir, namaKaryawan, nikUnitKerja, noRekening, atasNamaRekening, totalDitransfer, cicilan, keterangan, userRole }
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
      keterangan,
      userRole
    } = req.body;

    // Role Guard: Hanya Admin, Ketua Koperasi, dan Financial yang dapat menginput pengajuan pinjaman
    if (userRole && !isRoleAdmin(userRole) && !isRoleKetua(userRole) && !isRoleFinancial(userRole)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Hanya Admin, Ketua Koperasi, dan Financial yang dapat menginput pengajuan pinjaman/biaya.'
      });
    }

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

    if (isRoleAnggota(userRole) || isRolePegawai(userRole)) {
      return res.status(403).json({ success: false, message: 'Akses ditolak: Anggota dan Pegawai tidak memiliki hak untuk mengubah status pengajuan.' });
    }

    if (isRolePengurus(userRole) || isRoleKetua(userRole)) {
      if (!['On Process', 'Rejected'].includes(newStatus)) {
        return res.status(403).json({
          success: false,
          message: 'Role Pengurus dan Ketua Koperasi hanya memiliki otoritas untuk Persetujuan (On Process) atau Penolakan (Rejected).'
        });
      }
    }

    if (isRoleFinancial(userRole)) {
      if (!['Done', 'On Process'].includes(newStatus)) {
        return res.status(403).json({
          success: false,
          message: 'Role Financial hanya dapat memproses pembayaran (Done) atau pengajuan diproses (On Process).'
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

/**
 * PUT /api/pengajuan
 * Body: { rowIndex, role, tipeFormulir, namaKaryawan, nikUnitKerja, noRekening, atasNamaRekening, totalDitransfer, cicilan, keterangan, status }
 * Khusus Admin: Edit data pengajuan
 */
app.put('/api/pengajuan', async (req, res) => {
  try {
    const {
      rowIndex,
      role,
      tipeFormulir,
      namaKaryawan,
      nikUnitKerja,
      noRekening,
      atasNamaRekening,
      totalDitransfer,
      cicilan,
      keterangan,
      status
    } = req.body;

    if (!rowIndex || parseInt(rowIndex, 10) < 2) {
      return res.status(400).json({ success: false, message: 'Index baris (rowIndex) tidak valid.' });
    }

    if (role !== 'Admin') {
      return res.status(403).json({ success: false, message: 'Hanya Administrator yang memiliki akses untuk mengubah data pengajuan.' });
    }

    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'Form Responses 1'!B${rowIndex}:J${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          tipeFormulir || '',
          namaKaryawan || '',
          nikUnitKerja || '',
          String(noRekening || ''),
          atasNamaRekening || '',
          totalDitransfer || '0',
          cicilan || '-',
          keterangan || '-',
          status || 'Created'
        ]]
      }
    });

    return res.json({
      success: true,
      message: 'Data pengajuan baris ke-' + rowIndex + ' berhasil diperbarui.'
    });
  } catch (error) {
    console.error('Error PUT /api/pengajuan:', error);
    return res.status(500).json({ success: false, message: 'Gagal memperbarui data pengajuan: ' + error.message });
  }
});

/**
 * DELETE /api/pengajuan
 * Body: { rowIndex, role }
 * Khusus Admin: Hapus pengajuan dari spreadsheet
 */
app.delete('/api/pengajuan', async (req, res) => {
  try {
    const { rowIndex, role } = req.body;

    if (!rowIndex || parseInt(rowIndex, 10) < 2) {
      return res.status(400).json({ success: false, message: 'Index baris (rowIndex) tidak valid.' });
    }

    if (role !== 'Admin') {
      return res.status(403).json({ success: false, message: 'Hanya Administrator yang memiliki akses untuk menghapus data pengajuan.' });
    }

    const targetRow = parseInt(rowIndex, 10);
    const sheets = await getSheetsClient();

    const meta = await sheets.spreadsheets.get({
      spreadsheetId: SPREADSHEET_ID
    });
    const targetSheet = (meta.data.sheets || []).find(s => s.properties.title === 'Form Responses 1');
    const sheetId = targetSheet ? targetSheet.properties.sheetId : 0;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          {
            deleteDimension: {
              range: {
                sheetId: sheetId,
                dimension: 'ROWS',
                startIndex: targetRow - 1,
                endIndex: targetRow
              }
            }
          }
        ]
      }
    });

    return res.json({
      success: true,
      message: 'Data pengajuan baris ke-' + rowIndex + ' berhasil dihapus.'
    });
  } catch (error) {
    console.error('Error DELETE /api/pengajuan:', error);
    return res.status(500).json({ success: false, message: 'Gagal menghapus pengajuan: ' + error.message });
  }
});

// ==========================================
// 4. ENDPOINTS: MODUL ABSENSI & GEOFENCING HALIM PERDANAKUSUMA
// ==========================================

// Titik Terminal Bandara Internasional Halim Perdanakusuma (HLP)
const HALIM_TERMINAL_LAT = -6.2652;
const HALIM_TERMINAL_LNG = 106.8851;
const GEOFENCE_RADIUS_METERS = 1000; // Radius absensi 1 km (1.000 meter)

function calculateHaversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000; // meter
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c);
}

// Helper: Memastikan Sheet Tertentu Ada Beserta Header
async function ensureSheetExists(sheets, title, headers) {
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const existingTitles = (meta.data.sheets || []).map(s => s.properties.title);
    if (!existingTitles.includes(title)) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: {
          requests: [{
            addSheet: {
              properties: { title }
            }
          }]
        }
      });
      if (headers && headers.length > 0) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${title}'!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [headers] }
        });
      }
      console.log(`[Google Sheets] Berhasil membuat sheet on-demand: '${title}'`);
    }
  } catch (err) {
    console.warn(`[ensureSheetExists Warning for ${title}]:`, err.message);
  }
}

// Otomatis Cek & Buat Semua Sheet Di Google Spreadsheet
async function ensureAllSheetsExist() {
  if (!SPREADSHEET_ID) return;
  try {
    const sheets = await getSheetsClient();
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const existingTitles = (meta.data.sheets || []).map(s => s.properties.title);

    const requiredSheets = [
      {
        title: 'Users',
        headers: ['Username', 'Password', 'Role', 'Nama Lengkap', 'NIK', 'Email', 'ResetToken', 'AccountStatus']
      },
      {
        title: 'Form Responses 1',
        headers: ['Timestamp', 'Tipe Formulir', 'Nama Karyawan / Anggota', 'NIK - Unit Kerja', 'No. Rekening', 'Atas Nama Rekening', 'Total Ditransfer (Rp)', 'Cicilan', 'Keterangan / Keperluan', 'Status']
      },
      {
        title: 'absensi-harian',
        headers: ['Timestamp', 'Tanggal', 'NIK', 'Nama Lengkap', 'Unit Kerja', 'Clock-In', 'Clock-Out', 'Status Kehadiran', 'Lokasi Lat/Long', 'Jarak Terminal Halim (meter)', 'Status Verifikasi Wajah']
      },
      {
        title: 'wajah-karyawan',
        headers: ['NIK', 'Nama Lengkap', 'FaceDescriptor', 'Waktu Pendaftaran']
      },
      {
        title: 'operasional-executive-car',
        headers: ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Pengguna', 'Detail Pengguna', 'Jumlah Pax', 'Catatan']
      },
      {
        title: 'operasional-air-minum',
        headers: ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Penerima', 'Nama Tenant / Stakeholder', 'Jumlah Galon', 'Metode Pembayaran', 'Catatan']
      }
    ];

    for (const s of requiredSheets) {
      if (!existingTitles.includes(s.title)) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            requests: [{
              addSheet: {
                properties: { title: s.title }
              }
            }]
          }
        });

        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'${s.title}'!A1`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [s.headers]
          }
        });
        console.log(`[Google Sheets] Berhasil membuat sheet otomatis: '${s.title}'`);
      }
    }
  } catch (err) {
    console.error('[Google Sheets Auto-Setup Warning]:', err.message);
  }
}

/**
 * GET /api/absensi/config
 * Mengembalikan data koordinat patokan & radius geofencing Halim Perdanakusuma
 */
app.get('/api/absensi/config', (req, res) => {
  return res.json({
    success: true,
    data: {
      locationName: 'Terminal Bandara Internasional Halim Perdanakusuma (HLP)',
      halimLat: HALIM_TERMINAL_LAT,
      halimLng: HALIM_TERMINAL_LNG,
      radiusMeters: GEOFENCE_RADIUS_METERS
    }
  });
});

/**
 * GET /api/absensi/face-descriptors
 * Mengambil seluruh data descriptor wajah karyawan untuk face matching di browser
 */
app.get('/api/absensi/face-descriptors', async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'wajah-karyawan'!A2:D"
    });

    const rows = response.data.values || [];
    const descriptors = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      nik: String(r[0] || '').trim(),
      nama: String(r[1] || '').trim(),
      descriptor: r[2] || '',
      updatedAt: r[3] || ''
    })).filter(item => item.nik && item.descriptor);

    return res.json({
      success: true,
      data: descriptors
    });
  } catch (error) {
    console.error('Error /api/absensi/face-descriptors:', error);
    return res.status(500).json({ success: false, message: 'Gagal mengambil data descriptor wajah: ' + error.message });
  }
});

/**
 * POST /api/absensi/register-face
 * Mendaftarkan descriptor wajah (128-float embedding) ke sheet wajah-karyawan
 */
app.post('/api/absensi/register-face', async (req, res) => {
  try {
    const { nik, nama, descriptor, userRole, role } = req.body;
    const currentRole = userRole || role;

    if (currentRole && !isRoleAdmin(currentRole) && !isRoleFinancial(currentRole) && !isRolePegawai(currentRole) && !isRoleExecutiveCar(currentRole) && !isRoleAirMinumSender(currentRole)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Pendaftaran wajah hanya diperuntukkan bagi Pegawai, Financial, Executive Car, Air Minum Sender, dan Admin.'
      });
    }

    if (!nik || !descriptor) {
      return res.status(400).json({ success: false, message: 'NIK dan data descriptor wajah wajib disertakan.' });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'wajah-karyawan'!A2:D"
    });

    const rows = response.data.values || [];
    const cleanNik = String(nik).trim().toLowerCase();
    let foundRowIndex = -1;

    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0] || '').trim().toLowerCase() === cleanNik) {
        foundRowIndex = i + 2;
        break;
      }
    }

    const timestamp = getCurrentTimestamp();
    const descriptorString = typeof descriptor === 'string' ? descriptor : JSON.stringify(descriptor);

    if (foundRowIndex > 0) {
      // Update data wajah yang sudah ada
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'wajah-karyawan'!B${foundRowIndex}:D${foundRowIndex}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[nama || '', descriptorString, timestamp]]
        }
      });
    } else {
      // Tambah data wajah baru
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "'wajah-karyawan'!A:D",
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[nik, nama || '', descriptorString, timestamp]]
        }
      });
    }

    return res.json({
      success: true,
      message: `Wajah karyawan ${nama || nik} berhasil didaftarkan ke sistem!`
    });
  } catch (error) {
    console.error('Error /api/absensi/register-face:', error);
    return res.status(500).json({ success: false, message: 'Gagal mendaftarkan wajah: ' + error.message });
  }
});

/**
 * POST /api/absensi/clock-in
 * Mencatat jam masuk karyawan dengan validasi geofencing 1 km dari Terminal Bandara Halim
 */
app.post('/api/absensi/clock-in', async (req, res) => {
  try {
    const { nik, nama, unitKerja, latitude, longitude, faceVerified, userRole, role } = req.body;
    let currentRole = userRole || role;

    const sheets = await getSheetsClient();
    if (!currentRole && nik) {
      try {
        const usersResp = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: 'Users!A2:H'
        });
        const uRows = usersResp.data.values || [];
        const cNik = String(nik).trim().toLowerCase();
        const matched = uRows.find(u => String(u[4] || '').trim().toLowerCase() === cNik || String(u[0] || '').trim().toLowerCase() === cNik);
        if (matched) currentRole = matched[2];
      } catch (e) {
        console.warn('Role lookup warning in clock-in:', e.message);
      }
    }

    if (currentRole && !isRoleAdmin(currentRole) && !isRoleFinancial(currentRole) && !isRolePegawai(currentRole) && !isRoleExecutiveCar(currentRole) && !isRoleAirMinumSender(currentRole)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Fitur presensi clock-in hanya dapat digunakan oleh Pegawai, Financial, Executive Car, Air Minum Sender, dan Admin.'
      });
    }

    if (!nik || latitude === undefined || longitude === undefined) {
      return res.status(400).json({ success: false, message: 'NIK dan data koordinat GPS wajib disertakan.' });
    }

    const userLat = parseFloat(latitude);
    const userLng = parseFloat(longitude);
    const distanceMeters = calculateHaversineDistance(userLat, userLng, HALIM_TERMINAL_LAT, HALIM_TERMINAL_LNG);

    // Validasi Geofencing 1 km
    if (distanceMeters > GEOFENCE_RADIUS_METERS) {
      return res.status(400).json({
        success: false,
        message: `Di luar radius absensi! Jarak Anda saat ini ${distanceMeters.toLocaleString('id-ID')} meter dari Terminal Bandara Halim Perdanakusuma (Maksimal radius 1.000 meter / 1 KM).`
      });
    }

    
    if (!faceVerified || faceVerified === false || faceVerified === 'false') {
      return res.status(400).json({
        success: false,
        message: 'Presensi ditolak: Verifikasi biometrik wajah tidak valid atau wajah tidak cocok dengan data biometrik akun Anda!'
      });
    }

    // Best Practice: Validasi akun wajib sudah mendaftarkan wajah di sheet wajah-karyawan
    const wajahResponse = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'wajah-karyawan'!A2:D"
    });
    const cleanNikCheck = String(nik).trim().toLowerCase();
    const cleanNamaCheck = String(nama || '').trim().toLowerCase();
    const wajahRows = wajahResponse.data.values || [];
    const isEnrolled = wajahRows.some(r => {
      const rNik = String(r[0] || '').trim().toLowerCase();
      const rNama = String(r[1] || '').trim().toLowerCase();
      return (cleanNikCheck && rNik === cleanNikCheck) || (cleanNamaCheck && rNama === cleanNamaCheck);
    });

    if (!isEnrolled) {
      return res.status(400).json({
        success: false,
        message: 'Presensi ditolak: Akun Anda belum mendaftarkan biometrik wajah! Silakan lakukan pendaftaran wajah terlebih dahulu.'
      });
    }
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'absensi-harian'!A2:K"
    });

    const jkt = getJakartaDateTime();
    const dateStr = jkt.dateStr;
    const timeStr = jkt.timeStr;
    const timestampStr = jkt.timestampStr;
    const hour = jkt.hour;
    const minute = jkt.minute;
    const totalMinutes = jkt.totalMinutes;

    const cleanNik = String(nik).trim().toLowerCase();
    const rows = response.data.values || [];

    // Cek apakah sudah Clock-In hari ini
    for (let i = 0; i < rows.length; i++) {
      const rowTanggal = String(rows[i][1] || '').trim();
      const rowNik = String(rows[i][2] || '').trim().toLowerCase();

      if (rowTanggal === dateStr && rowNik === cleanNik) {
        return res.status(400).json({
          success: false,
          message: `Anda sudah melakukan Clock-In hari ini pada pukul ${rows[i][5] || '-'} WIB.`
        });
      }
    }

    // Pengecekan Jam Kerja & Penentuan Status Kehadiran Berdasarkan Role (WIB / UTC+7)
    let statusKehadiran = 'Hadir Tepat Waktu';

    if (isRoleExecutiveCar(currentRole)) {
      // Role Executive Car: bisa clockin dari jam 3 subuh (03:00 WIB) sampai jam 11 malam (23:00 WIB), asumsi kerja 12 jam
      if (hour < 3) {
        return res.status(400).json({
          success: false,
          message: `Belum memasuki jam operasional Executive Car! Clock-In dibuka mulai pukul 03:00 Subuh WIB (Waktu saat ini: ${timeStr} WIB).`
        });
      }
      if (hour >= 23) {
        return res.status(400).json({
          success: false,
          message: `Batas waktu Clock-In Executive Car telah berakhir (Maksimal pukul 23:00 WIB). Waktu saat ini: ${timeStr} WIB.`
        });
      }
      statusKehadiran = 'Hadir (Executive Car - 12 Jam)';
    } else if (isRoleAirMinumSender(currentRole)) {
      // Role Air Minum Sender: jam kerja pengantaran galon
      if (totalMinutes <= (8 * 60 + 30)) {
        statusKehadiran = 'Hadir Tepat Waktu (Pengantar Air Minum)';
      } else if (totalMinutes <= (16 * 60)) {
        statusKehadiran = 'Terlambat (Pengantar Air Minum)';
      } else {
        statusKehadiran = 'Terlambat (Lewat Jam Kantor 16:00)';
      }
    } else if (isRolePegawai(currentRole) || isRoleFinancial(currentRole)) {
      // Role Finance / Pegawai: jam 8 pagi sampai jam 4 sore (08:00 - 16:00 WIB)
      if (totalMinutes <= (8 * 60 + 30)) {
        statusKehadiran = 'Hadir Tepat Waktu';
      } else if (totalMinutes <= (16 * 60)) {
        statusKehadiran = 'Terlambat';
      } else {
        statusKehadiran = 'Terlambat (Lewat Jam Kantor 16:00)';
      }
    } else {
      // Admin / Lainnya
      statusKehadiran = totalMinutes <= (8 * 60 + 30) ? 'Hadir Tepat Waktu' : 'Terlambat';
    }
    const faceStatus = faceVerified ? 'Terverifikasi Face-API ✅' : 'Wajah Tidak Dikenali ❌';
    const lokasiStr = `Lat: ${userLat.toFixed(6)}, Lng: ${userLng.toFixed(6)}`;

    const newRow = [
      timestampStr,
      dateStr,
      nik,
      nama || '-',
      unitKerja || '-',
      timeStr,
      '-', // Clock-Out belum
      statusKehadiran,
      lokasiStr,
      distanceMeters,
      faceStatus
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'absensi-harian'!A:K",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [newRow]
      }
    });

    return res.json({
      success: true,
      message: `Clock-In berhasil dicatat pada pukul ${timeStr} WIB (${statusKehadiran})!`,
      data: {
        tanggal: dateStr,
        clockIn: timeStr,
        statusKehadiran,
        distanceMeters,
        faceStatus
      }
    });
  } catch (error) {
    console.error('Error /api/absensi/clock-in:', error);
    return res.status(500).json({ success: false, message: 'Gagal melakukan Clock-In: ' + error.message });
  }
});

/**
 * POST /api/absensi/clock-out
 * Mencatat jam pulang karyawan ke sheet absensi-harian
 */
app.post('/api/absensi/clock-out', async (req, res) => {
  try {
    const { nik, latitude, longitude, userRole, role, faceVerified } = req.body;
    let currentRole = userRole || role;

    const sheets = await getSheetsClient();
    if (!currentRole && nik) {
      try {
        const usersResp = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: 'Users!A2:H'
        });
        const uRows = usersResp.data.values || [];
        const cNik = String(nik).trim().toLowerCase();
        const matched = uRows.find(u => String(u[4] || '').trim().toLowerCase() === cNik || String(u[0] || '').trim().toLowerCase() === cNik);
        if (matched) currentRole = matched[2];
      } catch (e) {
        console.warn('Role lookup warning in clock-out:', e.message);
      }
    }

    if (currentRole && !isRoleAdmin(currentRole) && !isRoleFinancial(currentRole) && !isRolePegawai(currentRole) && !isRoleExecutiveCar(currentRole) && !isRoleAirMinumSender(currentRole)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Fitur presensi clock-out hanya dapat digunakan oleh Pegawai, Financial, Executive Car, Air Minum Sender, dan Admin.'
      });
    }

    if (!nik || latitude === undefined || longitude === undefined) {
      return res.status(400).json({ success: false, message: 'NIK dan data koordinat GPS wajib disertakan.' });
    }

    if (faceVerified !== undefined && (faceVerified === false || faceVerified === 'false')) {
      return res.status(400).json({ success: false, message: 'Clock-Out ditolak: Verifikasi biometrik wajah tidak cocok dengan akun Anda!' });
    }

    const userLat = parseFloat(latitude);
    const userLng = parseFloat(longitude);
    const distanceMeters = calculateHaversineDistance(userLat, userLng, HALIM_TERMINAL_LAT, HALIM_TERMINAL_LNG);

    // Validasi Geofencing 1 km
    if (distanceMeters > GEOFENCE_RADIUS_METERS) {
      return res.status(400).json({
        success: false,
        message: `Di luar radius absensi! Jarak Anda saat ini ${distanceMeters.toLocaleString('id-ID')} meter dari Terminal Bandara Halim Perdanakusuma (Maksimal radius 1.000 meter / 1 KM).`
      });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'absensi-harian'!A2:K"
    });

    const jkt = getJakartaDateTime();
    const dateStr = jkt.dateStr;
    const timeStr = jkt.timeStr;
    const hour = jkt.hour;
    const minute = jkt.minute;
    const totalMinutes = jkt.totalMinutes;

    const cleanNik = String(nik).trim().toLowerCase();
    const rows = response.data.values || [];
    let foundRowIndex = -1;
    let existingClockIn = '';

    for (let i = 0; i < rows.length; i++) {
      const rowTanggal = String(rows[i][1] || '').trim();
      const rowNik = String(rows[i][2] || '').trim().toLowerCase();

      if (rowTanggal === dateStr && rowNik === cleanNik) {
        foundRowIndex = i + 2;
        existingClockIn = String(rows[i][5] || '').trim();
        break;
      }
    }

    if (foundRowIndex === -1) {
      return res.status(400).json({
        success: false,
        message: 'Anda belum melakukan Clock-In hari ini! Silakan Clock-In terlebih dahulu.'
      });
    }

    // ================= VALIDASI WAJIB ISI DATA OPERASIONAL SEBELUM CLOCK-OUT =================
    // a. Role Executive Car: wajib isi data penggunaan mobil harian
    if (isRoleExecutiveCar(currentRole)) {
      await ensureSheetExists(sheets, 'operasional-executive-car', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Pengguna', 'Detail Pengguna', 'Jumlah Pax', 'Catatan']);
      const opCarResp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'operasional-executive-car'!A2:I"
      });
      const opCarRows = opCarResp.data.values || [];
      const hasCarLogToday = opCarRows.some(r => {
        const rTgl = String(r[1] || '').trim();
        const rNik = String(r[3] || '').trim().toLowerCase();
        return rTgl === dateStr && rNik === cleanNik;
      });

      if (!hasCarLogToday) {
        return res.status(400).json({
          success: false,
          isOperationalRequired: true,
          operationalType: 'executive-car',
          message: 'Clock-Out ditolak: Anda wajib mengisi Data Penggunaan Executive Car Harian sebelum melakukan Clock-Out! Silakan isi data di menu Data Operasional.'
        });
      }
    }

    // b. Role Air Minum Sender: wajib isi data pengantaran galon hari ini sebelum pulang
    if (isRoleAirMinumSender(currentRole)) {
      await ensureSheetExists(sheets, 'operasional-air-minum', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Penerima', 'Nama Tenant / Stakeholder', 'Jumlah Galon', 'Metode Pembayaran', 'Catatan']);
      const opAirResp = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'operasional-air-minum'!A2:J"
      });
      const opAirRows = opAirResp.data.values || [];
      const hasAirLogToday = opAirRows.some(r => {
        const rTgl = String(r[1] || '').trim();
        const rNik = String(r[3] || '').trim().toLowerCase();
        return rTgl === dateStr && rNik === cleanNik;
      });

      if (!hasAirLogToday) {
        return res.status(400).json({
          success: false,
          isOperationalRequired: true,
          operationalType: 'air-minum',
          message: 'Clock-Out ditolak: Anda wajib mengisi Data Pengantaran Galon Air Minum hari ini sebelum melakukan Clock-Out (pulang)! Silakan isi data di menu Data Operasional.'
        });
      }
    }

    // Hitung durasi kerja dari Clock-In
    let durationHours = 0;
    let durationMinutes = 0;
    if (existingClockIn && existingClockIn.includes(':')) {
      const parts = existingClockIn.split(':');
      const inHour = parseInt(parts[0], 10) || 0;
      const inMin = parseInt(parts[1], 10) || 0;
      const inTotal = inHour * 60 + inMin;
      let diffMins = totalMinutes - inTotal;
      if (diffMins < 0) diffMins = 0;
      durationHours = Math.floor(diffMins / 60);
      durationMinutes = diffMins % 60;
    }

    let outMessage = `Clock-Out berhasil dicatat pada pukul ${timeStr} WIB. Terima kasih atas kerja keras Anda hari ini!`;
    if (isRoleExecutiveCar(currentRole)) {
      // Role Executive Car: clock out sampai jam 11 malam (23:00 WIB), asumsi kerja 12 jam
      if (hour > 23 || (hour === 23 && minute > 0)) {
        outMessage = `Clock-Out berhasil pada pukul ${timeStr} WIB (Melewati batas reguler 23:00 WIB. Durasi kerja: ${durationHours} Jam ${durationMinutes} Menit).`;
      } else if (durationHours >= 12) {
        outMessage = `Clock-Out berhasil dicatat pada pukul ${timeStr} WIB (Durasi kerja: ${durationHours} Jam ${durationMinutes} Menit - Target 12 Jam Kerja Terpenuhi ✅). Terima kasih atas dedikasi Anda!`;
      } else {
        outMessage = `Clock-Out berhasil dicatat pada pukul ${timeStr} WIB (Durasi kerja: ${durationHours} Jam ${durationMinutes} Menit dari target 12 Jam). Terima kasih!`;
      }
    } else if (isRolePegawai(currentRole) || isRoleFinancial(currentRole)) {
      // Role Finance / Pegawai: jam 8 pagi sampai jam 4 sore (16:00 WIB)
      if (totalMinutes < (16 * 60)) {
        outMessage = `Clock-Out berhasil dicatat pada pukul ${timeStr} WIB (Catatan: Pulang lebih awal sebelum jam kantor 16:00 WIB). Terima kasih!`;
      } else {
        outMessage = `Clock-Out berhasil dicatat pada pukul ${timeStr} WIB. Jam kantor telah usai (16:00 WIB). Terima kasih atas dedikasi Anda hari ini!`;
      }
    }

    // Update Kolom G (Clock-Out)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'absensi-harian'!G${foundRowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[timeStr]]
      }
    });

    return res.json({
      success: true,
      message: outMessage,
      data: {
        tanggal: dateStr,
        clockOut: timeStr,
        distanceMeters,
        durationHours,
        durationMinutes
      }
    });
  } catch (error) {
    console.error('Error /api/absensi/clock-out:', error);
    return res.status(500).json({ success: false, message: 'Gagal melakukan Clock-Out: ' + error.message });
  }
});

/**
 * GET /api/absensi/today
 * Mengambil status presensi hari ini untuk NIK yang sedang login
 */
app.get('/api/absensi/today', async (req, res) => {
  try {
    const { nik } = req.query;
    if (!nik) return res.json({ success: true, data: null });

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'absensi-harian'!A2:K"
    });

    const jkt = getJakartaDateTime();
    const dateStr = jkt.dateStr;

    const cleanNik = String(nik).trim().toLowerCase();
    const rows = response.data.values || [];

    for (let i = rows.length - 1; i >= 0; i--) {
      const rowTanggal = String(rows[i][1] || '').trim();
      const rowNik = String(rows[i][2] || '').trim().toLowerCase();

      if (rowTanggal === dateStr && rowNik === cleanNik) {
        return res.json({
          success: true,
          data: {
            rowIndex: i + 2,
            timestamp: rows[i][0] || '',
            tanggal: rows[i][1] || '',
            nik: rows[i][2] || '',
            nama: rows[i][3] || '',
            unitKerja: rows[i][4] || '',
            clockIn: rows[i][5] || '-',
            clockOut: rows[i][6] || '-',
            statusKehadiran: rows[i][7] || '-',
            lokasi: rows[i][8] || '-',
            jarakMeters: rows[i][9] || '0',
            faceVerified: rows[i][10] || '-'
          }
        });
      }
    }

    return res.json({ success: true, data: null });
  } catch (error) {
    console.error('Error /api/absensi/today:', error);
    return res.status(500).json({ success: false, message: 'Gagal mengambil status absensi: ' + error.message });
  }
});

/**
 * GET /api/absensi/history
 * Mengambil riwayat data presensi
 * - Pegawai & Financial: data milik sendiri
 * - Admin, Ketua Koperasi, Pengurus: seluruh karyawan + filter tanggal/search
 */
app.get('/api/absensi/history', async (req, res) => {
  try {
    const { role, nik, nama, startDate, endDate, search, todayOnly } = req.query;

    if (isRoleAnggota(role)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Anggota koperasi tidak memiliki akses ke fitur absensi.'
      });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'absensi-harian'!A2:K"
    });

    const rows = response.data.values || [];
    let list = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      timestamp: r[0] || '',
      tanggal: r[1] || '',
      nik: r[2] || '',
      nama: r[3] || '',
      unitKerja: r[4] || '',
      clockIn: r[5] || '-',
      clockOut: r[6] || '-',
      statusKehadiran: r[7] || '-',
      lokasi: r[8] || '-',
      jarakMeters: r[9] || '0',
      faceVerified: r[10] || '-'
    }));

    // 1. Pegawai, Financial, dan Executive Car hanya melihat data absensi milik sendiri
    if (isRolePegawai(role) || isRoleFinancial(role) || isRoleExecutiveCar(role) || isRoleAirMinumSender(role)) {
      const userNik = (nik || '').trim().toLowerCase();
      const userNama = (nama || '').trim().toLowerCase();

      list = list.filter(item => {
        const itemNik = (item.nik || '').toLowerCase();
        const itemNama = (item.nama || '').toLowerCase();
        return (userNik && itemNik.includes(userNik)) || (userNama && itemNama.includes(userNama));
      });
    }

    // 2. Filter Tanggal: todayOnly atau rentang startDate s/d endDate
    const jkt = getJakartaDateTime();
    const todayStr = jkt.dateStr;

    if (todayOnly === 'true') {
      list = list.filter(item => String(item.tanggal).trim() === todayStr);
    } else {
      if (startDate) {
        list = list.filter(item => String(item.tanggal).trim() >= startDate);
      }
      if (endDate) {
        list = list.filter(item => String(item.tanggal).trim() <= endDate);
      }
    }

    // 3. Filter Search (Nama / NIK / Unit Kerja)
    if (search && search.trim()) {
      const q = search.trim().toLowerCase();
      list = list.filter(item => 
        (item.nama && item.nama.toLowerCase().includes(q)) ||
        (item.nik && item.nik.toLowerCase().includes(q)) ||
        (item.unitKerja && item.unitKerja.toLowerCase().includes(q))
      );
    }

    list.reverse(); // Terbaru di paling atas

    return res.json({
      success: true,
      data: list
    });
  } catch (error) {
    console.error('Error /api/absensi/history:', error);
    return res.status(500).json({ success: false, message: 'Gagal mengambil riwayat absensi: ' + error.message });
  }
});

/**
 * GET /api/absensi/today-employees
 * Mengambil rekap daftar seluruh pegawai yang masuk pada hari ini
 * Khusus: Admin, Ketua Koperasi, Pengurus
 */
app.get('/api/absensi/today-employees', async (req, res) => {
  try {
    const { role } = req.query;
    if (isRoleAnggota(role) || isRolePegawai(role) || isRoleFinancial(role) || isRoleExecutiveCar(role) || isRoleAirMinumSender(role)) {
      return res.status(403).json({
        success: false,
        message: 'Akses monitoring presensi hanya untuk Admin, Ketua Koperasi, dan Pengurus.'
      });
    }

    const sheets = await getSheetsClient();
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'absensi-harian'!A2:K"
    });

    const jkt = getJakartaDateTime();
    const todayStr = jkt.dateStr;

    const rows = response.data.values || [];
    const todayRows = rows
      .map((r, idx) => ({
        rowIndex: idx + 2,
        timestamp: r[0] || '',
        tanggal: r[1] || '',
        nik: r[2] || '',
        nama: r[3] || '',
        unitKerja: r[4] || '',
        clockIn: r[5] || '-',
        clockOut: r[6] || '-',
        statusKehadiran: r[7] || '-',
        lokasi: r[8] || '-',
        jarakMeters: r[9] || '0',
        faceVerified: r[10] || '-'
      }))
      .filter(item => String(item.tanggal).trim() === todayStr)
      .reverse();

    let tepatWaktu = 0;
    let terlambat = 0;
    todayRows.forEach(item => {
      if (item.statusKehadiran && item.statusKehadiran.toLowerCase().includes('tepat waktu')) {
        tepatWaktu++;
      } else {
        terlambat++;
      }
    });

    return res.json({
      success: true,
      date: todayStr,
      summary: {
        totalHadir: todayRows.length,
        tepatWaktu,
        terlambat
      },
      data: todayRows
    });
  } catch (error) {
    console.error('Error /api/absensi/today-employees:', error);
    return res.status(500).json({ success: false, message: 'Gagal memuat data pegawai masuk hari ini: ' + error.message });
  }
});


// ==========================================
// 6. ENDPOINTS: DATA OPERASIONAL ROLE
// ==========================================

/**
 * GET /api/operasional/status-today
 * Memeriksa apakah user (Executive Car / Air Minum Sender) sudah mengisi data operasional hari ini
 */
app.get('/api/operasional/status-today', async (req, res) => {
  try {
    const { role, nik } = req.query;
    if (!nik || !role) return res.json({ success: true, filledToday: false, count: 0 });

    const sheets = await getSheetsClient();
    const jkt = getJakartaDateTime();
    const todayStr = jkt.dateStr;
    const cleanNik = String(nik).trim().toLowerCase();

    if (isRoleExecutiveCar(role)) {
      await ensureSheetExists(sheets, 'operasional-executive-car', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Pengguna', 'Detail Pengguna', 'Jumlah Pax', 'Catatan']);
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'operasional-executive-car'!A2:I"
      });
      const rows = response.data.values || [];
      const userTodayRows = rows.filter(r => String(r[1] || '').trim() === todayStr && String(r[3] || '').trim().toLowerCase() === cleanNik);
      return res.json({
        success: true,
        type: 'executive-car',
        filledToday: userTodayRows.length > 0,
        count: userTodayRows.length,
        todayDate: todayStr
      });
    }

    if (isRoleAirMinumSender(role)) {
      await ensureSheetExists(sheets, 'operasional-air-minum', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Penerima', 'Nama Tenant / Stakeholder', 'Jumlah Galon', 'Metode Pembayaran', 'Catatan']);
      const response = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'operasional-air-minum'!A2:J"
      });
      const rows = response.data.values || [];
      const userTodayRows = rows.filter(r => String(r[1] || '').trim() === todayStr && String(r[3] || '').trim().toLowerCase() === cleanNik);
      return res.json({
        success: true,
        type: 'air-minum',
        filledToday: userTodayRows.length > 0,
        count: userTodayRows.length,
        todayDate: todayStr
      });
    }

    return res.json({ success: true, filledToday: true, count: 0 });
  } catch (error) {
    console.error('Error /api/operasional/status-today:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/operasional/executive-car
 * Menyimpan data harian penggunaan Executive Car
 */
app.post('/api/operasional/executive-car', async (req, res) => {
  try {
    const { nik, nama, kategoriPengguna, detailPengguna, jumlahPax, catatan } = req.body;

    if (!nik || !kategoriPengguna) {
      return res.status(400).json({ success: false, message: 'NIK dan Kategori Pengguna wajib diisi.' });
    }

    const paxNum = parseInt(jumlahPax, 10);
    if (isNaN(paxNum) || paxNum < 1) {
      return res.status(400).json({ success: false, message: 'Jumlah Pax harus berupa angka minimal 1.' });
    }

    if (kategoriPengguna === 'Lain-lain' && (!detailPengguna || !detailPengguna.trim())) {
      return res.status(400).json({ success: false, message: 'Harap tuliskan keterangan untuk kategori Lain-lain.' });
    }

    const sheets = await getSheetsClient();
    await ensureSheetExists(sheets, 'operasional-executive-car', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Pengguna', 'Detail Pengguna', 'Jumlah Pax', 'Catatan']);

    const jkt = getJakartaDateTime();
    const newRow = [
      jkt.timestampStr,
      jkt.dateStr,
      jkt.timeStr,
      String(nik).trim(),
      String(nama || '-').trim(),
      String(kategoriPengguna).trim(),
      String(detailPengguna || '-').trim(),
      paxNum,
      String(catatan || '-').trim()
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'operasional-executive-car'!A:I",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [newRow] }
    });

    return res.json({
      success: true,
      message: `Data operasional Executive Car (${kategoriPengguna} - ${paxNum} Pax) berhasil dicatat!`,
      data: {
        tanggal: jkt.dateStr,
        waktu: jkt.timeStr,
        kategoriPengguna,
        jumlahPax: paxNum
      }
    });
  } catch (error) {
    console.error('Error /api/operasional/executive-car:', error);
    return res.status(500).json({ success: false, message: 'Gagal mencatat data Executive Car: ' + error.message });
  }
});

/**
 * GET /api/operasional/executive-car
 * Mengambil riwayat penggunaan Executive Car
 */
app.get('/api/operasional/executive-car', async (req, res) => {
  try {
    const { role, nik, startDate, endDate, todayOnly } = req.query;
    const sheets = await getSheetsClient();
    await ensureSheetExists(sheets, 'operasional-executive-car', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Pengguna', 'Detail Pengguna', 'Jumlah Pax', 'Catatan']);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'operasional-executive-car'!A2:I"
    });

    const rows = response.data.values || [];
    let list = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      timestamp: r[0] || '',
      tanggal: r[1] || '',
      waktu: r[2] || '',
      nik: r[3] || '',
      nama: r[4] || '',
      kategoriPengguna: r[5] || '',
      detailPengguna: r[6] || '-',
      jumlahPax: parseInt(r[7], 10) || 0,
      catatan: r[8] || '-'
    }));

    // Driver hanya melihat miliknya sendiri
    if (isRoleExecutiveCar(role)) {
      const cleanNik = String(nik || '').trim().toLowerCase();
      list = list.filter(item => (item.nik || '').toLowerCase() === cleanNik);
    }

    const jkt = getJakartaDateTime();
    if (todayOnly === 'true') {
      list = list.filter(item => String(item.tanggal).trim() === jkt.dateStr);
    } else {
      if (startDate) list = list.filter(item => String(item.tanggal).trim() >= startDate);
      if (endDate) list = list.filter(item => String(item.tanggal).trim() <= endDate);
    }

    list.reverse();
    return res.json({ success: true, data: list });
  } catch (error) {
    console.error('Error GET /api/operasional/executive-car:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/operasional/air-minum
 * Menyimpan data harian pengantaran galon air minum
 */
app.post('/api/operasional/air-minum', async (req, res) => {
  try {
    const { nik, nama, kategoriPenerima, namaPenerima, jumlahGalon, metodePembayaran, catatan } = req.body;

    if (!nik || !kategoriPenerima) {
      return res.status(400).json({ success: false, message: 'NIK dan Kategori Penerima wajib diisi.' });
    }

    const galonNum = parseInt(jumlahGalon, 10);
    if (isNaN(galonNum) || galonNum < 1) {
      return res.status(400).json({ success: false, message: 'Jumlah Galon harus berupa angka minimal 1.' });
    }

    if ((kategoriPenerima === 'Tenant' || kategoriPenerima === 'Stakeholder') && (!namaPenerima || !namaPenerima.trim())) {
      return res.status(400).json({ success: false, message: `Harap tuliskan nama ${kategoriPenerima}.` });
    }

    const sheets = await getSheetsClient();
    await ensureSheetExists(sheets, 'operasional-air-minum', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Penerima', 'Nama Tenant / Stakeholder', 'Jumlah Galon', 'Metode Pembayaran', 'Catatan']);

    const jkt = getJakartaDateTime();
    const newRow = [
      jkt.timestampStr,
      jkt.dateStr,
      jkt.timeStr,
      String(nik).trim(),
      String(nama || '-').trim(),
      String(kategoriPenerima).trim(),
      String(namaPenerima || '-').trim(),
      galonNum,
      String(metodePembayaran || 'Cash').trim(),
      String(catatan || '-').trim()
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'operasional-air-minum'!A:J",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: { values: [newRow] }
    });

    return res.json({
      success: true,
      message: `Data pengantaran (${galonNum} Galon ke ${kategoriPenerima} ${namaPenerima ? '- ' + namaPenerima : ''}) berhasil dicatat!`,
      data: {
        tanggal: jkt.dateStr,
        waktu: jkt.timeStr,
        kategoriPenerima,
        namaPenerima,
        jumlahGalon: galonNum,
        metodePembayaran
      }
    });
  } catch (error) {
    console.error('Error /api/operasional/air-minum:', error);
    return res.status(500).json({ success: false, message: 'Gagal mencatat data pengantaran galon: ' + error.message });
  }
});

/**
 * GET /api/operasional/air-minum
 * Mengambil riwayat pengantaran galon air minum
 */
app.get('/api/operasional/air-minum', async (req, res) => {
  try {
    const { role, nik, startDate, endDate, todayOnly } = req.query;
    const sheets = await getSheetsClient();
    await ensureSheetExists(sheets, 'operasional-air-minum', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Penerima', 'Nama Tenant / Stakeholder', 'Jumlah Galon', 'Metode Pembayaran', 'Catatan']);

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'operasional-air-minum'!A2:J"
    });

    const rows = response.data.values || [];
    let list = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      timestamp: r[0] || '',
      tanggal: r[1] || '',
      waktu: r[2] || '',
      nik: r[3] || '',
      nama: r[4] || '',
      kategoriPenerima: r[5] || '',
      namaPenerima: r[6] || '-',
      jumlahGalon: parseInt(r[7], 10) || 0,
      metodePembayaran: r[8] || 'Cash',
      catatan: r[9] || '-'
    }));

    if (isRoleAirMinumSender(role)) {
      const cleanNik = String(nik || '').trim().toLowerCase();
      list = list.filter(item => (item.nik || '').toLowerCase() === cleanNik);
    }

    const jkt = getJakartaDateTime();
    if (todayOnly === 'true') {
      list = list.filter(item => String(item.tanggal).trim() === jkt.dateStr);
    } else {
      if (startDate) list = list.filter(item => String(item.tanggal).trim() >= startDate);
      if (endDate) list = list.filter(item => String(item.tanggal).trim() <= endDate);
    }

    list.reverse();
    return res.json({ success: true, data: list });
  } catch (error) {
    console.error('Error GET /api/operasional/air-minum:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/operasional/summary
 * Mengambil ringkasan analitik operasional terpadu (Executive Car & Air Minum) untuk Chart & Dashboard
 * Khusus: Admin, Ketua Koperasi, dan Pengurus
 */
app.get('/api/operasional/summary', async (req, res) => {
  try {
    const { role, period } = req.query; // period: 'today' | '7days' | 'month' | 'all'
    if (role && !isRoleAdmin(role) && !isRoleKetua(role) && !isRolePengurus(role)) {
      return res.status(403).json({ success: false, message: 'Akses dashboard hanya untuk Admin, Ketua Koperasi, dan Pengurus.' });
    }

    const sheets = await getSheetsClient();
    await ensureSheetExists(sheets, 'operasional-executive-car', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Pengguna', 'Detail Pengguna', 'Jumlah Pax', 'Catatan']);
    await ensureSheetExists(sheets, 'operasional-air-minum', ['Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Penerima', 'Nama Tenant / Stakeholder', 'Jumlah Galon', 'Metode Pembayaran', 'Catatan']);

    const [carResp, airResp] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "'operasional-executive-car'!A2:I" }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: "'operasional-air-minum'!A2:J" })
    ]);

    const carRows = carResp.data.values || [];
    const airRows = airResp.data.values || [];

    const jkt = getJakartaDateTime();
    const todayStr = jkt.dateStr;

    // Filter tanggal berdasarkan period
    const nowMs = Date.now();
    const filterByPeriod = (tglStr) => {
      if (!tglStr) return false;
      if (period === 'today') return tglStr === todayStr;
      if (period === '7days') {
        const d = new Date(tglStr);
        return (nowMs - d.getTime()) <= 7 * 24 * 60 * 60 * 1000;
      }
      if (period === 'month') {
        return tglStr.startsWith(todayStr.slice(0, 7)); // YYYY-MM
      }
      return true; // 'all'
    };

    const filteredCar = carRows.filter(r => filterByPeriod(String(r[1] || '').trim()));
    const filteredAir = airRows.filter(r => filterByPeriod(String(r[1] || '').trim()));

    // 1. Executive Car Analytics
    let totalCarPax = 0;
    const carKategoriMap = { 'UMUM': 0, 'DPR-RI': 0, 'FOC': 0, 'FS-PICA': 0, 'PRIVATE JET': 0, 'Lain-lain': 0 };
    const carTrendMap = {};

    filteredCar.forEach(r => {
      const tgl = String(r[1] || '').trim();
      const kat = String(r[5] || 'Lain-lain').trim();
      const pax = parseInt(r[7], 10) || 0;
      totalCarPax += pax;

      if (carKategoriMap.hasOwnProperty(kat)) {
        carKategoriMap[kat] += pax;
      } else {
        carKategoriMap['Lain-lain'] += pax;
      }

      carTrendMap[tgl] = (carTrendMap[tgl] || 0) + pax;
    });

    // 2. Air Minum Analytics
    let totalAirGalon = 0;
    const airKategoriMap = { 'ATS': 0, 'API': 0, 'Tenant': 0, 'Stakeholder': 0 };
    const airBayarMap = { 'Cash': 0, 'QRIS': 0, 'Pembayaran di Akhir Bulan': 0 };
    const airTrendMap = {};

    filteredAir.forEach(r => {
      const tgl = String(r[1] || '').trim();
      const kat = String(r[5] || 'Tenant').trim();
      const galon = parseInt(r[7], 10) || 0;
      const bayar = String(r[8] || 'Cash').trim();
      totalAirGalon += galon;

      if (airKategoriMap.hasOwnProperty(kat)) {
        airKategoriMap[kat] += galon;
      }

      if (airBayarMap.hasOwnProperty(bayar)) {
        airBayarMap[bayar] += galon;
      } else {
        airBayarMap['Cash'] += galon;
      }

      airTrendMap[tgl] = (airTrendMap[tgl] || 0) + galon;
    });

    // Urutan Tanggal Trend untuk Chart
    const allDates = Array.from(new Set([...Object.keys(carTrendMap), ...Object.keys(airTrendMap)])).sort();
    const trendDates = allDates.slice(-14); // Ambil maks 14 hari terakhir
    const trendCarPax = trendDates.map(d => carTrendMap[d] || 0);
    const trendAirGalon = trendDates.map(d => airTrendMap[d] || 0);

    return res.json({
      success: true,
      period: period || 'all',
      todayDate: todayStr,
      summary: {
        executiveCar: {
          totalTrips: filteredCar.length,
          totalPax: totalCarPax,
          byKategori: carKategoriMap
        },
        airMinum: {
          totalDeliveries: filteredAir.length,
          totalGalon: totalAirGalon,
          byKategori: airKategoriMap,
          byPembayaran: airBayarMap
        },
        chartTrends: {
          labels: trendDates,
          carPax: trendCarPax,
          airGalon: trendAirGalon
        }
      }
    });
  } catch (error) {
    console.error('Error /api/operasional/summary:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// Fallback route to serve frontend index.html
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Export Express app for Vercel Serverless Functions
module.exports = app;

// Start local server if not running on Vercel
if (require.main === module && process.env.VERCEL !== '1') {
  app.listen(PORT, async () => {
    console.log('===================================================');
    console.log(' Portal Koperasi Server running on port ' + PORT);
    console.log(' URL: http://localhost:' + PORT);
    console.log(' Spreadsheet ID: ' + (SPREADSHEET_ID ? SPREADSHEET_ID : 'BELUM DI-SET'));
    console.log(' Credentials: ' + (process.env.GOOGLE_CREDENTIALS_JSON ? '[Dari Environment Variable Vercel]' : CREDENTIALS_PATH));
    console.log('===================================================');

    // Inisialisasi sheet otomatis jika belum ada
    try {
      await ensureAllSheetsExist();
    } catch (e) {
      console.error('[Startup Sheet Setup Error]:', e.message);
    }
  });
}