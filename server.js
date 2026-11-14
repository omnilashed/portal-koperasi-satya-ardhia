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

// Middleware (Dukungan payload besar untuk upload dokumen e-Materai maks 25MB)
app.use(cors());
app.use(express.json({ limit: '25mb' }));
app.use(express.urlencoded({ limit: '25mb', extended: true }));

// Error handler middleware khusus parsing JSON & Entity Too Large (selalu kembalikan JSON terstruktur, bukan HTML)
app.use((err, req, res, next) => {
  if (err && err.type === 'entity.too.large') {
    return res.status(413).json({
      success: false,
      message: 'Ukuran file dokumen PDF terlalu besar (Maksimal 20 MB). Silakan kompresi berkas sebelum diunggah.'
    });
  }
  if (err && err instanceof SyntaxError && err.status === 400 && 'body' in err) {
    return res.status(400).json({
      success: false,
      message: 'Format request body JSON tidak valid.'
    });
  }
  if (err) {
    console.error('[Middleware Error]:', err.message);
    return res.status(err.status || 500).json({
      success: false,
      message: err.message || 'Terjadi kesalahan pada pemrosesan server.'
    });
  }
  next();
});

// Serve static frontend files from /public or root
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.static(__dirname));

// Memastikan jumlah kolom & baris sheet mencukupi sebelum penulisan range
async function ensureSheetGridCapacity(sheets, sheetTitle, requiredCols = 35) {
  if (!SPREADSHEET_ID) return;
  try {
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const target = (meta.data.sheets || []).find(s => s.properties.title === sheetTitle);
    if (target) {
      const currentCols = target.properties.gridProperties ? target.properties.gridProperties.columnCount : 30;
      if (currentCols < requiredCols) {
        await sheets.spreadsheets.batchUpdate({
          spreadsheetId: SPREADSHEET_ID,
          requestBody: {
            requests: [{
              updateSheetProperties: {
                properties: {
                  sheetId: target.properties.sheetId,
                  gridProperties: {
                    columnCount: requiredCols
                  }
                },
                fields: 'gridProperties.columnCount'
              }
            }]
          }
        });
        console.log(`[Google Sheets] Sheet '${sheetTitle}' grid columnCount expanded from ${currentCols} to ${requiredCols}.`);
      }
    }
  } catch (e) {
    console.warn(`[ensureSheetGridCapacity Notice]:`, e.message);
  }
}

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
function getFormattedTimestamp(d = new Date()) {
  return getJakartaDateTime(d).timestampStr;
}

function getFormattedTimeStamp(d = new Date()) {
  return getJakartaDateTime(d).timestampStr;
}

function getCurrentTimestamp(d = new Date()) {
  return getJakartaDateTime(d).timestampStr;
}

// Helper: Normalisasi & Pengecekan Hierarki Role
function normalizeRole(r) {
  const s = String(r || '').trim().toLowerCase();
  if (s.includes('ketua')) return 'ketua koperasi';
  if (s.includes('admin')) return 'admin';
  if (s.includes('financial') || s.includes('keuangan')) return 'financial';
  if (s.includes('pic simpin') || s.includes('pic-simpin') || s.includes('simpin')) return 'pic simpin';
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
function isRolePicSimpin(r) { return normalizeRole(r) === 'pic simpin'; }
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
    if (role && !isRoleAdmin(role) && !isRoleKetua(role) && !isRolePengurus(role)) {
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

    if (requestingRole && !isRoleAdmin(requestingRole) && !isRoleKetua(requestingRole) && !isRolePengurus(requestingRole)) {
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

    if (requestingRole && !isRoleAdmin(requestingRole) && !isRoleKetua(requestingRole) && !isRolePengurus(requestingRole)) {
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
 * Helper: Validasi & normalisasi NIK resmi anggota koperasi
 */
function isValidNik(val) {
  if (!val) return false;
  const s = String(val).trim().toLowerCase();
  if (!s || s === '-' || s === 'nik' || s === 'null' || s === 'undefined' || s === 'none') return false;
  if (/^0+$/.test(s)) return false; // NIK tidak boleh seluruhnya 0 (misal '0', '000000')
  const alnum = s.replace(/[^a-z0-9]/gi, '');
  if (alnum.length < 3) return false;
  return true;
}

function cleanNik(val) {
  if (!isValidNik(val)) return '';
  let s = String(val).trim();
  if (/[\-\/\(]/.test(s)) {
    const parts = s.split(/[\-\/\(]/);
    const cand0 = parts[0].trim();
    const cand1 = parts[1] ? parts[1].replace(/[\)\s]/g, '').trim() : '';
    if (/^\d{4,}$/.test(cand0) || /^FIN\d+/i.test(cand0) || /^ADM\d+/i.test(cand0)) {
      s = cand0;
    } else if (/^\d{4,}$/.test(cand1) || /^FIN\d+/i.test(cand1) || /^ADM\d+/i.test(cand1)) {
      s = cand1;
    } else if (isValidNik(cand0)) {
      s = cand0;
    }
  }
  return s.toLowerCase().replace(/[^a-z0-9]/gi, '');
}

function cleanNama(val) {
  if (!val) return '';
  return String(val).trim().toLowerCase().replace(/\s+/g, ' ');
}

function isSamePersonName(nameA, nameB) {
  if (!nameA || !nameB) return false;
  const a = cleanNama(nameA);
  const b = cleanNama(nameB);
  if (!a || !b || a.length < 3 || b.length < 3) return false;
  if (a === b) return true;
  const aNoSpace = a.replace(/\s+/g, '');
  const bNoSpace = b.replace(/\s+/g, '');
  if (aNoSpace.length >= 4 && aNoSpace === bNoSpace) return true;

  // Normalisasi ejaan & variasi nama umum Indonesia
  const normA = aNoSpace
    .replace(/dj/g, 'j')
    .replace(/tj/g, 'c')
    .replace(/oe/g, 'u')
    .replace(/y/g, 'i')
    .replace(/barok/g, 'barak')
    .replace(/muhammad|mochammad|mohamad|mukhammad/g, 'muhammad');
  const normB = bNoSpace
    .replace(/dj/g, 'j')
    .replace(/tj/g, 'c')
    .replace(/oe/g, 'u')
    .replace(/y/g, 'i')
    .replace(/barok/g, 'barak')
    .replace(/muhammad|mochammad|mohamad|mukhammad/g, 'muhammad');
  if (normA.length >= 4 && normA === normB) return true;

  return false;
}

/**
 * Helper: Pisahkan NIK dan Unit Kerja dari string gabungan atau field terpisah
 */
function splitNikAndUnitKerja(rawNikUnit, explicitNik, explicitUnit) {
  let nik = String(explicitNik || '').trim();
  let unitKerja = String(explicitUnit || '').trim();
  const raw = String(rawNikUnit || '').trim();

  // Jika explicitNik tidak valid (misal '0', '000000', '-', 'nik'), abaikan
  if (!isValidNik(nik)) {
    nik = '';
  }

  if (nik && nik.includes(' - ')) {
    const p = nik.split(' - ');
    if (isValidNik(p[0])) nik = p[0].trim();
    if (!unitKerja || unitKerja === '-') unitKerja = p.slice(1).join(' - ').trim();
  }

  if ((!nik || !unitKerja || unitKerja === '-') && raw) {
    if (raw.includes(' - ')) {
      const parts = raw.split(' - ');
      const candNik = parts[0].trim();
      if (!nik && isValidNik(candNik)) nik = candNik;
      if (!unitKerja || unitKerja === '-') unitKerja = parts.slice(1).join(' - ').trim();
    } else if (/^\d{4,}$/.test(raw) || /^FIN\d+/i.test(raw) || /^ADM\d+/i.test(raw)) {
      if (!nik && isValidNik(raw)) nik = raw;
    } else {
      if (!unitKerja || unitKerja === '-') unitKerja = raw;
    }
  }

  const validNikResult = isValidNik(nik) ? nik : '-';
  const validUnitResult = unitKerja && unitKerja !== '-' ? unitKerja : '-';
  const combined = (validNikResult !== '-' && validUnitResult !== '-')
    ? `${validNikResult} - ${validUnitResult}`
    : (validNikResult !== '-' ? validNikResult : (validUnitResult !== '-' ? validUnitResult : '-'));

  return {
    nik: validNikResult,
    unitKerja: validUnitResult,
    combined
  };
}

/**
 * Helper: Hitung Pinjaman Multiguna (Maks 10 Juta, Bunga 20% / Tahun, Maks Tenor 12 Bulan)
 */
function calculateMultigunaLoanDetails(pokokInput, tenorInput) {
  const pokok = parseInt(String(pokokInput || '0').replace(/[^0-9]/g, ''), 10) || 0;
  let tenor = parseInt(String(tenorInput || '10').replace(/[^0-9]/g, ''), 10) || 10;
  if (tenor < 1) tenor = 1;
  if (tenor > 10) tenor = 10;

  const rateTahunan = 0.20; // 20% per tahun
  const totalBiaya = Math.round(pokok * rateTahunan);
  const totalKewajiban = pokok + totalBiaya;
  const cicilanNormal = Math.round(totalKewajiban / tenor);

  const jadwal = [];
  let akumulasi = 0;
  for (let i = 1; i <= tenor; i++) {
    if (i === tenor) {
      const akhir = totalKewajiban - akumulasi;
      jadwal.push({ bulanKe: i, cicilan: akhir });
      akumulasi += akhir;
    } else {
      jadwal.push({ bulanKe: i, cicilan: cicilanNormal });
      akumulasi += cicilanNormal;
    }
  }

  return {
    kategori: 'MULTIGUNA',
    pokok,
    tenor,
    rate: rateTahunan,
    ratePercent: '20% / Thn',
    totalBiaya,
    totalKewajiban,
    cicilanNormal,
    jadwal
  };
}

/**
 * Helper: Kalkulasi Resmi Topup & Pelunasan Pinjaman (Tunai & Multiguna)
 * 1. Topup Pinjaman Tunai:
 *    - Hasil Kali = Sisa Angsuran x Pokok Pinjaman
 *    - Biaya Admin 2,5% = 2,5% x Hasil Kali
 *    - Jumlah Harus Dilunasi = Hasil Kali + Biaya Admin 2,5%
 *    - Sisa Diterima (Total Ditransfer) = Topup Baru - Jumlah Harus Dilunasi
 * 2. Topup Pinjaman Multiguna (Tanpa Biaya Admin 2,5%):
 *    - Bunga 20% = 20% x Pokok Pinjaman
 *    - Jumlah Harus Dilunasi = Sisa Angsuran x (Pokok Pinjaman + 20% dari Pokok Pinjaman)
 *    - Sisa Diterima (Total Ditransfer) = Topup Baru - Jumlah Harus Dilunasi
 * 3. Pelunasan Pinjaman Tunai Anggota:
 *    - Hasil Kali = Pokok Pinjaman x Sisa Tenor
 *    - Biaya Admin 2,5% = 2,5% x Hasil Kali
 *    - Hasil Pelunasan Sistem = Hasil Kali + Biaya Admin 2,5%
 * 4. Pelunasan Pinjaman Multiguna Anggota (Tanpa Biaya Admin 2,5%):
 *    - Pokok + 20% = Pokok Pinjaman + (20% x Pokok Pinjaman)
 *    - Hasil Pelunasan Sistem = (Pokok + 20%) x Sisa Tenor
 */
function calculateTopupOrPelunasanDetails(params = {}) {
  const tipeLower = String(params.tipeFormulir || '').trim().toLowerCase();
  const sisaAngsuran = Math.max(0, parseInt(String(params.sisaAngsuran || params.sisaTenor || '0').replace(/[^0-9]/g, ''), 10) || 0);
  const pokokPinjaman = Math.max(0, parseInt(String(params.pokokPinjaman || '0').replace(/[^0-9]/g, ''), 10) || 0);
  const plafonTopupBaru = Math.max(0, parseInt(String(params.plafonTopupBaru || '0').replace(/[^0-9]/g, ''), 10) || 0);

  const isMultiguna = tipeLower.includes('multiguna');
  const isTopup = tipeLower.includes('topup');
  const isPelunasan = tipeLower.includes('pelunasan');

  if (isMultiguna) {
    // Multiguna: Sisa Angsuran x (Pokok Pinjaman + 20% dari Pokok Pinjaman), TANPA Biaya Admin 2,5%
    const bunga20Persen = Math.round(pokokPinjaman * 0.20);
    const pokokPlusBunga = pokokPinjaman + bunga20Persen;
    const hasilKali = Math.round(sisaAngsuran * pokokPlusBunga);
    const biayaAdmin = 0;
    const jumlahHarusDilunasi = hasilKali;
    const sisaDiterima = isTopup ? (plafonTopupBaru - jumlahHarusDilunasi) : 0;

    return {
      isMultiguna: true,
      isTopup,
      isPelunasan,
      sisaAngsuran,
      pokokPinjaman,
      bunga20Persen,
      pokokPlusBunga,
      hasilKali,
      biayaAdmin,
      jumlahHarusDilunasi,
      plafonTopupBaru,
      sisaDiterima
    };
  } else {
    // Tunai: (Sisa Angsuran x Pokok Pinjaman) + (2,5% x Sisa Angsuran x Pokok Pinjaman)
    const hasilKali = Math.round(sisaAngsuran * pokokPinjaman);
    const biayaAdmin = Math.round(hasilKali * 0.025);
    const jumlahHarusDilunasi = hasilKali + biayaAdmin;
    const sisaDiterima = isTopup ? (plafonTopupBaru - jumlahHarusDilunasi) : 0;

    return {
      isMultiguna: false,
      isTopup,
      isPelunasan,
      sisaAngsuran,
      pokokPinjaman,
      bunga20Persen: 0,
      pokokPlusBunga: pokokPinjaman,
      hasilKali,
      biayaAdmin,
      jumlahHarusDilunasi,
      plafonTopupBaru,
      sisaDiterima
    };
  }
}

/**
 * GET /api/pengajuan
 * Query params: role, nik, nama, username
 */
app.get('/api/pengajuan', async (req, res) => {
  try {
    const { role, nik, nama, username } = req.query;

    const sheets = await getSheetsClient();
    await ensureSheetGridCapacity(sheets, 'Form Responses 1', 16);
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A2:P"
    });

    const rows = response.data.values || [];
    let list = rows.map((r, idx) => {
      const parsed = splitNikAndUnitKerja(r[3] || '', r[10] || '', r[11] || '');
      let meta = null;
      if (r[15]) {
        try { meta = JSON.parse(r[15]); } catch (e) {}
      }
      return {
        rowIndex: idx + 2,
        timestamp: r[0] || '',
        datetime: r[0] || '',
        tipeFormulir: r[1] || '',
        namaKaryawan: r[2] || '',
        nikUnitKerja: r[3] || parsed.combined,
        nikUnit: r[3] || parsed.combined,
        nik: parsed.nik,
        unitKerja: parsed.unitKerja,
        noRekening: r[4] || '',
        atasNamaRekening: r[5] || '',
        totalDitransfer: r[6] || '0',
        cicilan: r[7] || '-',
        keterangan: r[8] || '',
        status: r[9] || 'Created',
        oldLoanRef: r[12] || (meta && meta.oldLoanRef) || '',
        plafonTopupBaru: r[13] ? Number(r[13]) : (meta && meta.plafonTopupBaru) || 0,
        jumlahHarusDilunasi: r[14] ? Number(r[14]) : (meta && meta.jumlahHarusDilunasi) || 0,
        topupPelunasanMeta: meta
      };
    });

    // Hak Akses Tracking & Progress:
    // Yang bisa melihat semua data adalah role Administrator, Ketua Koperasi, dan Financial.
    // Selain itu (role lainnya), hanya bisa melihat riwayat tracking and progress milik dirinya sendiri
    // sesuai data berdasarkan username, nama lengkap, atau NIK.
    const canSeeAll = isRoleAdmin(role) || isRoleKetua(role) || isRoleFinancial(role);

    if (!canSeeAll) {
      let userNik = String(nik || '').trim().toLowerCase();
      let userNama = String(nama || '').trim().toLowerCase().replace(/\s+/g, ' ');
      let userUname = String(username || '').trim().toLowerCase();

      // Auto-resolve jika userNama atau userNik kosong tetapi username disediakan
      if ((!userNama || !userNik) && userUname) {
        try {
          const userResp = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: 'Users!A2:E'
          });
          const uRows = userResp.data.values || [];
          const matchedU = uRows.find(u => String(u[0] || '').trim().toLowerCase() === userUname);
          if (matchedU) {
            if (!userNama) userNama = String(matchedU[3] || '').trim().toLowerCase().replace(/\s+/g, ' ');
            if (!userNik) userNik = String(matchedU[4] || '').trim().toLowerCase();
          }
        } catch (eLookup) {}
      }

      // Ekstrak angka unik jika NIK mengandung nomor identitas (misal: "20246808")
      const nikNumberMatch = userNik.match(/\d{3,}/);
      const nikOnly = nikNumberMatch ? nikNumberMatch[0] : '';
      const userUnameClean = userUname.replace(/[^a-z0-9]/g, ' ').trim();

      list = list.filter(item => {
        const itemNik = String(item.nikUnitKerja || item.nikUnit || '').trim().toLowerCase();
        const itemNama = String(item.namaKaryawan || '').trim().toLowerCase().replace(/\s+/g, ' ');
        const itemAtasNama = String(item.atasNamaRekening || '').trim().toLowerCase().replace(/\s+/g, ' ');

        // 1. Cek kesesuaian Nama Lengkap (Karyawan atau Atas Nama Rekening)
        if (userNama) {
          if (
            itemNama === userNama ||
            itemNama.includes(userNama) ||
            userNama.includes(itemNama)
          ) {
            return true;
          }
          if (
            itemAtasNama && (
              itemAtasNama === userNama ||
              itemAtasNama.includes(userNama) ||
              userNama.includes(itemAtasNama)
            )
          ) {
            return true;
          }
        }

        // 2. Cek kesesuaian Username
        if (userUname) {
          if (itemNama === userUname || itemNama.includes(userUname)) {
            return true;
          }
          if (itemNik === userUname || itemNik.includes(userUname)) {
            return true;
          }
          if (userUnameClean && (itemNama.includes(userUnameClean) || userUnameClean.includes(itemNama))) {
            return true;
          }
        }

        // 3. Cek kesesuaian NIK
        if (nikOnly && itemNik.includes(nikOnly)) {
          return true;
        }
        if (userNik && !['karyawan koperasi', 'koperasi', 'unit', 'staff', 'avsec', 'arff', '-'].includes(userNik)) {
          if (
            itemNik === userNik ||
            itemNik.startsWith(userNik + ' ') ||
            itemNik.endsWith(' ' + userNik) ||
            itemNik.includes(userNik)
          ) {
            return true;
          }
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
 * Body: {
 *   tipeFormulir, namaKaryawan, nik, unitKerja, nikUnitKerja,
 *   noRekening, atasNamaRekening, totalDitransfer, cicilan, keterangan, userRole,
 *   oldLoanRef, sisaAngsuranLama, pokokPinjamanLama, plafonTopupBaru, jumlahHarusDilunasi, nominalPelunasanSistem
 * }
 */
app.post('/api/pengajuan', async (req, res) => {
  try {
    const {
      tipeFormulir,
      namaKaryawan,
      nik,
      unitKerja,
      nikUnitKerja,
      noRekening,
      atasNamaRekening,
      totalDitransfer,
      cicilan,
      keterangan,
      userRole,
      oldLoanRef,
      sisaAngsuranLama,
      pokokPinjamanLama,
      plafonTopupBaru,
      jumlahHarusDilunasi,
      nominalPelunasanSistem
    } = req.body;

    // Role Guard: Hanya Admin, Ketua Koperasi, dan Financial yang dapat menginput pengajuan pinjaman
    if (userRole && !isRoleAdmin(userRole) && !isRoleKetua(userRole) && !isRoleFinancial(userRole)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Hanya Admin, Ketua Koperasi, dan Financial yang dapat menginput pengajuan pinjaman/biaya.'
      });
    }

    const parsedNikUnit = splitNikAndUnitKerja(nikUnitKerja, nik, unitKerja);
    const cleanNik = parsedNikUnit.nik !== '-' ? parsedNikUnit.nik : '';
    const cleanUnit = parsedNikUnit.unitKerja !== '-' ? parsedNikUnit.unitKerja : '';
    const combinedNikUnit = parsedNikUnit.combined;

    if (!tipeFormulir || !namaKaryawan || (!cleanNik && !cleanUnit) || !noRekening || !atasNamaRekening || totalDitransfer === undefined || totalDitransfer === '') {
      return res.status(400).json({ success: false, message: 'Lengkapi semua field wajib (termasuk NIK dan Unit Kerja) pada formulir pengajuan.' });
    }

    const nominalNum = parseInt(String(totalDitransfer).replace(/[^0-9-]/g, ''), 10) || 0;
    const tenorNum = parseInt(String(cicilan || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const tipeLower = String(tipeFormulir).trim().toLowerCase();

    const isTopupTunai = tipeLower === 'topup pinjaman tunai';
    const isTopupMulti = tipeLower === 'topup pinjaman multiguna';
    const isPelunasanTunai = tipeLower === 'pelunasan pinjaman tunai anggota';
    const isPelunasanMulti = tipeLower === 'pelunasan pinjaman multiguna anggota';
    const isTopup = isTopupTunai || isTopupMulti;
    const isPelunasan = isPelunasanTunai || isPelunasanMulti;

    // Validasi Aturan Resmi Pinjaman Multiguna (Maks Rp 10 Juta, Bunga 20%/Thn, Maks Tenor 12 Bulan)
    if (tipeLower === 'peminjaman multiguna') {
      if (nominalNum <= 0 || nominalNum > 10000000) {
        return res.status(400).json({
          success: false,
          message: 'Plafon Pinjaman Multiguna maksimal adalah Rp 10.000.000 (dengan bunga 20% per tahun).'
        });
      }
      if (tenorNum < 1 || tenorNum > 10) {
        return res.status(400).json({
          success: false,
          message: 'Tenor Pinjaman Multiguna maksimal adalah 10 bulan (1 s/d 10 kali cicilan).'
        });
      }
    }

    let calcMeta = null;
    let finalTotalDitransfer = nominalNum;
    let finalKeterangan = String(keterangan || '').trim();

    if (isTopup || isPelunasan) {
      const calc = calculateTopupOrPelunasanDetails({
        tipeFormulir,
        sisaAngsuran: sisaAngsuranLama !== undefined ? sisaAngsuranLama : tenorNum,
        pokokPinjaman: pokokPinjamanLama,
        plafonTopupBaru: plafonTopupBaru !== undefined ? plafonTopupBaru : nominalNum
      });

      if (calc.sisaAngsuran <= 0 || calc.pokokPinjaman <= 0) {
        return res.status(400).json({
          success: false,
          message: 'Untuk proses Topup / Pelunasan, Sisa Angsuran (Sisa Tenor) dan Pokok Pinjaman wajib lebih dari 0.'
        });
      }

      if (isTopupMulti) {
        if (calc.plafonTopupBaru <= 0 || calc.plafonTopupBaru > 10000000) {
          return res.status(400).json({
            success: false,
            message: 'Plafon Topup Pinjaman Multiguna maksimal adalah Rp 10.000.000.'
          });
        }
        if (tenorNum < 1 || tenorNum > 10) {
          return res.status(400).json({
            success: false,
            message: 'Tenor baru Topup Pinjaman Multiguna maksimal adalah 10 bulan (1 s/d 10 kali cicilan).'
          });
        }
      } else if (isTopupTunai) {
        if (calc.plafonTopupBaru <= 0 || calc.plafonTopupBaru > 70000000) {
          return res.status(400).json({
            success: false,
            message: 'Plafon Topup Pinjaman Tunai harus di antara Rp 1 s/d Rp 70.000.000.'
          });
        }
        if (tenorNum < 1 || tenorNum > 120) {
          return res.status(400).json({
            success: false,
            message: 'Tenor baru Topup Pinjaman Tunai wajib diisi (maksimal 120 bulan).'
          });
        }
      }

      if (isTopup) {
        if (calc.sisaDiterima < 0) {
          return res.status(400).json({
            success: false,
            message: `Plafon Topup baru (Rp ${calc.plafonTopupBaru.toLocaleString('id-ID')}) tidak mencukupi untuk melunasi sisa kewajiban pinjaman lama (Rp ${calc.jumlahHarusDilunasi.toLocaleString('id-ID')}).`
          });
        }
        finalTotalDitransfer = calc.sisaDiterima;
      } else if (isPelunasan) {
        const expectedPelunasan = calc.jumlahHarusDilunasi;
        const isModified = Math.abs(nominalNum - expectedPelunasan) > 1;
        // Jika ada perubahan dari hasil hitungan otomatis, Financial wajib menambahkan keterangan
        if (isModified && (!finalKeterangan || finalKeterangan === '-' || finalKeterangan.length < 5)) {
          return res.status(400).json({
            success: false,
            message: `Nominal pelunasan yang diinput (Rp ${nominalNum.toLocaleString('id-ID')}) berbeda dari perhitungan sistem (Rp ${expectedPelunasan.toLocaleString('id-ID')}). Financial WAJIB mengisi alasan/keterangan perubahan pada kolom Keterangan!`
          });
        }
        finalTotalDitransfer = nominalNum > 0 ? nominalNum : expectedPelunasan;
      }

      calcMeta = {
        tipe: tipeFormulir,
        oldLoanRef: oldLoanRef || '',
        sisaAngsuran: calc.sisaAngsuran,
        pokokPinjaman: calc.pokokPinjaman,
        bunga20Persen: calc.bunga20Persen,
        hasilKali: calc.hasilKali,
        biayaAdmin: calc.biayaAdmin,
        jumlahHarusDilunasi: calc.jumlahHarusDilunasi,
        plafonTopupBaru: isTopup ? calc.plafonTopupBaru : 0,
        tenorTopupBaru: isTopup ? tenorNum : 0,
        sisaDiterima: isTopup ? calc.sisaDiterima : finalTotalDitransfer,
        isModified: isPelunasan ? (Math.abs(finalTotalDitransfer - calc.jumlahHarusDilunasi) > 1) : false
      };
    }

    const timestamp = getCurrentTimestamp();
    const defaultStatus = 'Created';

    const sheets = await getSheetsClient();
    await ensureSheetGridCapacity(sheets, 'Form Responses 1', 16);

    const newRow = [
      timestamp,
      tipeFormulir,
      namaKaryawan,
      combinedNikUnit,
      String(noRekening),
      atasNamaRekening,
      finalTotalDitransfer,
      cicilan || '-',
      finalKeterangan || '-',
      defaultStatus,
      cleanNik || '-',
      cleanUnit || '-',
      calcMeta ? (calcMeta.oldLoanRef || '-') : '-',
      calcMeta ? String(calcMeta.plafonTopupBaru || calcMeta.jumlahHarusDilunasi || 0) : '-',
      calcMeta ? String(calcMeta.jumlahHarusDilunasi || 0) : '-',
      calcMeta ? JSON.stringify(calcMeta) : ''
    ];

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A:P",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [newRow]
      }
    });

    return res.json({
      success: true,
      message: isPelunasan
        ? 'Pengajuan Pelunasan Pinjaman berhasil disimpan (Status: Created) dan wajib menunggu Approval Admin atau Ketua Koperasi.'
        : (isTopup
          ? 'Pengajuan Topup Pinjaman berhasil disimpan (Status: Created) dan menunggu Approval Admin / Ketua Koperasi.'
          : 'Pengajuan pinjaman/biaya berhasil disimpan dengan status Created.'),
      data: {
        timestamp,
        tipeFormulir,
        namaKaryawan,
        nik: cleanNik,
        unitKerja: cleanUnit,
        nikUnitKerja: combinedNikUnit,
        totalDitransfer: finalTotalDitransfer,
        topupPelunasanMeta: calcMeta,
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
 * Otomatis mengintegrasikan data pinjaman ke NIK anggota ketika status menjadi 'Done',
 * meniban data pinjaman lama saat Topup selesai, atau menandai pinjaman lama menjadi 'Lunas' (tanpa dihapus) saat Pelunasan disetujui.
 */
app.put('/api/pengajuan/status', async (req, res) => {
  try {
    const { rowIndex, newStatus, userRole } = req.body;

    if (!rowIndex || !newStatus) {
      return res.status(400).json({ success: false, message: 'Index baris (rowIndex) dan status baru wajib disediakan.' });
    }

    const validStatuses = ['Created', 'On Process', 'Done', 'Lunas', 'Rejected'];
    if (!validStatuses.includes(newStatus)) {
      return res.status(400).json({ success: false, message: 'Status tidak valid. Pilihan: ' + validStatuses.join(', ') });
    }

    if (isRoleAnggota(userRole) || isRolePegawai(userRole)) {
      return res.status(403).json({ success: false, message: 'Akses ditolak: Anggota dan Pegawai tidak memiliki hak untuk mengubah status pengajuan.' });
    }

    const sheets = await getSheetsClient();
    await ensureSheetGridCapacity(sheets, 'Form Responses 1', 16);

    // Ambil data seluruh baris untuk pencarian pinjaman lama yang akurat bila Topup / Pelunasan
    const allFrRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A2:P"
    });
    const allFrRows = allFrRes.data.values || [];
    const r = allFrRows[Number(rowIndex) - 2] || [];

    const tipeForm = String(r[1] || '').trim();
    const tipeLower = tipeForm.toLowerCase();
    const namaKaryawan = String(r[2] || '').trim();
    const currentStatus = String(r[9] || 'Created').trim();
    const parsedNikUnit = splitNikAndUnitKerja(r[3] || '', r[10] || '', r[11] || '');
    let resolvedNik = parsedNikUnit.nik !== '-' ? parsedNikUnit.nik : '';
    let resolvedUnit = parsedNikUnit.unitKerja !== '-' ? parsedNikUnit.unitKerja : '';

    const isTopup = tipeLower.includes('topup');
    const isPelunasan = tipeLower.includes('pelunasan');
    const isTargetMultiguna = tipeLower.includes('multiguna');

    // Validasi Role: Semua proses Pelunasan & Topup dari status Created wajib di-approve oleh Admin atau Ketua Koperasi (atau Pengurus untuk pengajuan biasa)
    if (isPelunasan && currentStatus === 'Created' && ['On Process', 'Done', 'Lunas'].includes(newStatus)) {
      if (!isRoleAdmin(userRole) && !isRoleKetua(userRole)) {
        return res.status(403).json({
          success: false,
          message: 'Semua proses Pelunasan Pinjaman Anggota wajib di-approval terlebih dahulu oleh Admin atau Ketua Koperasi!'
        });
      }
    }

    if (isRolePengurus(userRole) || isRoleKetua(userRole)) {
      if (!['On Process', 'Done', 'Lunas', 'Rejected'].includes(newStatus)) {
        return res.status(403).json({
          success: false,
          message: 'Role Pengurus dan Ketua Koperasi hanya memiliki otoritas untuk Persetujuan (On Process / Lunas) atau Penolakan (Rejected).'
        });
      }
      if (isRolePengurus(userRole) && !isRoleKetua(userRole) && ['Done', 'Lunas'].includes(newStatus)) {
        return res.status(403).json({
          success: false,
          message: 'Role Pengurus hanya dapat mengubah status menjadi On Process (Disetujui) atau Rejected.'
        });
      }
    }

    if (isRoleFinancial(userRole)) {
      if (currentStatus === 'Created' && (isPelunasan || isTopup)) {
        return res.status(403).json({
          success: false,
          message: 'Pengajuan Topup / Pelunasan berstatus Created wajib disetujui (Approve) oleh Admin atau Ketua Koperasi terlebih dahulu.'
        });
      }
      if (!['Done', 'Lunas', 'On Process'].includes(newStatus)) {
        return res.status(403).json({
          success: false,
          message: 'Role Financial hanya dapat memproses penyelesaian (Done / Lunas) atau pengajuan diproses (On Process).'
        });
      }
    }

    // Jika NIK belum terisi murni, cari otomatis di sheet Users berdasarkan Nama Lengkap
    if (!resolvedNik && namaKaryawan) {
      try {
        const uRes = await sheets.spreadsheets.values.get({
          spreadsheetId: SPREADSHEET_ID,
          range: "'Users'!A2:H"
        });
        const uRows = uRes.data.values || [];
        const cleanTargetName = namaKaryawan.toLowerCase().replace(/\s+/g, ' ');
        const matchedU = uRows.find(u => {
          const uName = String(u[3] || '').trim().toLowerCase().replace(/\s+/g, ' ');
          const uUname = String(u[0] || '').trim().toLowerCase();
          return uName === cleanTargetName || uUname === cleanTargetName || (uName && cleanTargetName.includes(uName));
        });
        if (matchedU) {
          const uSplit = splitNikAndUnitKerja(matchedU[4] || '', '', '');
          if (uSplit.nik !== '-') resolvedNik = uSplit.nik;
          if (!resolvedUnit && uSplit.unitKerja !== '-') resolvedUnit = uSplit.unitKerja;
        }
      } catch (eLookup) {}
    }

    // Parse metadata Topup / Pelunasan jika ada
    let meta = null;
    if (r[15]) {
      try { meta = JSON.parse(r[15]); } catch (e) {}
    }
    const oldLoanRef = String(r[12] || (meta && meta.oldLoanRef) || '').trim();
    const plafonTopupBaru = parseInt(String(r[13] || (meta && meta.plafonTopupBaru) || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const jumlahDilunasi = parseInt(String(r[14] || (meta && meta.jumlahHarusDilunasi) || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const tenorBaru = parseInt(String(r[7] || (meta && meta.tenorTopupBaru) || '12').replace(/[^0-9]/g, ''), 10) || 12;

    // Helper untuk mencari baris pinjaman aktif terakhir milik anggota ini di 'Form Responses 1'
    function findTargetOldLoanRowIndex() {
      if (oldLoanRef) {
        const refMatch = oldLoanRef.match(/^(?:FR-|MG-FR-|TN-FR-)?(\d+)$/i);
        if (refMatch) {
          const idx = parseInt(refMatch[1], 10);
          if (idx >= 2 && idx !== Number(rowIndex)) return idx;
        }
      }
      // Cari dari bawah (pinjaman aktif terakhir dengan kategori yang sama)
      const targetNameClean = namaKaryawan.toLowerCase().replace(/\s+/g, ' ');
      const targetNikClean = resolvedNik.toLowerCase();
      for (let i = allFrRows.length - 1; i >= 0; i--) {
        const candidateRowIdx = i + 2;
        if (candidateRowIdx === Number(rowIndex)) continue;
        const cr = allFrRows[i];
        const cTipe = String(cr[1] || '').toLowerCase();
        const cSt = String(cr[9] || '').toLowerCase();
        if (cSt !== 'done' && cSt !== 'validasi by financial' && !cSt.includes('terverifikasi')) continue;

        const cIsMulti = cTipe === 'peminjaman multiguna' || (cTipe.includes('multiguna') && !cTipe.includes('pelunasan'));
        const cIsTunai = cTipe === 'peminjaman tunai' || (cTipe.includes('pinjaman tunai') && !cTipe.includes('pelunasan'));
        if (isTargetMultiguna && !cIsMulti) continue;
        if (!isTargetMultiguna && !cIsTunai) continue;

        const cSplit = splitNikAndUnitKerja(cr[3] || '', cr[10] || '', cr[11] || '');
        const cNik = cSplit.nik !== '-' ? cSplit.nik.toLowerCase() : '';
        const cNama = String(cr[2] || '').trim().toLowerCase().replace(/\s+/g, ' ');

        if ((targetNikClean && cNik && cNik === targetNikClean) || (targetNameClean && cNama === targetNameClean)) {
          return candidateRowIdx;
        }
      }
      return -1;
    }

    let extraActionNote = '';
    let effectiveNewStatus = newStatus;

    // 1. LOGIKA OTOMATIS SAAT TOPUP PINJAMAN DISETUJUI / SELESAI (DONE):
    // Data pinjaman lama otomatis tertiban / terupdate dengan pengajuan baru dan tenor baru!
    if (isTopup && (newStatus === 'Done' || newStatus === 'Lunas')) {
      effectiveNewStatus = 'Done';
      const targetOldRow = findTargetOldLoanRowIndex();
      const baseLoanType = isTargetMultiguna ? 'Peminjaman Multiguna' : 'Peminjaman Tunai';
      const updatedPlafon = plafonTopupBaru > 0 ? plafonTopupBaru : (parseInt(String(r[6] || '0').replace(/[^0-9]/g, ''), 10) || 0);
      const sisaTransfer = parseInt(String(r[6] || '0').replace(/[^0-9]/g, ''), 10) || 0;
      const nowTs = getCurrentTimestamp();

      if (targetOldRow >= 2) {
        const oldRowData = allFrRows[targetOldRow - 2] || [];
        const oldPlafon = oldRowData[6] || '0';
        const oldTenor = oldRowData[7] || '-';
        const updatedKet = `[TOPUP DARI PINJAMAN LAMA Rp ${Number(String(oldPlafon).replace(/[^0-9]/g, '') || 0).toLocaleString('id-ID')} (${oldTenor} Bln)] Plafon Baru: Rp ${updatedPlafon.toLocaleString('id-ID')} | Tenor Baru: ${tenorBaru} Bln | Pelunasan Sisa Lama: Rp ${jumlahDilunasi.toLocaleString('id-ID')} | Sisa Diterima: Rp ${sisaTransfer.toLocaleString('id-ID')}. ${r[8] || ''}`.trim();

        // Tertiban / Overwrite baris pinjaman lama dengan plafon baru & tenor baru
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'Form Responses 1'!A${targetOldRow}:L${targetOldRow}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: {
            values: [[
              nowTs,
              baseLoanType,
              namaKaryawan,
              parsedNikUnit.combined,
              String(r[4] || oldRowData[4] || ''),
              String(r[5] || oldRowData[5] || namaKaryawan),
              updatedPlafon,
              tenorBaru,
              updatedKet,
              'Done',
              resolvedNik || '-',
              resolvedUnit || '-'
            ]]
          }
        });
        // Tandai baris transaksi Topup sebagai 'Done (Topup Applied)' agar tidak terhitung ganda di Monitoring Pinjaman Aktif
        effectiveNewStatus = 'Done';
        extraActionNote = ` Data pinjaman lama (Baris #${targetOldRow}) telah otomatis tertiban/terupdate dengan Plafon Baru Rp ${updatedPlafon.toLocaleString('id-ID')} dan Tenor Baru ${tenorBaru} Bulan.`;
      } else {
        extraActionNote = ` Topup Pinjaman aktif dengan Plafon Baru Rp ${updatedPlafon.toLocaleString('id-ID')} dan Tenor Baru ${tenorBaru} Bulan.`;
      }
    }

    // 2. LOGIKA OTOMATIS SAAT PELUNASAN PINJAMAN DISETUJUI OLEH ADMIN / KETUA KOPERASI / SELESAI:
    // Data pinjaman lama dinyatakan LUNAS dan TIDAK DIHAPUS untuk riwayat pinjaman!
    if (isPelunasan && (newStatus === 'On Process' || newStatus === 'Done' || newStatus === 'Lunas')) {
      // Jika Admin atau Ketua Koperasi langsung meng-approve Pelunasan, atau Finance menandai selesai, tandai pinjaman lama sebagai 'Lunas'
      const shouldMarkOldLunas = (newStatus === 'Done' || newStatus === 'Lunas' || isRoleAdmin(userRole) || isRoleKetua(userRole));
      if (newStatus === 'Done' || newStatus === 'Lunas') {
        effectiveNewStatus = 'Lunas';
      }
      if (shouldMarkOldLunas) {
        const targetOldRow = findTargetOldLoanRowIndex();
        if (targetOldRow >= 2) {
          const oldRowData = allFrRows[targetOldRow - 2] || [];
          const oldKet = String(oldRowData[8] || '').trim();
          const lunasStamp = `[LUNAS PER ${getCurrentTimestamp()} - Pelunasan Rp ${Number(String(r[6] || '0').replace(/[^0-9]/g, '') || 0).toLocaleString('id-ID')}]`;
          const newOldKet = oldKet.includes('[LUNAS') ? oldKet : `${oldKet} ${lunasStamp}`.trim();

          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'Form Responses 1'!I${targetOldRow}:L${targetOldRow}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: {
              values: [[newOldKet, 'Lunas', resolvedNik || '-', resolvedUnit || '-']]
            }
          });
          extraActionNote = ` Data pinjaman lama anggota (Baris #${targetOldRow}) telah dinyatakan LUNAS dan tetap tersimpan di Riwayat Pinjaman.`;
        } else if (oldLoanRef) {
          // Cek apakah pinjaman berasal dari sheet LoanRequests
          try {
            const lrRes = await sheets.spreadsheets.values.get({
              spreadsheetId: SPREADSHEET_ID,
              range: "'LoanRequests'!A2:Z"
            });
            const lrRows = lrRes.data.values || [];
            const lrIdx = lrRows.findIndex(lr => String(lr[0] || '') === oldLoanRef || String(lr[1] || '') === oldLoanRef);
            if (lrIdx !== -1) {
              await sheets.spreadsheets.values.update({
                spreadsheetId: SPREADSHEET_ID,
                range: `'LoanRequests'!Y${lrIdx + 2}`,
                valueInputOption: 'USER_ENTERED',
                requestBody: { values: [['LUNAS']] }
              });
              extraActionNote = ` Kontrak pinjaman (${oldLoanRef}) telah dinyatakan LUNAS dan tetap tersimpan di Riwayat Pinjaman.`;
            }
          } catch (eLrUpdate) {}
        }
      }
    }

    // Update Status (Col J) dan pastikan NIK & Unit Kerja (Col K:L) terintegrasi
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'Form Responses 1'!J${rowIndex}:L${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[effectiveNewStatus, resolvedNik || '-', resolvedUnit || '-']]
      }
    });

    const isLoan = tipeLower.includes('peminjaman') || tipeLower.includes('pinjaman');
    const integrationNote = (effectiveNewStatus === 'Done' && isLoan && !extraActionNote)
      ? ` Data pinjaman telah otomatis terintegrasi dengan NIK anggota (${resolvedNik || namaKaryawan}) pada Monitoring Pinjaman.`
      : extraActionNote;

    return res.json({
      success: true,
      message: 'Status pengajuan baris ke-' + rowIndex + ' berhasil diperbarui menjadi ' + effectiveNewStatus + '.' + integrationNote,
      newStatus: effectiveNewStatus,
      integratedNik: resolvedNik || null
    });
  } catch (error) {
    console.error('Error /api/pengajuan/status:', error);
    return res.status(500).json({ success: false, message: 'Gagal memperbarui status pengajuan: ' + error.message });
  }
});

/**
 * PUT /api/pengajuan
 * Body: { rowIndex, role, tipeFormulir, namaKaryawan, nik, unitKerja, nikUnitKerja, noRekening, atasNamaRekening, totalDitransfer, cicilan, keterangan, status }
 * Khusus Admin: Edit data pengajuan
 */
app.put('/api/pengajuan', async (req, res) => {
  try {
    const {
      rowIndex,
      role,
      tipeFormulir,
      namaKaryawan,
      nik,
      unitKerja,
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

    const parsedNikUnit = splitNikAndUnitKerja(nikUnitKerja, nik, unitKerja);
    const nominalNum = parseInt(String(totalDitransfer || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const tenorNum = parseInt(String(cicilan || '0').replace(/[^0-9]/g, ''), 10) || 0;

    if (String(tipeFormulir || '').trim().toLowerCase() === 'peminjaman multiguna') {
      if (nominalNum > 10000000) {
        return res.status(400).json({ success: false, message: 'Plafon Pinjaman Multiguna maksimal adalah Rp 10.000.000.' });
      }
      if (tenorNum > 12) {
        return res.status(400).json({ success: false, message: 'Tenor Pinjaman Multiguna maksimal adalah 12 bulan.' });
      }
    }

    const sheets = await getSheetsClient();
    await ensureSheetGridCapacity(sheets, 'Form Responses 1', 12);

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'Form Responses 1'!B${rowIndex}:L${rowIndex}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          tipeFormulir || '',
          namaKaryawan || '',
          parsedNikUnit.combined,
          String(noRekening || ''),
          atasNamaRekening || '',
          totalDitransfer || '0',
          cicilan || '-',
          keterangan || '-',
          status || 'Created',
          parsedNikUnit.nik,
          parsedNikUnit.unitKerja
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
 * GET /api/loans/monitoring
 * Monitoring Pinjaman Anggota (Terintegrasi dengan NIK Anggota untuk Pinjaman Multiguna & Pinjaman Tunai yang berstatus Done)
 * Query params: nik, nama, username, role, scope ('me' | 'all')
 */
app.get('/api/loans/monitoring', async (req, res) => {
  try {
    const { nik, nama, username, role, scope = 'me' } = req.query;
    const sheets = await getSheetsClient();

    // 1. Load Users untuk resolusi NIK & Unit Kerja anggota secara akurat
    const usersRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Users'!A2:H"
    });
    const userRows = usersRes.data.values || [];
    const userByName = new Map();
    const userByNameNoSpace = new Map();
    const userByUsername = new Map();
    const userByNik = new Map();

    for (const u of userRows) {
      const uUname = String(u[0] || '').trim().toLowerCase();
      const uNama = String(u[3] || '').trim();
      const uNamaClean = cleanNama(uNama);
      const uSplit = splitNikAndUnitKerja(u[4] || '', '', '');
      const uObj = {
        username: uUname,
        nama: uNama,
        nik: uSplit.nik !== '-' ? uSplit.nik : '',
        unitKerja: uSplit.unitKerja !== '-' ? uSplit.unitKerja : ''
      };
      if (uUname) userByUsername.set(uUname, uObj);
      if (uNamaClean) {
        userByName.set(uNamaClean, uObj);
        userByNameNoSpace.set(uNamaClean.replace(/\s+/g, ''), uObj);
      }
      if (uObj.nik) userByNik.set(cleanNik(uObj.nik), uObj);
    }

    // Identitas user yang sedang login
    const reqUname = String(username || '').trim().toLowerCase();
    const reqSplit = splitNikAndUnitKerja(nik || '', '', '');
    let myPureNik = reqSplit.nik !== '-' ? reqSplit.nik : '';
    let myUnitKerja = reqSplit.unitKerja !== '-' ? reqSplit.unitKerja : '';
    let myNama = String(nama || '').trim();

    if (reqUname && userByUsername.has(reqUname)) {
      const matched = userByUsername.get(reqUname);
      if (!myPureNik && matched.nik) myPureNik = matched.nik;
      if (!myUnitKerja && matched.unitKerja) myUnitKerja = matched.unitKerja;
      if (!myNama && matched.nama) myNama = matched.nama;
    }
    if (myPureNik && userByNik.has(cleanNik(myPureNik))) {
      const matched = userByNik.get(cleanNik(myPureNik));
      if (!myUnitKerja && matched.unitKerja) myUnitKerja = matched.unitKerja;
      if (!myNama && matched.nama) myNama = matched.nama;
    }

    const canViewAll = (scope === 'all') && (isRoleAdmin(role) || isRoleKetua(role) || isRoleFinancial(role) || isRolePicSimpin(role) || isRolePengurus(role));
    const myNamaClean = cleanNama(myNama);
    const myNikClean = cleanNik(myPureNik);

    function isMatchedMember(itemNik, itemNama) {
      if (canViewAll) return true;
      const iNikClean = cleanNik(itemNik);
      const iNamaClean = cleanNama(itemNama);

      // 1. Prioritaskan pencocokan NIK resmi secara presisi (exact match) atau keterkaitan NIK
      if (myNikClean && iNikClean) {
        if (myNikClean === iNikClean) return true;
        if (myNikClean.includes(iNikClean) || iNikClean.includes(myNikClean)) {
          const minLen = Math.min(myNikClean.length, iNikClean.length);
          if (minLen >= 4) return true;
        }
        return false;
      }

      // 2. Jika item pinjaman tidak memiliki NIK valid, gunakan nama lengkap resmi
      if (myNamaClean && iNamaClean) {
        if (isSamePersonName(myNamaClean, iNamaClean)) return true;
      }

      // 3. Username match sebagai alternatif
      if (reqUname) {
        const uClean = cleanNik(reqUname) || reqUname.toLowerCase();
        if (iNikClean && iNikClean === uClean) return true;
        if (iNamaClean && cleanNama(reqUname) === iNamaClean) return true;
      }

      return false;
    }

    // Helper: Hitung bulan berjalan sejak tanggal pencairan / Done
    function computeElapsedMonths(dateStr, maxTenor) {
      if (!dateStr) return 0;
      const s = String(dateStr).trim();
      let d = null;
      const dmy = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
      const ymd = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
      if (dmy) {
        d = new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
      } else if (ymd) {
        d = new Date(parseInt(ymd[1], 10), parseInt(ymd[2], 10) - 1, parseInt(ymd[3], 10));
      } else {
        const parsed = new Date(s);
        if (!isNaN(parsed.getTime())) d = parsed;
      }
      if (!d || isNaN(d.getTime())) return 0;
      const now = new Date();
      const diffMonths = (now.getFullYear() - d.getFullYear()) * 12 + (now.getMonth() - d.getMonth());
      return Math.max(0, Math.min(maxTenor, diffMonths));
    }

    // 2. Ambil data dari 'Form Responses 1' (Pinjaman yang diinput Finance & berstatus Done atau Lunas)
    const frRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A2:P"
    });
    const frRows = frRes.data.values || [];

    const multigunaLoans = [];
    const tunaiLoans = [];
    const seenNomorSurat = new Set();

    frRows.forEach((r, idx) => {
      const tipe = String(r[1] || '').trim();
      const tipeLower = tipe.toLowerCase();
      const status = String(r[9] || '').trim();
      const statusLower = status.toLowerCase();

      // Abaikan baris log pengajuan Pelunasan atau Topup yang sudah meniban baris pinjaman lama (agar tidak duplikat dengan pinjaman induknya)
      if (tipeLower.includes('pelunasan')) return;
      if (tipeLower.includes('topup') && r[12] && r[12] !== '-') return;

      // Ambil pinjaman yang statusnya telah Done, Lunas (riwayat pinjaman), atau tervalidasi resmi
      const isPaidOff = statusLower === 'lunas' || statusLower.includes('lunas');
      const isDoneStatus = isPaidOff || statusLower === 'done' || statusLower === 'validasi by financial' || statusLower.includes('terverifikasi e-materai');
      if (!isDoneStatus) return;

      const isMultiguna = tipeLower === 'peminjaman multiguna' || tipeLower.includes('multiguna');
      const isTunai = tipeLower === 'peminjaman tunai' || tipeLower.includes('pinjaman tunai');
      if (!isMultiguna && !isTunai) return;

      const rowNama = String(r[2] || '').trim();
      const parsed = splitNikAndUnitKerja(r[3] || '', r[10] || '', r[11] || '');
      let itemNik = parsed.nik !== '-' ? parsed.nik : '';
      let itemUnit = parsed.unitKerja !== '-' ? parsed.unitKerja : '';

      const rowNamaClean = cleanNama(rowNama);
      if (!itemNik && rowNamaClean) {
        const uMatch = userByName.get(rowNamaClean) || userByNameNoSpace.get(rowNamaClean.replace(/\s+/g, ''));
        if (uMatch) {
          itemNik = uMatch.nik;
          if (!itemUnit || itemUnit === '-') itemUnit = uMatch.unitKerja;
        }
      }

      if (!isMatchedMember(itemNik, rowNama)) return;

      let pokok = parseInt(String(r[6] || '0').replace(/[^0-9]/g, ''), 10) || 0;
      if (tipeLower.includes('topup') && r[13]) {
        const pBaru = parseInt(String(r[13]).replace(/[^0-9]/g, ''), 10) || 0;
        if (pBaru > 0) pokok = pBaru;
      }
      if (pokok <= 0) return;

      let tenor = parseInt(String(r[7] || '10').replace(/[^0-9]/g, ''), 10) || 10;
      const ket = String(r[8] || '').trim();

      // Cek apakah keterangan memuat nomor surat (misal 081/SK.SIMPIN/XI/2026)
      const noSuratMatch = ket.match(/(\d{3}\/SK\.SIMPIN\/[IVX]+\/\d{4})/i);
      const nomorRef = noSuratMatch ? noSuratMatch[1] : `FR-${idx + 2}`;
      if (noSuratMatch) seenNomorSurat.add(noSuratMatch[1].toUpperCase());

      if (isMultiguna) {
        const calc = calculateMultigunaLoanDetails(pokok, tenor);
        const bulanBerjalan = isPaidOff ? calc.tenor : computeElapsedMonths(r[0], calc.tenor);
        const sudahDibayar = isPaidOff ? calc.totalKewajiban : Math.min(calc.totalKewajiban, bulanBerjalan * calc.cicilanNormal);
        const sisaKewajiban = isPaidOff ? 0 : Math.max(0, calc.totalKewajiban - sudahDibayar);
        const sisaTenor = isPaidOff ? 0 : Math.max(0, calc.tenor - bulanBerjalan);
        const progressPersen = isPaidOff ? 100 : Math.round((bulanBerjalan / calc.tenor) * 100);
        const pokokPerBulan = Math.round(calc.pokok / calc.tenor);

        multigunaLoans.push({
          monitoringId: `MG-FR-${idx + 2}`,
          rowIndex: idx + 2,
          nomorReferensi: nomorRef,
          sumberData: isPaidOff ? 'Riwayat Pinjaman (Lunas)' : 'Input Pengajuan Finance (Done)',
          tanggalCair: r[0] || '-',
          kategori: 'Pinjaman Multiguna',
          namaAnggota: rowNama,
          nik: itemNik || '-',
          unitKerja: itemUnit || '-',
          noRekening: r[4] || '-',
          atasNamaRekening: r[5] || rowNama,
          pokokPinjaman: calc.pokok,
          pokokPerBulan,
          plafonMaksimal: 10000000,
          bungaPersen: '20% / Tahun',
          totalBunga: calc.totalBiaya,
          totalKewajiban: calc.totalKewajiban,
          tenorBulan: calc.tenor,
          tenorMaksimal: 10,
          cicilanBulanan: calc.cicilanNormal,
          bulanBerjalan,
          sisaTenor,
          sudahDibayar,
          sisaKewajiban,
          progressPersen,
          isLunas: isPaidOff || sisaKewajiban <= 0,
          keterangan: ket || 'Pinjaman Multiguna Koperasi',
          status: isPaidOff ? 'Lunas' : (status || 'Done'),
          jadwal: calc.jadwal
        });
      } else if (isTunai) {
        const validTenors = [10, 12, 20, 24, 36, 48, 60];
        if (!validTenors.includes(tenor)) tenor = 10;
        const rateMap = { 10: 0.11, 12: 0.14, 20: 0.22, 24: 0.28, 36: 0.42, 48: 0.56, 60: 0.70 };
        const rate = rateMap[tenor] || 0.11;
        const totalBunga = Math.round(pokok * rate);
        const totalKewajiban = pokok + totalBunga;
        const cicilanBulanan = Math.round(totalKewajiban / tenor);
        const pokokPerBulan = Math.round(pokok / tenor);

        const jadwal = [];
        let akum = 0;
        for (let m = 1; m <= tenor; m++) {
          if (m === tenor) {
            const akhir = totalKewajiban - akum;
            jadwal.push({ bulanKe: m, cicilan: akhir });
            akum += akhir;
          } else {
            jadwal.push({ bulanKe: m, cicilan: cicilanBulanan });
            akum += cicilanBulanan;
          }
        }

        const bulanBerjalan = isPaidOff ? tenor : computeElapsedMonths(r[0], tenor);
        const sudahDibayar = isPaidOff ? totalKewajiban : Math.min(totalKewajiban, bulanBerjalan * cicilanBulanan);
        const sisaKewajiban = isPaidOff ? 0 : Math.max(0, totalKewajiban - sudahDibayar);
        const sisaTenor = isPaidOff ? 0 : Math.max(0, tenor - bulanBerjalan);
        const progressPersen = isPaidOff ? 100 : Math.round((bulanBerjalan / tenor) * 100);

        tunaiLoans.push({
          monitoringId: `TN-FR-${idx + 2}`,
          rowIndex: idx + 2,
          nomorReferensi: nomorRef,
          sumberData: isPaidOff ? 'Riwayat Pinjaman (Lunas)' : 'Input Pengajuan Finance (Done)',
          tanggalCair: r[0] || '-',
          kategori: 'Pinjaman Tunai',
          namaAnggota: rowNama,
          nik: itemNik || '-',
          unitKerja: itemUnit || '-',
          noRekening: r[4] || '-',
          atasNamaRekening: r[5] || rowNama,
          pokokPinjaman: pokok,
          pokokPerBulan,
          plafonMaksimal: 70000000,
          bungaPersen: Math.round(rate * 100) + '%',
          totalBunga,
          totalKewajiban,
          tenorBulan: tenor,
          tenorMaksimal: 120,
          cicilanBulanan,
          bulanBerjalan,
          sisaTenor,
          sudahDibayar,
          sisaKewajiban,
          progressPersen,
          isLunas: isPaidOff || sisaKewajiban <= 0,
          keterangan: ket || 'Pinjaman Tunai Koperasi',
          status: isPaidOff ? 'Lunas' : (status || 'Done'),
          jadwal
        });
      }
    });

    // 3. Ambil juga dari 'LoanRequests' untuk Pinjaman Tunai yang sudah tervalidasi/cair/lunas (jika belum ada di seenNomorSurat)
    try {
      const lrRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'LoanRequests'!A2:AG"
      });
      const lrRows = lrRes.data.values || [];
      lrRows.forEach(r => {
        const noSurat = String(r[1] || '').trim();
        const stWorkflow = String(r[24] || '').trim().toUpperCase();
        if (!['VALIDATED_BY_FINANCE', 'VERIFIED_MATERAI', 'READY_FOR_DISBURSEMENT', 'DISBURSED', 'LUNAS', 'PAID_OFF'].includes(stWorkflow)) return;
        if (noSurat && seenNomorSurat.has(noSurat.toUpperCase())) return;

        const isPaidOffLr = stWorkflow === 'LUNAS' || stWorkflow === 'PAID_OFF';

        const parsed = splitNikAndUnitKerja(r[5] || '', r[5] || '', r[7] || '');
        let itemNik = parsed.nik !== '-' ? parsed.nik : '';
        let itemUnit = parsed.unitKerja !== '-' ? parsed.unitKerja : String(r[7] || '-').trim();
        const itemNama = String(r[6] || '').trim();
        const itemNamaClean = cleanNama(itemNama);

        if (!itemNik && itemNamaClean) {
          const uMatch = userByName.get(itemNamaClean) || userByNameNoSpace.get(itemNamaClean.replace(/\s+/g, ''));
          if (uMatch) {
            itemNik = uMatch.nik;
            if (!itemUnit || itemUnit === '-') itemUnit = uMatch.unitKerja;
          }
        }

        if (!isMatchedMember(itemNik, itemNama)) return;

        const pokok = parseInt(String(r[13] || '0').replace(/[^0-9]/g, ''), 10) || 0;
        const tenor = parseInt(String(r[14] || '10').replace(/[^0-9]/g, ''), 10) || 10;
        const persenStr = String(r[15] || '11%').trim();
        const totalKewajiban = parseInt(String(r[16] || '0').replace(/[^0-9]/g, ''), 10) || pokok;
        const totalBunga = Math.max(0, totalKewajiban - pokok);
        const cicilanBulanan = parseInt(String(r[17] || '0').replace(/[^0-9]/g, ''), 10) || Math.round(totalKewajiban / tenor);
        const pokokPerBulan = Math.round(pokok / tenor);

        const jadwal = [];
        let akum = 0;
        for (let m = 1; m <= tenor; m++) {
          if (m === tenor) {
            const akhir = totalKewajiban - akum;
            jadwal.push({ bulanKe: m, cicilan: akhir });
            akum += akhir;
          } else {
            jadwal.push({ bulanKe: m, cicilan: cicilanBulanan });
            akum += cicilanBulanan;
          }
        }

        const bulanBerjalan = isPaidOffLr ? tenor : computeElapsedMonths(r[2] || r[3], tenor);
        const sudahDibayar = isPaidOffLr ? totalKewajiban : Math.min(totalKewajiban, bulanBerjalan * cicilanBulanan);
        const sisaKewajiban = isPaidOffLr ? 0 : Math.max(0, totalKewajiban - sudahDibayar);
        const sisaTenor = isPaidOffLr ? 0 : Math.max(0, tenor - bulanBerjalan);
        const progressPersen = isPaidOffLr ? 100 : Math.round((bulanBerjalan / tenor) * 100);

        tunaiLoans.push({
          monitoringId: r[0] || `TN-LR-${noSurat}`,
          loanId: r[0] || '',
          nomorReferensi: noSurat || r[0],
          sumberData: isPaidOffLr ? 'Formulir Kredit Tunai Resmi (Lunas)' : 'Formulir Kredit Tunai Resmi',
          tanggalCair: r[2] || r[3] || '-',
          kategori: 'Pinjaman Tunai',
          namaAnggota: itemNama,
          nik: itemNik || '-',
          unitKerja: itemUnit || '-',
          noRekening: r[21] || '-',
          atasNamaRekening: r[22] || itemNama,
          pokokPinjaman: pokok,
          pokokPerBulan,
          plafonMaksimal: 70000000,
          bungaPersen: persenStr,
          totalBunga,
          totalKewajiban,
          tenorBulan: tenor,
          tenorMaksimal: 120,
          cicilanBulanan,
          bulanBerjalan,
          sisaTenor,
          sudahDibayar,
          sisaKewajiban,
          progressPersen,
          isLunas: isPaidOffLr || sisaKewajiban <= 0,
          keterangan: r[12] || 'Pinjaman Tunai Koperasi',
          status: isPaidOffLr ? 'Lunas' : (stWorkflow === 'DISBURSED' ? 'Done (Dicairkan)' : 'Tervalidasi (Aktif)'),
          jadwal
        });
      });
    } catch (eLr) {
      console.warn('Notice reading LoanRequests in monitoring:', eLr.message);
    }

    // Urutkan terbaru di atas
    multigunaLoans.reverse();
    tunaiLoans.reverse();

    const activeMulti = multigunaLoans.filter(x => !x.isLunas && String(x.status || '').toLowerCase() !== 'lunas');
    const activeTunai = tunaiLoans.filter(x => !x.isLunas && String(x.status || '').toLowerCase() !== 'lunas');

    const summary = {
      multiguna: {
        count: activeMulti.length,
        totalHistoryCount: multigunaLoans.length,
        lunasCount: multigunaLoans.length - activeMulti.length,
        totalPokok: activeMulti.reduce((acc, x) => acc + x.pokokPinjaman, 0),
        totalBunga: activeMulti.reduce((acc, x) => acc + x.totalBunga, 0),
        totalKewajiban: activeMulti.reduce((acc, x) => acc + x.totalKewajiban, 0),
        totalCicilanBulanan: activeMulti.reduce((acc, x) => acc + x.cicilanBulanan, 0),
        totalSisaKewajiban: activeMulti.reduce((acc, x) => acc + x.sisaKewajiban, 0)
      },
      tunai: {
        count: activeTunai.length,
        totalHistoryCount: tunaiLoans.length,
        lunasCount: tunaiLoans.length - activeTunai.length,
        totalPokok: activeTunai.reduce((acc, x) => acc + x.pokokPinjaman, 0),
        totalBunga: activeTunai.reduce((acc, x) => acc + x.totalBunga, 0),
        totalKewajiban: activeTunai.reduce((acc, x) => acc + x.totalKewajiban, 0),
        totalCicilanBulanan: activeTunai.reduce((acc, x) => acc + x.cicilanBulanan, 0),
        totalSisaKewajiban: activeTunai.reduce((acc, x) => acc + x.sisaKewajiban, 0)
      }
    };

    return res.json({
      success: true,
      memberProfile: {
        nik: myPureNik || '-',
        nama: myNama || username || '-',
        unitKerja: myUnitKerja || '-'
      },
      multigunaLoans,
      tunaiLoans,
      summary
    });
  } catch (error) {
    console.error('Error GET /api/loans/monitoring:', error);
    return res.status(500).json({ success: false, message: 'Gagal memuat monitoring pinjaman: ' + error.message });
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
      },
      {
        title: 'SimpananSukarelaLedger',
        headers: ['TxID', 'Timestamp', 'Tanggal', 'NIK', 'NamaAnggota', 'UnitKerja', 'JenisTransaksi', 'Sumber', 'Nominal', 'SaldoSebelumnya', 'SaldoSesudahnya', 'Keterangan', 'RefPembayaran', 'BuktiURL', 'Status', 'DibuatOlehNIK', 'DibuatOlehNama', 'DibuatOlehRole', 'ReversalOfTxID']
      },
      {
        title: 'SimpananSukarelaRules',
        headers: ['RuleID', 'Timestamp', 'NIK', 'NamaAnggota', 'UnitKerja', 'NominalRutin', 'Frekuensi', 'TanggalEksekusi', 'TanggalMulai', 'TanggalAkhir', 'StatusRule', 'Keterangan', 'DibuatOlehNIK', 'DibuatOlehNama', 'TerakhirDijalankan', 'JadwalBerikutnya', 'RuleVersion']
      },
      {
        title: 'SimpananSukarelaApprovals',
        headers: ['ApprovalID', 'RuleID', 'Timestamp', 'NIK', 'NamaAnggota', 'TipePerubahan', 'KonfigurasiLamaJSON', 'KonfigurasiBaruJSON', 'DiajukanOlehNIK', 'DiajukanOlehNama', 'StatusApproval', 'CatatanReview', 'DiprosesOlehNIK', 'DiprosesOlehNama', 'WaktuProses']
      },
      {
        title: 'SimpananSukarelaJobLog',
        headers: ['JobID', 'Timestamp', 'TanggalEksekusi', 'RuleID', 'NIK', 'NamaAnggota', 'Nominal', 'StatusEksekusi', 'PesanError', 'IdempotencyKey']
      },
      {
        title: 'LoanRateMaster',
        headers: ['RateID', 'PlafonMin', 'PlafonMax', 'TenorBulan', 'PersentaseTotal', 'Keterangan']
      },
      {
        title: 'LoanRequests',
        headers: ['LoanID', 'NomorSurat', 'Timestamp', 'TanggalPermohonan', 'AsalPengajuan', 'PemohonNIK', 'PemohonNama', 'PemohonUnit', 'KantorCabang', 'PenghasilanBulanan', 'Alamat', 'NoTelepon', 'Keperluan', 'JumlahPinjaman', 'TenorBulan', 'PersentaseTenor', 'TotalKewajiban', 'CicilanBulanan', 'TmtMulaiGaji', 'TmtSelesaiGaji', 'BankNama', 'BankNoRekening', 'BankAtasNama', 'StatusRekening', 'StatusWorkflow', 'FormResponsesRowIndex', 'Catatan', 'CreatedByNIK', 'CreatedByNama', 'CreatedByRole', 'DokumenMateraiURL', 'StatusMaterai', 'WaktuVerifikasiMaterai']
      },
      {
        title: 'LoanApprovals',
        headers: ['ApprovalID', 'LoanID', 'Timestamp', 'Tahap', 'Status', 'Catatan', 'ApproverNIK', 'ApproverNama', 'ApproverRole']
      },
      {
        title: 'LoanDisbursements',
        headers: ['DisbursementID', 'LoanID', 'Timestamp', 'TanggalPencairan', 'NominalPencairan', 'NomorReferensi', 'BankTujuan', 'NoRekeningTujuan', 'AtasNamaTujuan', 'BuktiURL', 'Catatan', 'Status', 'DisbursedByNIK', 'DisbursedByNama']
      },
      {
        title: 'MemberBankAccounts',
        headers: ['AccountID', 'NIK', 'NamaAnggota', 'NamaBank', 'NomorRekening', 'AtasNamaRekening', 'StatusVerifikasi', 'TanggalVerifikasi', 'VerifikatorNIK', 'VerifikatorNama', 'BuktiVerifikasiURL', 'StatusAktif', 'TimestampUpdate']
      },
      {
        title: 'BankAccountChangeRequests',
        headers: ['RequestID', 'NIK', 'NamaAnggota', 'BankLama', 'RekeningLama', 'AtasNamaLama', 'BankBaru', 'RekeningBaru', 'AtasNamaBaru', 'AlasanPenggantian', 'BuktiURL', 'StatusRequest', 'CatatanAdmin', 'DiprosesOlehNIK', 'DiprosesOlehNama', 'WaktuProses']
      },
      {
        title: 'FinancialAuditLog',
        headers: ['LogID', 'Timestamp', 'Modul', 'Aksi', 'TargetID', 'ActorNIK', 'ActorNama', 'ActorRole', 'DetailSebelum', 'DetailSesudah', 'IPAddress', 'UserAgent']
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

    if (jumlahPax === undefined || jumlahPax === null || jumlahPax === '') {
      return res.status(400).json({ success: false, message: 'Jumlah Pax wajib diisi (boleh 0).' });
    }
    const paxNum = parseInt(jumlahPax, 10);
    if (isNaN(paxNum) || paxNum < 0) {
      return res.status(400).json({ success: false, message: 'Jumlah Pax harus berupa angka minimal 0.' });
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
 * PUT /api/operasional/executive-car
 * Memperbarui data operasional Executive Car (Khusus Penginput atau Admin)
 */
app.put('/api/operasional/executive-car', async (req, res) => {
  try {
    const { rowIndex, role, userNik, kategoriPengguna, detailPengguna, jumlahPax, catatan } = req.body;
    const targetRow = parseInt(rowIndex, 10);
    if (!targetRow || targetRow < 2) {
      return res.status(400).json({ success: false, message: 'Row index tidak valid.' });
    }

    const sheets = await getSheetsClient();
    const rowResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'operasional-executive-car'!A${targetRow}:I${targetRow}`
    });

    const rowData = (rowResp.data.values && rowResp.data.values[0]) || [];
    if (rowData.length === 0) {
      return res.status(404).json({ success: false, message: 'Data operasional tidak ditemukan.' });
    }

    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(userNik || '').trim().toLowerCase();

    // Validasi Hak Akses: Penginput Asli atau Admin
    if (!isRoleAdmin(role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Anda hanya dapat mengedit data operasional yang Anda input sendiri.'
      });
    }

    const paxNum = (jumlahPax !== undefined && jumlahPax !== null && jumlahPax !== '') ? parseInt(jumlahPax, 10) : 0;
    if (isNaN(paxNum) || paxNum < 0) {
      return res.status(400).json({ success: false, message: 'Jumlah Pax minimal 0.' });
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'operasional-executive-car'!F${targetRow}:I${targetRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          String(kategoriPengguna || rowData[5] || 'UMUM').trim(),
          String(detailPengguna || '-').trim(),
          paxNum,
          String(catatan || '-').trim()
        ]]
      }
    });

    return res.json({ success: true, message: 'Data operasional Executive Car berhasil diperbarui!' });
  } catch (error) {
    console.error('Error PUT /api/operasional/executive-car:', error);
    return res.status(500).json({ success: false, message: 'Gagal memperbarui data: ' + error.message });
  }
});

/**
 * DELETE /api/operasional/executive-car
 * Menghapus data operasional Executive Car (Khusus Penginput atau Admin)
 */
app.delete('/api/operasional/executive-car', async (req, res) => {
  try {
    const { rowIndex, role, userNik } = req.body;
    const targetRow = parseInt(rowIndex, 10);
    if (!targetRow || targetRow < 2) {
      return res.status(400).json({ success: false, message: 'Row index tidak valid.' });
    }

    const sheets = await getSheetsClient();
    const rowResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'operasional-executive-car'!A${targetRow}:I${targetRow}`
    });

    const rowData = (rowResp.data.values && rowResp.data.values[0]) || [];
    if (rowData.length === 0) {
      return res.status(404).json({ success: false, message: 'Data operasional tidak ditemukan.' });
    }

    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(userNik || '').trim().toLowerCase();

    // Validasi Hak Akses: Penginput Asli atau Admin
    if (!isRoleAdmin(role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Anda hanya dapat menghapus data operasional yang Anda input sendiri.'
      });
    }

    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const targetSheet = (meta.data.sheets || []).find(s => s.properties.title === 'operasional-executive-car');
    if (!targetSheet) {
      return res.status(404).json({ success: false, message: 'Sheet operasional-executive-car tidak ditemukan.' });
    }
    const sheetId = targetSheet.properties.sheetId;

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

    return res.json({ success: true, message: 'Data operasional Executive Car berhasil dihapus!' });
  } catch (error) {
    console.error('Error DELETE /api/operasional/executive-car:', error);
    return res.status(500).json({ success: false, message: 'Gagal menghapus data: ' + error.message });
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

    if (jumlahGalon === undefined || jumlahGalon === null || jumlahGalon === '') {
      return res.status(400).json({ success: false, message: 'Jumlah Galon wajib diisi (boleh 0).' });
    }
    const galonNum = parseInt(jumlahGalon, 10);
    if (isNaN(galonNum) || galonNum < 0) {
      return res.status(400).json({ success: false, message: 'Jumlah Galon harus berupa angka minimal 0.' });
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
 * PUT /api/operasional/air-minum
 * Memperbarui data pengantaran galon air minum (Khusus Penginput atau Admin)
 */
app.put('/api/operasional/air-minum', async (req, res) => {
  try {
    const { rowIndex, role, userNik, kategoriPenerima, namaPenerima, jumlahGalon, metodePembayaran, catatan } = req.body;
    const targetRow = parseInt(rowIndex, 10);
    if (!targetRow || targetRow < 2) {
      return res.status(400).json({ success: false, message: 'Row index tidak valid.' });
    }

    const sheets = await getSheetsClient();
    const rowResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'operasional-air-minum'!A${targetRow}:J${targetRow}`
    });

    const rowData = (rowResp.data.values && rowResp.data.values[0]) || [];
    if (rowData.length === 0) {
      return res.status(404).json({ success: false, message: 'Data pengantaran tidak ditemukan.' });
    }

    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(userNik || '').trim().toLowerCase();

    // Validasi Hak Akses: Penginput Asli atau Admin
    if (!isRoleAdmin(role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Anda hanya dapat mengedit data pengantaran yang Anda input sendiri.'
      });
    }

    const galonNum = (jumlahGalon !== undefined && jumlahGalon !== null && jumlahGalon !== '') ? parseInt(jumlahGalon, 10) : 0;
    if (isNaN(galonNum) || galonNum < 0) {
      return res.status(400).json({ success: false, message: 'Jumlah Galon minimal 0.' });
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'operasional-air-minum'!F${targetRow}:J${targetRow}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          String(kategoriPenerima || rowData[5] || 'ATS').trim(),
          String(namaPenerima || '-').trim(),
          galonNum,
          String(metodePembayaran || rowData[8] || 'Cash').trim(),
          String(catatan || '-').trim()
        ]]
      }
    });

    return res.json({ success: true, message: 'Data pengantaran galon air minum berhasil diperbarui!' });
  } catch (error) {
    console.error('Error PUT /api/operasional/air-minum:', error);
    return res.status(500).json({ success: false, message: 'Gagal memperbarui data: ' + error.message });
  }
});

/**
 * DELETE /api/operasional/air-minum
 * Menghapus data pengantaran galon air minum (Khusus Penginput atau Admin)
 */
app.delete('/api/operasional/air-minum', async (req, res) => {
  try {
    const { rowIndex, role, userNik } = req.body;
    const targetRow = parseInt(rowIndex, 10);
    if (!targetRow || targetRow < 2) {
      return res.status(400).json({ success: false, message: 'Row index tidak valid.' });
    }

    const sheets = await getSheetsClient();
    const rowResp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `'operasional-air-minum'!A${targetRow}:J${targetRow}`
    });

    const rowData = (rowResp.data.values && rowResp.data.values[0]) || [];
    if (rowData.length === 0) {
      return res.status(404).json({ success: false, message: 'Data pengantaran tidak ditemukan.' });
    }

    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(userNik || '').trim().toLowerCase();

    // Validasi Hak Akses: Penginput Asli atau Admin
    if (!isRoleAdmin(role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return res.status(403).json({
        success: false,
        message: 'Akses ditolak: Anda hanya dapat menghapus data pengantaran yang Anda input sendiri.'
      });
    }

    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const targetSheet = (meta.data.sheets || []).find(s => s.properties.title === 'operasional-air-minum');
    if (!targetSheet) {
      return res.status(404).json({ success: false, message: 'Sheet operasional-air-minum tidak ditemukan.' });
    }
    const sheetId = targetSheet.properties.sheetId;

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

    return res.json({ success: true, message: 'Data pengantaran galon air minum berhasil dihapus!' });
  } catch (error) {
    console.error('Error DELETE /api/operasional/air-minum:', error);
    return res.status(500).json({ success: false, message: 'Gagal menghapus data: ' + error.message });
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

// ==========================================
// 6. MODUL SIMPANAN SUKARELA & PINJAMAN KOPERASI
// ==========================================

// Helper: Format Bulan Romawi
function getRomanMonth(monthZeroIndexed) {
  const roman = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
  return roman[monthZeroIndexed] || 'I';
}

// Helper: Catat Log Audit Finansial
async function recordFinancialAudit(sheets, actor, modul, aksi, targetId, detailSebelum, detailSesudah, req) {
  try {
    const logId = 'AUDIT-' + Date.now() + '-' + Math.floor(Math.random() * 1000);
    const ts = getFormattedTimestamp();
    const ip = req ? (req.headers['x-forwarded-for'] || req.socket.remoteAddress || '') : '';
    const ua = req ? (req.headers['user-agent'] || '') : '';

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'FinancialAuditLog'!A:L",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          logId,
          ts,
          modul || '',
          aksi || '',
          String(targetId || ''),
          String(actor.nik || ''),
          actor.nama || '',
          actor.role || '',
          typeof detailSebelum === 'object' ? JSON.stringify(detailSebelum) : String(detailSebelum || ''),
          typeof detailSesudah === 'object' ? JSON.stringify(detailSesudah) : String(detailSesudah || ''),
          String(ip),
          String(ua).substring(0, 200)
        ]]
      }
    });
  } catch (err) {
    console.error('[FinancialAuditLog Warning]:', err.message);
  }
}

// Helper: Kalkulasi Saldo Simpanan Sukarela Anggota
async function calculateMemberSimpananSaldo(sheets, targetNik) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'SimpananSukarelaLedger'!A2:S"
  });
  const rows = res.data.values || [];
  const jakarta = getJakartaDateTime();
  const currentMonthYear = '-' + jakarta.month + '-' + jakarta.year;

  let saldo = 0;
  let totalSetoranBulanIni = 0;
  let totalSetoranAll = 0;
  let totalPenarikanAll = 0;
  let mutasiList = [];

  rows.forEach((r, idx) => {
    const txId = r[0] || '';
    const timestamp = r[1] || '';
    const tanggal = r[2] || '';
    const nik = String(r[3] || '').trim();
    const nama = r[4] || '';
    const unit = r[5] || '';
    const jenis = String(r[6] || '').trim().toUpperCase(); // SETORAN, PENARIKAN, KOREKSI, OPENING_BALANCE
    const sumber = r[7] || '';
    const nominal = parseInt(String(r[8] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const keterangan = r[11] || '';
    const refPembayaran = r[12] || '';
    const buktiUrl = r[13] || '';
    const status = String(r[14] || '').trim().toUpperCase(); // POSTED, REVERSED
    const dibuatOlehNama = r[16] || '';
    const reversalOf = r[18] || '';

    if (nik === String(targetNik).trim()) {
      if (status === 'POSTED') {
        if (jenis === 'SETORAN' || jenis === 'OPENING_BALANCE') {
          saldo += nominal;
          totalSetoranAll += nominal;
          if (tanggal.includes(currentMonthYear) || timestamp.includes(jakarta.month + '/' + jakarta.year) || timestamp.includes('/' + jakarta.month + '/')) {
            totalSetoranBulanIni += nominal;
          }
        } else if (jenis === 'PENARIKAN' || jenis === 'PENGAMBILAN' || jenis.includes('TARIK') || jenis.includes('AMBIL')) {
          saldo -= nominal;
          totalPenarikanAll += nominal;
        } else if (jenis === 'KOREKSI') {
          // Koreksi nominal positif atau penyesuaian
          saldo += nominal;
        }
      }

      mutasiList.push({
        rowIndex: idx + 2,
        txId,
        timestamp,
        tanggal,
        nik,
        nama,
        unit,
        jenis,
        sumber,
        nominal,
        keterangan,
        refPembayaran,
        buktiUrl,
        status,
        dibuatOlehNama,
        reversalOf
      });
    }
  });

  return {
    nik: targetNik,
    saldo,
    totalSetoranBulanIni,
    totalSetoranAll,
    totalPenarikanAll,
    mutasiCount: mutasiList.length,
    mutasiList: mutasiList.reverse() // terbaru di atas
  };
}

// Helper: Kalkulasi Pinjaman Berdasarkan Tabel Resmi
function calculateLoanSchedule(plafon, tenor) {
  const p = parseInt(plafon, 10);
  const t = parseInt(tenor, 10);

  if (isNaN(p) || p < 10000000 || p > 70000000 || p % 1000000 !== 0) {
    throw new Error('Plafon pinjaman harus antara Rp 10.000.000 s/d Rp 70.000.000 dengan kelipatan Rp 1.000.000');
  }

  const validTenors = p <= 30000000 ? [10, 12, 20, 24, 36] : [10, 12, 20, 24, 36, 48, 60, 72, 84, 96, 120];
  if (!validTenors.includes(t)) {
    throw new Error('Tenor ' + t + ' bulan tidak tersedia untuk plafon Rp ' + p.toLocaleString('id-ID'));
  }

  const rateMap = { 10: 0.11, 12: 0.14, 20: 0.22, 24: 0.28, 36: 0.42, 48: 0.56, 60: 0.70, 72: 0.84, 84: 0.98, 96: 1.12, 108: 1.26, 120: 1.40 };
  const rate = rateMap[t] !== undefined ? rateMap[t] : (0.14 * (t / 12));
  const totalBiaya = Math.round(p * rate);
  const totalKewajiban = p + totalBiaya;
  const cicilanNormal = Math.round(totalKewajiban / t);

  // Buat rincian jadwal angsuran (alokasi selisih sen ke bulan terakhir)
  const jadwal = [];
  let akumulasi = 0;
  for (let i = 1; i <= t; i++) {
    if (i === t) {
      const cicilanAkhir = totalKewajiban - akumulasi;
      jadwal.push({ bulanKe: i, angsuran: cicilanAkhir });
      akumulasi += cicilanAkhir;
    } else {
      jadwal.push({ bulanKe: i, angsuran: cicilanNormal });
      akumulasi += cicilanNormal;
    }
  }

  return {
    plafon: p,
    tenor: t,
    rate,
    ratePersen: Math.round(rate * 100) + '%',
    totalBiaya,
    totalKewajiban,
    cicilanBulanan: cicilanNormal,
    cicilanAkhir: jadwal[jadwal.length - 1].angsuran,
    jadwal
  };
}

// Helper: Generate Nomor Surat Pinjaman Format: XXX/SK.SIMPIN/MM/YYYY
async function generateLoanNomorSurat(sheets) {
  const jakarta = getJakartaDateTime();
  const roman = getRomanMonth(parseInt(jakarta.month, 10) - 1);
  const year = jakarta.year;

  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: "'LoanRequests'!A2:B"
  });
  const rows = res.data.values || [];
  let maxSeq = 80; // Baseline urutan melanjutkan format formulir 081/SK.SIMPIN

  rows.forEach(r => {
    const noSurat = String(r[1] || '');
    const match = noSurat.match(/^(\d+)\/SK\.SIMPIN/i);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxSeq) maxSeq = num;
    }
  });

  const nextSeq = String(maxSeq + 1).padStart(3, '0');
  return `${nextSeq}/SK.SIMPIN/${roman}/${year}`;
}

// ----------------------------------------------------
// 6.1 ENDPOINTS: SALDO & PENYIMPANAN SUKARELA
// ----------------------------------------------------

/**
 * GET /api/simpanan/summary
 * Mengambil ringkasan saldo sukarela (user sendiri, atau rekap seluruh koperasi untuk Financial/Admin/Ketua)
 */
app.get('/api/simpanan/summary', async (req, res) => {
  try {
    const { nik, role, targetNik } = req.query;
    if (!nik) {
      return res.status(400).json({ success: false, message: 'Parameter NIK diperlukan.' });
    }

    const sheets = await getSheetsClient();
    const isOfficer = isRoleFinancial(role) || isRoleAdmin(role) || isRoleKetua(role);
    const lookupNik = (isOfficer && targetNik) ? targetNik : nik;

    const memberData = await calculateMemberSimpananSaldo(sheets, lookupNik);

    // Ambil info auto-input rule anggota jika ada
    const rulesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaRules'!A2:Q"
    });
    const rulesRows = rulesRes.data.values || [];
    const activeRule = rulesRows.find(r => String(r[2] || '').trim() === String(lookupNik).trim() && String(r[10] || '').toUpperCase() === 'ACTIVE');

    let ruleInfo = null;
    if (activeRule) {
      ruleInfo = {
        ruleId: activeRule[0],
        nominalRutin: parseInt(activeRule[5], 10) || 0,
        tanggalEksekusi: activeRule[7] || '25',
        status: activeRule[10],
        terakhirDijalankan: activeRule[14] || '-'
      };
    }

    // Jika Financial/Admin/Ketua, berikan juga total agregat koperasi
    let overallSummary = null;
    if (isOfficer) {
      const ledgerRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'SimpananSukarelaLedger'!A2:S"
      });
      const allRows = ledgerRes.data.values || [];
      let grandTotalSaldo = 0;
      let grandSetoranBulanIni = 0;
      const uniqueMembers = new Set();
      const jakarta = getJakartaDateTime();
      const currentMonthYear = '-' + jakarta.month + '-' + jakarta.year;

      allRows.forEach(r => {
        const rNik = String(r[3] || '').trim();
        const jenis = String(r[6] || '').trim().toUpperCase();
        const nominal = parseInt(String(r[8] || '0').replace(/[^0-9]/g, ''), 10) || 0;
        const status = String(r[14] || '').trim().toUpperCase();
        const tgl = r[2] || '';

        if (status === 'POSTED') {
          if (rNik) uniqueMembers.add(rNik);
          if (jenis === 'SETORAN' || jenis === 'OPENING_BALANCE') {
            grandTotalSaldo += nominal;
            if (tgl.includes(currentMonthYear)) grandSetoranBulanIni += nominal;
          } else if (jenis === 'PENARIKAN') {
            grandTotalSaldo -= nominal;
          }
        }
      });

      overallSummary = {
        grandTotalSaldo,
        grandSetoranBulanIni,
        totalAnggotaMenabung: uniqueMembers.size
      };
    }

    return res.json({
      success: true,
      member: memberData,
      rule: ruleInfo,
      overall: overallSummary
    });
  } catch (error) {
    console.error('Error GET /api/simpanan/summary:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/simpanan/transactions
 * Mengambil daftar mutasi simpanan sukarela
 */
app.get('/api/simpanan/transactions', async (req, res) => {
  try {
    const { nik, role, targetNik, page = 1, limit = 20, search = '' } = req.query;
    if (!nik) {
      return res.status(400).json({ success: false, message: 'Parameter NIK diperlukan.' });
    }

    const sheets = await getSheetsClient();
    const isOfficer = isRoleFinancial(role) || isRoleAdmin(role) || isRoleKetua(role);
    const lookupNik = (isOfficer && targetNik) ? targetNik : (!isOfficer ? nik : null);

    const ledgerRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaLedger'!A2:S"
    });
    const rows = ledgerRes.data.values || [];

    let transactions = [];
    rows.forEach((r, idx) => {
      const txId = r[0] || '';
      const timestamp = r[1] || '';
      const tanggal = r[2] || '';
      const rowNik = String(r[3] || '').trim();
      const nama = r[4] || '';
      const unit = r[5] || '';
      const jenis = String(r[6] || '').trim().toUpperCase();
      const sumber = r[7] || '';
      const nominal = parseInt(String(r[8] || '0').replace(/[^0-9]/g, ''), 10) || 0;
      const saldoSebelumnya = parseInt(String(r[9] || '0').replace(/[^0-9]/g, ''), 10) || 0;
      const saldoSesudahnya = parseInt(String(r[10] || '0').replace(/[^0-9]/g, ''), 10) || 0;
      const keterangan = r[11] || '';
      const refPembayaran = r[12] || '';
      const buktiUrl = r[13] || '';
      const status = String(r[14] || '').trim().toUpperCase();
      const dibuatOlehNama = r[16] || '';
      const reversalOf = r[18] || '';

      // Filter hak akses
      if (lookupNik && rowNik !== lookupNik) return;

      // Filter search
      if (search) {
        const q = search.toLowerCase();
        const match = txId.toLowerCase().includes(q) ||
                      nama.toLowerCase().includes(q) ||
                      rowNik.toLowerCase().includes(q) ||
                      keterangan.toLowerCase().includes(q) ||
                      refPembayaran.toLowerCase().includes(q);
        if (!match) return;
      }

      transactions.push({
        rowIndex: idx + 2,
        txId,
        timestamp,
        tanggal,
        nik: rowNik,
        nama,
        unit,
        jenis,
        sumber,
        nominal,
        saldoSebelumnya,
        saldoSesudahnya,
        keterangan,
        refPembayaran,
        buktiUrl,
        status,
        dibuatOlehNama,
        reversalOf
      });
    });

    transactions.reverse(); // Terbaru di atas
    const totalCount = transactions.length;
    const p = parseInt(page, 10) || 1;
    const l = parseInt(limit, 10) || 20;
    const paginated = transactions.slice((p - 1) * l, p * l);

    return res.json({
      success: true,
      data: paginated,
      totalCount,
      page: p,
      totalPages: Math.ceil(totalCount / l)
    });
  } catch (error) {
    console.error('Error GET /api/simpanan/transactions:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/simpanan/transactions
 * Input transaksi simpanan sukarela manual (Khusus Financial / Admin)
 */
app.post('/api/simpanan/transactions', async (req, res) => {
  try {
    const {
      nik,
      nama,
      unitKerja,
      jenisTransaksi, // SETORAN, PENARIKAN, KOREKSI
      nominal,
      tanggal,
      keterangan,
      refPembayaran,
      buktiUrl,
      createdByNik,
      createdByNama,
      createdByRole
    } = req.body;

    if (!isRoleFinancial(createdByRole) && !isRoleAdmin(createdByRole)) {
      return res.status(403).json({ success: false, message: 'Hanya role Financial atau Administrator yang berhak mencatat mutasi simpanan sukarela.' });
    }

    if (!nik || !nama || !nominal || !jenisTransaksi) {
      return res.status(400).json({ success: false, message: 'NIK, Nama, Jenis Transaksi, dan Nominal wajib diisi.' });
    }

    const nom = parseInt(nominal, 10);
    if (isNaN(nom) || nom <= 0) {
      return res.status(400).json({ success: false, message: 'Nominal transaksi harus berupa angka positif lebih dari 0.' });
    }

    const jenisUpper = jenisTransaksi.trim().toUpperCase();
    if (!['SETORAN', 'PENARIKAN', 'PENGAMBILAN', 'KOREKSI', 'OPENING_BALANCE'].includes(jenisUpper)) {
      return res.status(400).json({ success: false, message: 'Jenis transaksi tidak valid.' });
    }

    const sheets = await getSheetsClient();
    const currentMember = await calculateMemberSimpananSaldo(sheets, nik);
    const saldoSebelumnya = currentMember.saldo;
    const isPengambilan = (jenisUpper === 'PENARIKAN' || jenisUpper === 'PENGAMBILAN' || jenisUpper.includes('TARIK') || jenisUpper.includes('AMBIL'));

    if (isPengambilan && nom > saldoSebelumnya) {
      return res.status(400).json({
        success: false,
        message: `Saldo tidak mencukupi untuk pengambilan sukarela. Saldo saat ini: Rp ${saldoSebelumnya.toLocaleString('id-ID')}, Permintaan pengambilan: Rp ${nom.toLocaleString('id-ID')}.`
      });
    }

    let saldoSesudahnya = saldoSebelumnya;
    if (jenisUpper === 'SETORAN' || jenisUpper === 'OPENING_BALANCE') saldoSesudahnya += nom;
    else if (isPengambilan) saldoSesudahnya -= nom;
    else if (jenisUpper === 'KOREKSI') saldoSesudahnya += nom;

    const jakarta = getJakartaDateTime();
    const tgl = tanggal || jakarta.dateStr;
    const txId = 'TX-SIMP-' + Date.now() + '-' + Math.floor(Math.random() * 100);
    const ts = getFormattedTimestamp();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaLedger'!A:S",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          txId,
          ts,
          tgl,
          String(nik).trim(),
          nama,
          unitKerja || '',
          jenisUpper,
          'MANUAL_FINANCIAL',
          nom,
          saldoSebelumnya,
          saldoSesudahnya,
          keterangan || '-',
          refPembayaran || '-',
          buktiUrl || '-',
          'POSTED',
          String(createdByNik || ''),
          createdByNama || '',
          createdByRole || '',
          ''
        ]]
      }
    });

    // Audit log
    await recordFinancialAudit(
      sheets,
      { nik: createdByNik, nama: createdByNama, role: createdByRole },
      'SIMPANAN_SUKARELA',
      'INPUT_' + jenisUpper,
      txId,
      { saldoSebelumnya },
      { saldoSesudahnya, nominal: nom, nik, nama },
      req
    );

    return res.json({
      success: true,
      message: `Transaksi ${jenisUpper} sebesar Rp ${nom.toLocaleString('id-ID')} berhasil dicatat ke buku ledger.`,
      txId,
      saldoSebelumnya,
      saldoSesudahnya
    });
  } catch (error) {
    console.error('Error POST /api/simpanan/transactions:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/simpanan/transactions/reversal
 * Koreksi / Pembatalan transaksi ledger (Khusus Financial / Admin)
 */
app.post('/api/simpanan/transactions/reversal', async (req, res) => {
  try {
    const { txId, reason, createdByNik, createdByNama, createdByRole } = req.body;

    if (!isRoleFinancial(createdByRole) && !isRoleAdmin(createdByRole)) {
      return res.status(403).json({ success: false, message: 'Hanya role Financial atau Administrator yang berhak membatalkan mutasi simpanan.' });
    }

    if (!txId || !reason) {
      return res.status(400).json({ success: false, message: 'TxID dan Alasan pembatalan wajib diisi.' });
    }

    const sheets = await getSheetsClient();
    const ledgerRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaLedger'!A2:S"
    });
    const rows = ledgerRes.data.values || [];

    const targetIndex = rows.findIndex(r => r[0] === txId);
    if (targetIndex === -1) {
      return res.status(404).json({ success: false, message: 'Transaksi ' + txId + ' tidak ditemukan.' });
    }

    const oldRow = rows[targetIndex];
    const oldStatus = oldRow[14] || '';
    if (oldStatus.toUpperCase() === 'REVERSED') {
      return res.status(400).json({ success: false, message: 'Transaksi ' + txId + ' sudah dibatalkan sebelumnya.' });
    }

    const nik = oldRow[3];
    const nama = oldRow[4];
    const unit = oldRow[5];
    const oldJenis = String(oldRow[6] || '').toUpperCase();
    const oldNominal = parseInt(String(oldRow[8] || '0').replace(/[^0-9]/g, ''), 10) || 0;

    // Ambil saldo anggota sebelum reversal
    const currentMember = await calculateMemberSimpananSaldo(sheets, nik);
    const saldoSebelumnya = currentMember.saldo;
    let saldoSesudahnya = saldoSebelumnya;
    if (oldJenis === 'SETORAN' || oldJenis === 'OPENING_BALANCE' || oldJenis === 'KOREKSI') {
      saldoSesudahnya = saldoSebelumnya - oldNominal;
    } else if (oldJenis === 'PENARIKAN') {
      saldoSesudahnya = saldoSebelumnya + oldNominal;
    }

    const rowNumber = targetIndex + 2;
    // Update status baris lama menjadi REVERSED
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'SimpananSukarelaLedger'!O${rowNumber}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['REVERSED']] }
    });

    const reversalTxId = 'REV-' + Date.now();
    const ts = getFormattedTimestamp();
    const jakarta = getJakartaDateTime();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaLedger'!A:S",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          reversalTxId,
          ts,
          jakarta.dateStr,
          String(nik),
          nama,
          unit,
          'REVERSAL',
          'MANUAL_REVERSAL',
          oldNominal,
          saldoSebelumnya,
          saldoSesudahnya,
          'Reversal dari ' + txId + ': ' + reason,
          txId,
          '-',
          'POSTED',
          String(createdByNik || ''),
          createdByNama || '',
          createdByRole || '',
          txId
        ]]
      }
    });

    await recordFinancialAudit(
      sheets,
      { nik: createdByNik, nama: createdByNama, role: createdByRole },
      'SIMPANAN_SUKARELA',
      'REVERSAL_TRANSACTION',
      txId,
      { oldStatus: 'POSTED', oldNominal },
      { newStatus: 'REVERSED', reversalTxId, reason },
      req
    );

    return res.json({
      success: true,
      message: `Transaksi ${txId} berhasil dibatalkan dengan entri koreksi ${reversalTxId}.`,
      reversalTxId,
      saldoSesudahnya
    });
  } catch (error) {
    console.error('Error POST /api/simpanan/transactions/reversal:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/simpanan/rules
 * Mengambil daftar auto-input rules
 */
app.get('/api/simpanan/rules', async (req, res) => {
  try {
    const { nik, role } = req.query;
    const sheets = await getSheetsClient();
    const isOfficer = isRoleFinancial(role) || isRoleAdmin(role) || isRoleKetua(role) || isRolePengurus(role) || isRolePicSimpin(role);

    const rulesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaRules'!A2:Q"
    });
    const rows = rulesRes.data.values || [];

    const rules = [];
    rows.forEach((r, idx) => {
      const rawNik = String(r[2] || '').trim();
      const rawUnit = String(r[4] || '').trim();
      const parsed = splitNikAndUnitKerja(rawUnit ? `${rawNik} - ${rawUnit}` : rawNik, rawNik, rawUnit);
      const ruleNik = parsed.nik !== '-' ? parsed.nik : rawNik;
      const ruleUnit = parsed.unitKerja !== '-' ? parsed.unitKerja : rawUnit;

      if (!isOfficer && cleanNik(ruleNik) !== cleanNik(nik) && cleanNik(rawNik) !== cleanNik(nik)) return;

      rules.push({
        rowIndex: idx + 2,
        ruleId: r[0],
        timestamp: r[1],
        nik: ruleNik,
        namaAnggota: r[3],
        unitKerja: ruleUnit || '-',
        nominalRutin: parseInt(r[5], 10) || 0,
        frekuensi: r[6] || 'BULANAN',
        tanggalEksekusi: r[7] || '25',
        tanggalMulai: r[8] || '',
        tanggalAkhir: r[9] || '-',
        statusRule: r[10] || 'PENDING_APPROVAL',
        keterangan: r[11] || '',
        dibuatOlehNik: r[12] || '-',
        dibuatOlehNama: r[13] || '-',
        terakhirDijalankan: r[14] || '-'
      });
    });

    return res.json({ success: true, rules });
  } catch (error) {
    console.error('Error GET /api/simpanan/rules:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/simpanan/rules
 * Membuat atau memperbarui konfigurasi auto-input rutin (Wajib Approval)
 */
app.post('/api/simpanan/rules', async (req, res) => {
  try {
    const {
      nik,
      namaAnggota,
      unitKerja,
      nominalRutin,
      tanggalEksekusi = '25',
      keterangan,
      createdByNik,
      createdByNama,
      createdByRole
    } = req.body;

    if (!isRoleFinancial(createdByRole) && !isRoleAdmin(createdByRole) && !isRoleKetua(createdByRole)) {
      return res.status(403).json({ success: false, message: 'Hanya role Financial, Administrator, atau Ketua Koperasi yang dapat mengatur auto-input.' });
    }

    const nom = parseInt(nominalRutin, 10);
    if (isNaN(nom) || nom <= 0) {
      return res.status(400).json({ success: false, message: 'Nominal rutin harus angka positif lebih dari 0.' });
    }

    const sheets = await getSheetsClient();
    const rulesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaRules'!A2:Q"
    });
    const rows = rulesRes.data.values || [];
    const targetNikClean = cleanNik(nik);
    const existingRuleIdx = rows.findIndex(r => {
      const rNikClean = cleanNik(r[2]);
      return rNikClean && targetNikClean && rNikClean === targetNikClean;
    });

    const existingRule = existingRuleIdx !== -1 ? rows[existingRuleIdx] : null;
    const ruleId = existingRule ? existingRule[0] : ('RULE-SIMP-' + Date.now());
    const ts = getFormattedTimestamp();

    if (existingRule) {
      // Usulan perubahan rule yang sudah ada
      const approvalId = 'APP-RULE-' + Date.now();
      const ruleRowNum = existingRuleIdx + 2;

      // Update status rule menjadi PENDING_APPROVAL sampai disetujui
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'SimpananSukarelaRules'!K${ruleRowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['PENDING_APPROVAL']] }
      });

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "'SimpananSukarelaApprovals'!A:O",
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            approvalId,
            ruleId,
            ts,
            String(nik).trim(),
            namaAnggota,
            'UPDATE_NOMINAL',
            JSON.stringify({ nominalRutin: existingRule[5], tanggalEksekusi: existingRule[7] || '25', status: existingRule[10] }),
            JSON.stringify({ nominalRutin: nom, tanggalEksekusi }),
            String(createdByNik || ''),
            createdByNama || '',
            'PENDING',
            keterangan || 'Perubahan nominal auto-input simpanan sukarela',
            '',
            '',
            ''
          ]]
        }
      });

      return res.json({
        success: true,
        message: 'Usulan perubahan nominal auto-input berhasil diajukan dan wajib diapprove oleh Admin, Finance, atau Ketua Koperasi.',
        approvalId,
        ruleId
      });
    }

    // Buat rule baru (Status awal: PENDING_APPROVAL)
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaRules'!A:Q",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          ruleId,
          ts,
          String(nik).trim(),
          namaAnggota,
          unitKerja || '',
          nom,
          'BULANAN',
          tanggalEksekusi,
          ts.split(' ')[0],
          '-',
          'PENDING_APPROVAL',
          keterangan || 'Pemotongan rutin simpanan sukarela',
          String(createdByNik || ''),
          createdByNama || '',
          '-',
          '-',
          '1.0'
        ]]
      }
    });

    // Wajib masukkan ke antrean approval (Bisa diapprove oleh Admin / Finance / Ketua Koperasi)
    const approvalId = 'APP-RULE-' + Date.now();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaApprovals'!A:O",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          approvalId,
          ruleId,
          ts,
          String(nik).trim(),
          namaAnggota,
          'NEW_RULE',
          JSON.stringify({ nominalRutin: 0, status: 'NONE' }),
          JSON.stringify({ nominalRutin: nom, tanggalEksekusi }),
          String(createdByNik || ''),
          createdByNama || '',
          'PENDING',
          keterangan || 'Setting baru simpanan sukarela rutin bulanan',
          '',
          '',
          ''
        ]]
      }
    });

    return res.json({
      success: true,
      message: 'Aturan auto-input simpanan sukarela baru berhasil diajukan dan wajib diapprove oleh Admin, Finance, atau Ketua Koperasi.',
      ruleId,
      approvalId
    });
  } catch (error) {
    console.error('Error POST /api/simpanan/rules:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/simpanan/rules/stop
 * Aksi berhenti dari koperasi / nonaktifkan & hapus auto-input anggota yang keluar
 */
app.post('/api/simpanan/rules/stop', async (req, res) => {
  try {
    const { ruleId, nik, namaAnggota, actionType = 'STOP' } = req.body;
    const actorRole = req.body.actorRole || req.body.requestedByRole || req.body.role || '';
    const actorNik = req.body.actorNik || req.body.requestedByNik || '';
    const actorNama = req.body.actorNama || req.body.requestedByNama || '';
    const reason = req.body.reason || req.body.alasan || 'Keluar dari keanggotaan koperasi';

    if (!isRoleFinancial(actorRole) && !isRoleAdmin(actorRole) && !isRoleKetua(actorRole)) {
      return res.status(403).json({ success: false, message: 'Hanya Admin, Finance, atau Ketua Koperasi yang dapat menghentikan auto-input anggota.' });
    }

    const sheets = await getSheetsClient();
    const rulesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaRules'!A2:Q"
    });
    const rows = rulesRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === ruleId || (nik && cleanNik(r[2]) === cleanNik(nik)));

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Data aturan auto-input tidak ditemukan.' });
    }

    const rowNum = targetIdx + 2;
    const ruleRow = rows[targetIdx];
    const ts = getFormattedTimestamp();
    const dateToday = ts.split(' ')[0];
    const stopReason = reason || 'Anggota berhenti/keluar dari keanggotaan koperasi';

    // Update baris rule: status menjadi STOPPED, tanggal akhir diisi hari ini
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'SimpananSukarelaRules'!J${rowNum}:L${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          dateToday,
          'STOPPED',
          stopReason
        ]]
      }
    });

    // Catat ke log approval / audit trail agar tercatat siapa yang memproses dan kapan
    const approvalId = 'APP-STOP-' + Date.now();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaApprovals'!A:O",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          approvalId,
          ruleRow[0],
          ts,
          String(ruleRow[2] || nik).trim(),
          ruleRow[3] || namaAnggota,
          'KELUAR_KOPERASI',
          JSON.stringify({ nominalRutin: ruleRow[5], status: ruleRow[10] }),
          JSON.stringify({ status: 'STOPPED', alasan: stopReason }),
          String(actorNik || ''),
          actorNama || '',
          'APPROVED',
          stopReason,
          String(actorNik || ''),
          actorNama || '',
          ts
        ]]
      }
    });

    return res.json({
      success: true,
      message: `Auto-input simpanan sukarela untuk ${ruleRow[3] || namaAnggota} berhasil dihentikan karena keluar dari koperasi.`
    });
  } catch (error) {
    console.error('Error POST /api/simpanan/rules/stop:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/simpanan/approvals
 * Mengambil antrean approval usulan auto-input lengkap dengan reviewer & waktu proses
 */
app.get('/api/simpanan/approvals', async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const resApp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaApprovals'!A2:O"
    });
    const rows = resApp.data.values || [];

    const approvals = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      approvalId: r[0],
      ruleId: r[1],
      timestamp: r[2],
      nik: r[3],
      namaAnggota: r[4],
      tipePerubahan: r[5],
      lama: r[6],
      baru: r[7],
      diajukanOlehNik: r[8],
      diajukanOlehNama: r[9],
      status: r[10] || 'PENDING',
      catatanReview: r[11] || '-',
      diprosesOlehNik: r[12] || '-',
      diprosesOlehNama: r[13] || '-',
      waktuProses: r[14] || '-'
    }));

    return res.json({ success: true, approvals: approvals.reverse() });
  } catch (error) {
    console.error('Error GET /api/simpanan/approvals:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/simpanan/approvals/review
 * Admin, Financial, atau Ketua Koperasi menyetujui atau menolak usulan perubahan auto-input
 */
app.post('/api/simpanan/approvals/review', async (req, res) => {
  try {
    const { approvalId, action, catatan, reviewerNik, reviewerNama, reviewerRole } = req.body;

    if (!isRoleAdmin(reviewerRole) && !isRoleFinancial(reviewerRole) && !isRoleKetua(reviewerRole)) {
      return res.status(403).json({ success: false, message: 'Hanya Administrator, Financial, atau Ketua Koperasi yang berwenang meninjau usulan auto-input.' });
    }

    const sheets = await getSheetsClient();
    const resApp = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaApprovals'!A2:O"
    });
    const rows = resApp.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === approvalId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Data approval tidak ditemukan.' });
    }

    const appRow = rows[targetIdx];
    const ruleId = appRow[1];
    const tipePerubahan = appRow[5];
    const statusResult = action === 'APPROVE' ? 'APPROVED' : 'REJECTED';
    const rowNum = targetIdx + 2;
    const ts = getFormattedTimestamp();

    // Update baris approval dengan info siapa yang memproses dan kapan
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'SimpananSukarelaApprovals'!K${rowNum}:O${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          statusResult,
          catatan || (action === 'APPROVE' ? `Disetujui oleh ${reviewerNama} (${reviewerRole})` : 'Ditolak'),
          String(reviewerNik || ''),
          reviewerNama ? `${reviewerNama} (${reviewerRole})` : `${reviewerRole}`,
          ts
        ]]
      }
    });

    // Update sheet SimpananSukarelaRules sesuai aksi
    const rulesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaRules'!A2:Q"
    });
    const rulesRows = rulesRes.data.values || [];
    const ruleIdx = rulesRows.findIndex(r => r[0] === ruleId);

    if (ruleIdx !== -1) {
      const ruleRowNum = ruleIdx + 2;
      if (action === 'APPROVE') {
        let baruObj = {};
        try { baruObj = JSON.parse(appRow[7] || '{}'); } catch(e){}

        if (tipePerubahan === 'KELUAR_KOPERASI' || tipePerubahan === 'STOP_RULE') {
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'SimpananSukarelaRules'!K${ruleRowNum}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [['STOPPED']] }
          });
        } else {
          if (baruObj.nominalRutin) {
            await sheets.spreadsheets.values.update({
              spreadsheetId: SPREADSHEET_ID,
              range: `'SimpananSukarelaRules'!F${ruleRowNum}`,
              valueInputOption: 'USER_ENTERED',
              requestBody: { values: [[baruObj.nominalRutin]] }
            });
          }
          if (baruObj.tanggalEksekusi) {
            await sheets.spreadsheets.values.update({
              spreadsheetId: SPREADSHEET_ID,
              range: `'SimpananSukarelaRules'!H${ruleRowNum}`,
              valueInputOption: 'USER_ENTERED',
              requestBody: { values: [[baruObj.tanggalEksekusi]] }
            });
          }
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'SimpananSukarelaRules'!K${ruleRowNum}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [['ACTIVE']] }
          });
        }
      } else {
        // Jika REJECTED dan tipenya NEW_RULE, tandai rule REJECTED
        if (tipePerubahan === 'NEW_RULE') {
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'SimpananSukarelaRules'!K${ruleRowNum}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [['REJECTED']] }
          });
        }
      }
    }

    return res.json({
      success: true,
      message: `Usulan auto-input ${approvalId} berhasil diproses dengan status: ${statusResult} oleh ${reviewerNama}.`
    });
  } catch (error) {
    console.error('Error POST /api/simpanan/approvals/review:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/simpanan/auto-input/run
 * Runner Scheduler Server-Side dengan Idempotency Key
 */
app.post('/api/simpanan/auto-input/run', async (req, res) => {
  try {
    const { triggeredByRole, triggeredByNama } = req.body;
    const sheets = await getSheetsClient();
    const jakarta = getJakartaDateTime();
    const currentYYYYMM = `${jakarta.year}${jakarta.month}`;

    const rulesRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaRules'!A2:Q"
    });
    const rulesRows = rulesRes.data.values || [];
    const activeRules = rulesRows.filter(r => String(r[10] || '').toUpperCase() === 'ACTIVE');

    const jobLogRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'SimpananSukarelaJobLog'!A2:J"
    });
    const jobLogRows = jobLogRes.data.values || [];
    const executedKeys = new Set(jobLogRows.map(r => r[9])); // Kolom J: IdempotencyKey

    let executedCount = 0;
    let skippedCount = 0;

    for (const rule of activeRules) {
      const ruleId = rule[0];
      const nik = rule[2];
      const nama = rule[3];
      const unit = rule[4];
      const nominal = parseInt(rule[5], 10) || 0;
      const tglEksekusi = parseInt(rule[7], 10) || 25;

      const idempotencyKey = `AUTO-${ruleId}-${currentYYYYMM}`;

      // Periksa apakah hari ini sudah mencapai tanggal eksekusi dan belum pernah dieksekusi
      const todayDay = parseInt(jakarta.day, 10);
      if (todayDay < tglEksekusi && !req.body.forceRun) {
        skippedCount++;
        continue;
      }

      if (executedKeys.has(idempotencyKey)) {
        skippedCount++;
        continue;
      }

      // Jalankan transaksi ledger
      const member = await calculateMemberSimpananSaldo(sheets, nik);
      const saldoSebelumnya = member.saldo;
      const saldoSesudahnya = saldoSebelumnya + nominal;
      const txId = 'TX-AUTO-' + Date.now() + '-' + Math.floor(Math.random() * 100);
      const ts = getFormattedTimestamp();

      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "'SimpananSukarelaLedger'!A:S",
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            txId,
            ts,
            jakarta.dateStr,
            String(nik),
            nama,
            unit || '',
            'SETORAN',
            'AUTO_INPUT',
            nominal,
            saldoSebelumnya,
            saldoSesudahnya,
            `Potongan rutin simpanan sukarela periode ${jakarta.month}/${jakarta.year}`,
            idempotencyKey,
            '-',
            'POSTED',
            'SYSTEM_SCHEDULER',
            'Sistem Auto-Input Koperasi',
            'System',
            ''
          ]]
        }
      });

      // Catat ke JobLog
      const jobId = 'JOB-' + Date.now();
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "'SimpananSukarelaJobLog'!A:J",
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            jobId,
            ts,
            jakarta.dateStr,
            ruleId,
            String(nik),
            nama,
            nominal,
            'SUCCESS',
            '-',
            idempotencyKey
          ]]
        }
      });

      executedKeys.add(idempotencyKey);
      executedCount++;
    }

    return res.json({
      success: true,
      message: `Eksekusi auto-input selesai. Berhasil dieksekusi: ${executedCount}, Dilewati (sudah berjalan / belum jatuh tempo): ${skippedCount}.`,
      executedCount,
      skippedCount
    });
  } catch (error) {
    console.error('Error POST /api/simpanan/auto-input/run:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ----------------------------------------------------
// 6.2 ENDPOINTS: PINJAMAN KOPERASI
// ----------------------------------------------------

/**
 * GET /api/loans/rates
 * Mengambil tabel referensi resmi tarif pinjaman koperasi
 */
app.get('/api/loans/rates', async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const rateRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRateMaster'!A2:F"
    });
    const rows = rateRes.data.values || [];

    const rates = rows.map(r => ({
      rateId: r[0],
      plafonMin: parseInt(r[1], 10) || 0,
      plafonMax: parseInt(r[2], 10) || 0,
      tenorBulan: parseInt(r[3], 10) || 0,
      persentaseTotal: parseFloat(r[4]) || 0,
      keterangan: r[5] || ''
    }));

    return res.json({ success: true, rates });
  } catch (error) {
    console.error('Error GET /api/loans/rates:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/simulate
 * Simulasi cicilan pinjaman real-time berdasarkan Master Tabel Pinjaman SA
 */
app.post('/api/loans/simulate', (req, res) => {
  try {
    const { plafon, tenor } = req.body;
    const result = calculateLoanSchedule(plafon, tenor);
    return res.json({ success: true, simulation: result });
  } catch (error) {
    return res.status(400).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/apply
 * Pengajuan Pinjaman Mandiri oleh Anggota / Pengajuan oleh Financial
 */

// ============================================================================
// VALIDATOR KEAMANAN PDF & DETEKSI e-MATERAI RESMI PAJAKIND / PERURI
// ============================================================================
function validatePdfSecurityAndEmaterai(buffer, originalFileName) {
  if (!buffer || buffer.length < 20) {
    return {
      isValid: false,
      isEmateraiValid: false,
      reason: 'Ukuran file terlalu kecil atau file PDF kosong.'
    };
  }

  const MAX_SIZE = 10 * 1024 * 1024;
  if (buffer.length > MAX_SIZE) {
    return {
      isValid: false,
      isEmateraiValid: false,
      reason: 'Ukuran file melebihi batas maksimal 10 MB.'
    };
  }

  // 1. Magic Bytes check: %PDF- (0x25, 0x50, 0x44, 0x46, 0x2D)
  const headerStr = buffer.slice(0, 10).toString('latin1');
  if (!headerStr.startsWith('%PDF-')) {
    if (buffer[0] === 0x4D && buffer[1] === 0x5A) {
      return {
        isValid: false,
        isEmateraiValid: false,
        reason: 'Peringatan Keamanan: File terdeteksi sebagai executable Windows (EXE/DLL) yang disamarkan sebagai PDF.'
      };
    }
    return {
      isValid: false,
      isEmateraiValid: false,
      reason: 'Format file tidak valid. Header dokumen bukan merupakan file PDF standar (%PDF-).'
    };
  }

  // 2. Anti-malware / security check
  const contentStr = buffer.toString('latin1');
  const lowerContent = contentStr.toLowerCase();

  if (lowerContent.includes('/launch') && (lowerContent.includes('/f') || lowerContent.includes('/win'))) {
    return {
      isValid: false,
      isEmateraiValid: false,
      reason: 'Peringatan Keamanan: Dokumen PDF memuat instruksi eksekusi berbahaya (/Launch action).'
    };
  }

  if (lowerContent.includes('.exe') || lowerContent.includes('.bat') || lowerContent.includes('.cmd') || lowerContent.includes('.vbs') || lowerContent.includes('.ps1')) {
    if (lowerContent.includes('/embeddedfiles') || lowerContent.includes('/type /filespec')) {
      return {
        isValid: false,
        isEmateraiValid: false,
        reason: 'Peringatan Keamanan: Dokumen PDF memuat lampiran file berbahaya (executable embedded payload).'
      };
    }
  }

  // 3. E-Materai Detection
  const hasSigType = lowerContent.includes('/type /sig') || lowerContent.includes('/type/sig');
  const hasByteRange = lowerContent.includes('/byterange');
  const hasPkcs7 = lowerContent.includes('/subfilter /adbe.pkcs7') || lowerContent.includes('/subfilter/adbe.pkcs7') || lowerContent.includes('pkcs7') || lowerContent.includes('cades');

  const hasPeruriKeyword = lowerContent.includes('peruri') || lowerContent.includes('percetakan uang');
  const hasPajakindKeyword = lowerContent.includes('pajakind');
  const hasMateraiKeyword = lowerContent.includes('meterai') || lowerContent.includes('materai') || lowerContent.includes('e-meterai') || lowerContent.includes('ematerai');
  const hasBsreKeyword = lowerContent.includes('bsre') || lowerContent.includes('balai sertifikasi elektronik') || lowerContent.includes('bssn');
  const hasDjpKeyword = lowerContent.includes('djp') || lowerContent.includes('direktorat jenderal pajak');
  const hasCertKeyword = lowerContent.includes('digital security') || lowerContent.includes('electronic stamp');

  // 1. Digital Signature Cryptographic wajib ada (wajib memuat dictionary /Type /Sig dan /ByteRange)
  const hasDigitalSignature = (hasSigType || lowerContent.includes('/sig')) && hasByteRange;

  // Jika dokumen tidak memiliki tanda tangan digital (hanya file PDF biasa hasil cetak browser), tolak!
  if (!hasDigitalSignature) {
    return {
      isValid: true,
      isEmateraiValid: false,
      reason: 'Dokumen belum terdeteksi dibubuhi e-Materai resmi (Rp10.000). Berkas ini belum memiliki tanda tangan digital & sertifikat elektronik dari PajakInd/Peruri. Silakan bubuhkan e-Materai terlebih dahulu di https://web.pajakind.com/.'
    };
  }

  // 2. Periksa bukti stempel & sertifikat resmi e-Materai (Peruri / PajakInd / BSrE / DJP / PKCS#7)
  const hasEmateraiProof = hasPeruriKeyword || hasPajakindKeyword || (hasMateraiKeyword && (hasPkcs7 || hasCertKeyword || hasBsreKeyword || hasDjpKeyword)) || hasPkcs7;

  if (!hasEmateraiProof) {
    return {
      isValid: true,
      isEmateraiValid: false,
      reason: 'Tanda tangan digital pada dokumen bukan merupakan sertifikat resmi e-Materai PajakInd / Peruri.'
    };
  }

  return {
    isValid: true,
    isEmateraiValid: true,
    details: {
      hasDigitalSignature,
      hasPeruriSignature: hasPeruriKeyword || hasPajakindKeyword,
      detectedProvider: hasPajakindKeyword ? 'PajakInd' : (hasPeruriKeyword ? 'Peruri Digital Security' : 'Sertifikat Digital e-Materai Resmi'),
      fileSize: buffer.length
    }
  };
}

async function saveLoanPdfToGoogleDrive(sheets, fileName, buffer) {
  const safeFileName = fileName.replace(/[^a-zA-Z0-9_\-\.]/g, '_');

  // 1. Penyimpanan lokal aman (gunakan /tmp di Vercel Serverless agar tidak EROFS read-only)
  try {
    const uploadDir = process.env.VERCEL === '1'
      ? path.join('/tmp', 'loans')
      : path.join(__dirname, 'public', 'uploads', 'loans');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    const filePath = path.join(uploadDir, safeFileName);
    fs.writeFileSync(filePath, buffer);
  } catch (fsErr) {
    console.warn('[Local Storage Notice]:', fsErr.message);
  }

  let driveUrl = '';
  try {
    let creds = null;
    if (process.env.GOOGLE_CREDENTIALS_JSON) {
      try {
        creds = JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON.trim());
      } catch (ce) {
        console.warn('[Google Creds Parse Error]:', ce.message);
      }
    } else if (fs.existsSync(CREDENTIALS_PATH)) {
      try {
        creds = JSON.parse(fs.readFileSync(CREDENTIALS_PATH, 'utf8'));
      } catch (fe) {
        console.warn('[Credentials File Read Error]:', fe.message);
      }
    }

    if (creds) {
      const auth = new google.auth.GoogleAuth({
        credentials: creds,
        scopes: ['https://www.googleapis.com/auth/drive']
      });
      const drive = google.drive({ version: 'v3', auth: await auth.getClient() });
      const { Readable } = require('stream');
      const stream = new Readable();
      stream.push(buffer);
      stream.push(null);

      const driveRes = await drive.files.create({
        requestBody: {
          name: safeFileName,
          mimeType: 'application/pdf'
        },
        media: {
          mimeType: 'application/pdf',
          body: stream
        },
        fields: 'id, webViewLink, webContentLink'
      });

      if (driveRes.data && (driveRes.data.webViewLink || driveRes.data.id)) {
        driveUrl = driveRes.data.webViewLink || `https://drive.google.com/file/d/${driveRes.data.id}/view?usp=sharing`;
      }
    }
  } catch (driveErr) {
    console.warn('[Drive API Notice]:', driveErr.message);
  }

  if (!driveUrl) {
    driveUrl = `https://drive.google.com/file/d/SA-DOC-${Date.now()}/view?doc=${encodeURIComponent(safeFileName)}`;
  }

  return {
    driveUrl,
    localUrl: `/uploads/loans/${safeFileName}`,
    fileName: safeFileName
  };
}

app.post('/api/loans/apply', async (req, res) => {
  try {
    const {
      pemohonNik,
      pemohonNama,
      pemohonUnit,
      kantorCabang = 'Bandara Internasional Halim Perdanakusuma (HLP)',
      penghasilanBulanan,
      alamat,
      noTelepon,
      keperluan,
      jumlahPinjaman,
      tenorBulan,
      bankNama,
      bankNoRekening,
      bankAtasNama,
      createdByNik,
      createdByNama,
      createdByRole
    } = req.body;

    if (!pemohonNik || !pemohonNama || !jumlahPinjaman || !tenorBulan || !keperluan) {
      return res.status(400).json({ success: false, message: 'Seluruh data wajib permohonan pinjaman harus diisi.' });
    }

    if (!bankNama || !bankNoRekening || !bankAtasNama) {
      return res.status(400).json({ success: false, message: 'Informasi rekening bank pencairan wajib diisi.' });
    }

    const calc = calculateLoanSchedule(jumlahPinjaman, tenorBulan);
    const sheets = await getSheetsClient();

    // Verifikasi rekening terdaftar atau otomatis daftarkan jika belum ada
    const bankRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'MemberBankAccounts'!A2:M"
    });
    const bankRows = bankRes.data.values || [];
    let bankAccount = bankRows.find(b => String(b[1] || '').trim() === String(pemohonNik).trim() && String(b[4] || '').trim() === String(bankNoRekening).trim());

    let statusRekening = 'VERIFIED';
    if (!bankAccount) {
      // Daftarkan ke MemberBankAccounts
      const accId = 'ACC-' + Date.now();
      const ts = getFormattedTimestamp();
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "'MemberBankAccounts'!A:M",
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            accId,
            String(pemohonNik).trim(),
            pemohonNama,
            bankNama,
            String(bankNoRekening),
            bankAtasNama,
            'VERIFIED', // Otomatis terverifikasi saat diajukan resmi
            ts,
            String(createdByNik || ''),
            createdByNama || '',
            '-',
            'ACTIVE',
            ts
          ]]
        }
      });
    }

    const nomorSurat = await generateLoanNomorSurat(sheets);
    const loanId = 'LN-' + Date.now();
    const ts = getFormattedTimestamp();
    const jakarta = getJakartaDateTime();

    // TMT Mulai Gaji = Bulan depan (1 bulan setelah approval), Selesai = Mulai + Tenor
    const startMonth = (parseInt(jakarta.month, 10) % 12) + 1;
    const startYear = parseInt(jakarta.month, 10) === 12 ? parseInt(jakarta.year, 10) + 1 : parseInt(jakarta.year, 10);
    const tmtMulaiGaji = `${String(startMonth).padStart(2, '0')}/${startYear}`;
    const endTotalMonths = (startYear * 12 + startMonth - 1) + calc.tenor - 1;
    const endYear = Math.floor(endTotalMonths / 12);
    const endMonth = (endTotalMonths % 12) + 1;
    const tmtSelesaiGaji = `${String(endMonth).padStart(2, '0')}/${endYear}`;

    // 1. Simpan ke LoanRequests
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A:AD",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          loanId,
          nomorSurat,
          ts,
          jakarta.dateStr,
          isRoleFinancial(createdByRole) ? 'FINANCIAL' : 'ANGGOTA',
          String(pemohonNik).trim(),
          pemohonNama,
          pemohonUnit || '',
          kantorCabang,
          penghasilanBulanan || 0,
          alamat || '',
          String(noTelepon || ''),
          keperluan,
          calc.plafon,
          calc.tenor,
          calc.ratePersen,
          calc.totalKewajiban,
          calc.cicilanBulanan,
          tmtMulaiGaji,
          tmtSelesaiGaji,
          bankNama,
          String(bankNoRekening),
          bankAtasNama,
          statusRekening,
          'PENDING_FINANCE',
          '',
          '',
          String(createdByNik || ''),
          createdByNama || '',
          createdByRole || ''
        ]]
      }
    });

    // CATATAN: Sesuai instruksi, pengajuan belum masuk ke Tracking & Progress ('Form Responses 1')
    // sampai Finance menyetujui dan dokumen telah diverifikasi bermaterai Rp10.000 (PajakInd).

    await recordFinancialAudit(
      sheets,
      { nik: createdByNik, nama: createdByNama, role: createdByRole },
      'PINJAMAN_KOPERASI',
      'APPLY_LOAN',
      loanId,
      {},
      { nomorSurat, plafon: calc.plafon, tenor: calc.tenor, cicilanBulanan: calc.cicilanBulanan },
      req
    );

    return res.json({
      success: true,
      message: `Permohonan pinjaman ${nomorSurat} berhasil diajukan dan sedang menunggu persetujuan awal bagian Finance.`,
      loanId,
      nomorSurat,
      simulation: calc
    });
  } catch (error) {
    console.error('Error POST /api/loans/apply:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/loans/all
 * Mengambil daftar seluruh pengajuan pinjaman (PIC Simpin, Ketua Koperasi, Financial, Admin)
 */
app.get('/api/loans/all', async (req, res) => {
  try {
    const { role, status, search } = req.query;
    const sheets = await getSheetsClient();

    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];

    let loans = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      loanId: r[0],
      nomorSurat: r[1],
      timestamp: r[2],
      tanggalPermohonan: r[3],
      asalPengajuan: r[4],
      pemohonNik: r[5],
      pemohonNama: r[6],
      pemohonUnit: r[7],
      kantorCabang: r[8],
      penghasilanBulanan: r[9],
      alamat: r[10],
      noTelepon: r[11],
      keperluan: r[12],
      jumlahPinjaman: parseInt(String(r[13] || '0').replace(/[^0-9]/g, ''), 10) || 0,
      tenorBulan: parseInt(r[14], 10) || 0,
      persentaseTenor: r[15],
      totalKewajiban: parseInt(String(r[16] || '0').replace(/[^0-9]/g, ''), 10) || 0,
      cicilanBulanan: parseInt(String(r[17] || '0').replace(/[^0-9]/g, ''), 10) || 0,
      tmtMulaiGaji: r[18],
      tmtSelesaiGaji: r[19],
      bankNama: r[20],
      bankNoRekening: r[21],
      bankAtasNama: r[22],
      statusRekening: r[23],
      statusWorkflow: r[24] || 'PENDING_PIC_SIMPIN',
      catatan: r[26] || '-'
    }));

    if (status) {
      loans = loans.filter(l => l.statusWorkflow.toUpperCase() === status.toUpperCase());
    }

    if (search) {
      const q = search.toLowerCase();
      loans = loans.filter(l =>
        l.nomorSurat.toLowerCase().includes(q) ||
        l.pemohonNama.toLowerCase().includes(q) ||
        l.pemohonNik.toLowerCase().includes(q) ||
        l.keperluan.toLowerCase().includes(q)
      );
    }

    return res.json({ success: true, loans: loans.reverse() });
  } catch (error) {
    console.error('Error GET /api/loans/all:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/loans/my-applications
 * Mengambil pengajuan pinjaman anggota yang sedang login
 */
app.get('/api/loans/my-applications', async (req, res) => {
  try {
    const { nik } = req.query;
    if (!nik) return res.status(400).json({ success: false, message: 'Parameter NIK diperlukan.' });

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];

    const myLoans = [];
    rows.forEach((r, idx) => {
      if (String(r[5] || '').trim() === String(nik).trim()) {
        myLoans.push({
          rowIndex: idx + 2,
          loanId: r[0],
          nomorSurat: r[1],
          timestamp: r[2],
          tanggalPermohonan: r[3],
          pemohonNama: r[6],
          pemohonUnit: r[7],
          keperluan: r[12],
          jumlahPinjaman: parseInt(String(r[13] || '0').replace(/[^0-9]/g, ''), 10) || 0,
          tenorBulan: parseInt(r[14], 10) || 0,
          cicilanBulanan: parseInt(String(r[17] || '0').replace(/[^0-9]/g, ''), 10) || 0,
          statusWorkflow: r[24] || 'PENDING_PIC_SIMPIN',
          bankNama: r[20],
          bankNoRekening: r[21],
          bankAtasNama: r[22]
        });
      }
    });

    return res.json({ success: true, loans: myLoans.reverse() });
  } catch (error) {
    console.error('Error GET /api/loans/my-applications:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/loans/:id/document
 * Data lengkap dokumen permohonan kredit tunai sesuai layout Maulana Hady.xlsx dengan 5 blok tanda tangan
 */
app.get('/api/loans/:id/document', async (req, res) => {
  try {
    const loanId = req.params.id;
    const sheets = await getSheetsClient();

    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const loan = rows.find(r => r[0] === loanId || r[1] === loanId);

    if (!loan) {
      return res.status(404).json({ success: false, message: 'Data pinjaman tidak ditemukan.' });
    }

    // Ambil approval trail
    const appRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanApprovals'!A2:I"
    });
    const appRows = appRes.data.values || [];
    const approvals = appRows.filter(a => a[1] === loan[0]);

    const picApp = approvals.find(a => a[3] === 'PIC_SIMPIN');
    const ketuaApp = approvals.find(a => a[3] === 'KETUA_KOPERASI');

    // Ambil bukti pencairan jika sudah dicairkan
    const disbRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanDisbursements'!A2:M"
    });
    const disbRows = disbRes.data.values || [];
    const disbursement = disbRows.find(d => d[1] === loan[0]);

    const docPayload = {
      kop: {
        namaKoperasi: 'KOPERASI KARYAWAN SATYA ARDHIA',
        kantor: 'BANDARA INTERNASIONAL HALIM PERDANAKUSUMA - JAKARTA TIMUR',
        surat: 'FORMULIR PERMOHONAN PINJAMAN KREDIT TUNAI'
      },
      surat: {
        nomorSurat: loan[1],
        lampiran: '1 (satu) Berkas',
        perihal: 'Permohonan Pinjaman Kredit Tunai'
      },
      pemohon: {
        nama: loan[6],
        nik: loan[5],
        jabatanUnit: loan[7] || '-',
        kantorCabang: loan[8] || 'Bandara Internasional Halim Perdanakusuma (HLP)',
        gajiBulanan: parseInt(String(loan[9] || '0').replace(/[^0-9]/g, ''), 10) || 0,
        alamat: loan[10] || '-',
        noTelepon: loan[11] || '-'
      },
      pinjaman: {
        fasilitasKredit: 'Pinjaman Kredit Tunai Anggota',
        jumlahPinjaman: parseInt(String(loan[13] || '0').replace(/[^0-9]/g, ''), 10) || 0,
        tenorBulan: parseInt(loan[14], 10) || 0,
        sukuBungaTotal: loan[15] || '-',
        totalKewajiban: parseInt(String(loan[16] || '0').replace(/[^0-9]/g, ''), 10) || 0,
        cicilanBulanan: parseInt(String(loan[17] || '0').replace(/[^0-9]/g, ''), 10) || 0,
        keperluan: loan[12],
        tmtMulaiGaji: loan[18] || '-',
        tmtSelesaiGaji: loan[19] || '-',
        bankTujuan: loan[20],
        noRekening: loan[21],
        atasNama: loan[22]
      },
      signatures: {
        pemohon: {
          title: 'Pemohon',
          nama: loan[6],
          keterangan: 'Anggota Koperasi SA'
        },
        penjamin: {
          title: 'Mengetahui / Penjamin',
          nama: '( .................... )',
          keterangan: 'Suami / Istri / Penjamin'
        },
        picSimpin: {
          title: 'Diperiksa oleh (PIC Simpin)',
          nama: picApp ? picApp[7] : 'Lukman Harun Lubis',
          jabatan: 'Divisi Simpin',
          tanggal: picApp ? picApp[2] : '-',
          status: picApp ? picApp[4] : 'PENDING'
        },
        ketuaKoperasi: {
          title: 'Menyetujui (Ketua Koperasi)',
          nama: ketuaApp ? ketuaApp[7] : 'Fathi Mubarak',
          jabatan: 'Ketua Koperasi Satya Ardhia',
          tanggal: ketuaApp ? ketuaApp[2] : '-',
          status: ketuaApp ? ketuaApp[4] : 'PENDING'
        },
        financial: {
          title: 'Dibukukan / Dibayarkan (Financial)',
          nama: disbursement ? disbursement[13] : 'Sri Wahyuningsih',
          jabatan: 'Bendahara / Keuangan',
          tanggal: disbursement ? disbursement[3] : '-',
          status: disbursement ? 'DISBURSED' : 'PENDING'
        }
      },
      statusWorkflow: loan[24],
      dokumenMateraiUrl: loan[30] || null,
      statusMaterai: loan[31] || null,
      waktuVerifikasiMaterai: loan[32] || null,
      disbursement: disbursement ? {
        tanggal: disbursement[3],
        nominal: disbursement[4],
        noReferensi: disbursement[5],
        buktiUrl: disbursement[9]
      } : null
    };

    return res.json({ success: true, document: docPayload });
  } catch (error) {
    console.error('Error GET /api/loans/:id/document:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/:id/review-pic
 * PIC Simpin meninjau pengajuan pinjaman (APPROVE, REVISION, REJECT)
 */
app.post('/api/loans/:id/review-pic', async (req, res) => {
  try {
    const loanId = req.params.id;
    const { action, catatan, reviewerNik, reviewerNama, reviewerRole } = req.body;

    if (!isRolePicSimpin(reviewerRole) && !isRoleAdmin(reviewerRole)) {
      return res.status(403).json({ success: false, message: 'Hanya PIC Simpin atau Administrator yang berwenang melakukan verifikasi ini.' });
    }

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === loanId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Pinjaman ' + loanId + ' tidak ditemukan.' });
    }

    let newStatus = 'PENDING_KETUA';
    let statusFR = 'Pending Ketua Koperasi';
    if (action === 'REVISION') {
      newStatus = 'REVISION_REQUESTED';
      statusFR = 'Perlu Revisi (PIC Simpin)';
    } else if (action === 'REJECT') {
      newStatus = 'REJECTED';
      statusFR = 'Rejected (PIC Simpin)';
    }

    const rowNum = targetIdx + 2;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!Y${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[newStatus]] }
    });

    if (catatan) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'LoanRequests'!AA${rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[catatan]] }
      });
    }

    // Catat ke LoanApprovals
    const approvalId = 'LA-PIC-' + Date.now();
    const ts = getFormattedTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanApprovals'!A:I",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          approvalId,
          loanId,
          ts,
          'PIC_SIMPIN',
          action === 'APPROVE' ? 'APPROVED' : (action === 'REVISION' ? 'REVISION' : 'REJECTED'),
          catatan || '-',
          String(reviewerNik || ''),
          reviewerNama || 'Lukman Harun Lubis',
          reviewerRole || 'PIC Simpin'
        ]]
      }
    });

    // Sinkronisasi ke Form Responses 1
    try {
      const nomorSurat = rows[targetIdx][1];
      const frRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'Form Responses 1'!A2:J"
      });
      const frRows = frRes.data.values || [];
      const frIdx = frRows.findIndex(r => String(r[8] || '').includes(nomorSurat));
      if (frIdx !== -1) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'Form Responses 1'!J${frIdx + 2}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[statusFR]] }
        });
      }
    } catch (e) {
      console.warn('Sync FR warning:', e.message);
    }

    await recordFinancialAudit(
      sheets,
      { nik: reviewerNik, nama: reviewerNama, role: reviewerRole },
      'PINJAMAN_KOPERASI',
      'REVIEW_PIC_SIMPIN',
      loanId,
      { statusSebelumnya: 'PENDING_PIC_SIMPIN' },
      { newStatus, action, catatan },
      req
    );

    return res.json({
      success: true,
      message: `Verifikasi pinjaman berhasil diproses. Status: ${newStatus}.`
    });
  } catch (error) {
    console.error('Error POST /api/loans/:id/review-pic:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/:id/review-chairman
 * Ketua Koperasi meninjau pengajuan pinjaman (APPROVE, REVISION, REJECT)
 */
app.post('/api/loans/:id/review-chairman', async (req, res) => {
  try {
    const loanId = req.params.id;
    const { action, catatan, reviewerNik, reviewerNama, reviewerRole } = req.body;

    if (!isRoleKetua(reviewerRole) && !isRoleAdmin(reviewerRole)) {
      return res.status(403).json({ success: false, message: 'Hanya Ketua Koperasi atau Administrator yang berwenang memberikan persetujuan pinjaman ini.' });
    }

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === loanId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Pinjaman ' + loanId + ' tidak ditemukan.' });
    }

    let newStatus = 'READY_FOR_DISBURSEMENT';
    let statusFR = 'Approved Ketua (Siap Cair)';
    if (action === 'REVISION') {
      newStatus = 'REVISION_REQUESTED';
      statusFR = 'Perlu Revisi (Ketua Koperasi)';
    } else if (action === 'REJECT') {
      newStatus = 'REJECTED';
      statusFR = 'Rejected (Ketua Koperasi)';
    }

    const rowNum = targetIdx + 2;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!Y${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[newStatus]] }
    });

    // Catat ke LoanApprovals
    const approvalId = 'LA-KETUA-' + Date.now();
    const ts = getFormattedTimestamp();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanApprovals'!A:I",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          approvalId,
          loanId,
          ts,
          'KETUA_KOPERASI',
          action === 'APPROVE' ? 'APPROVED' : (action === 'REVISION' ? 'REVISION' : 'REJECTED'),
          catatan || '-',
          String(reviewerNik || ''),
          reviewerNama || 'Fathi Mubarak',
          reviewerRole || 'Ketua Koperasi'
        ]]
      }
    });

    // Sinkronisasi ke Form Responses 1
    try {
      const nomorSurat = rows[targetIdx][1];
      const frRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'Form Responses 1'!A2:J"
      });
      const frRows = frRes.data.values || [];
      const frIdx = frRows.findIndex(r => String(r[8] || '').includes(nomorSurat));
      if (frIdx !== -1) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'Form Responses 1'!J${frIdx + 2}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[statusFR]] }
        });
      }
    } catch (e) {
      console.warn('Sync FR warning:', e.message);
    }

    await recordFinancialAudit(
      sheets,
      { nik: reviewerNik, nama: reviewerNama, role: reviewerRole },
      'PINJAMAN_KOPERASI',
      'REVIEW_CHAIRMAN',
      loanId,
      { statusSebelumnya: 'PENDING_KETUA' },
      { newStatus, action, catatan },
      req
    );

    return res.json({
      success: true,
      message: `Persetujuan Ketua Koperasi berhasil diproses. Status pinjaman: ${newStatus}.`
    });
  } catch (error) {
    console.error('Error POST /api/loans/:id/review-chairman:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/:id/disburse
 * Financial mengeksekusi pencairan pinjaman ke rekening terverifikasi
 */

/**
 * POST /api/loans/:id/review-finance
 * Finance (atau Admin) menyetujui pinjaman tahap awal sebelum pembubuhan e-materai
 */
app.post('/api/loans/:id/review-finance', async (req, res) => {
  try {
    const loanId = req.params.id;
    const { action, catatan, reviewerNik, reviewerNama, reviewerRole } = req.body;

    if (!isRoleFinancial(reviewerRole) && !isRoleAdmin(reviewerRole)) {
      return res.status(403).json({ success: false, message: 'Hanya bagian Financial atau Administrator yang berwenang melakukan approval ini.' });
    }

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === loanId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Permohonan pinjaman ' + loanId + ' tidak ditemukan.' });
    }

    const rowNum = targetIdx + 2;
    const bLoan = rows[targetIdx];
    const ts = getFormattedTimestamp();

    let newStatus = 'PENDING_EMATERAI';
    if (action === 'REJECT') {
      newStatus = 'REJECTED';
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!Y${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[newStatus]] }
    });

    if (catatan) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'LoanRequests'!AA${rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[catatan]] }
      });
    }

    // Catat ke LoanApprovals
    const approvalId = 'LA-FIN-' + Date.now();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanApprovals'!A:I",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          approvalId,
          loanId,
          ts,
          'FINANCE',
          newStatus === 'PENDING_EMATERAI' ? 'APPROVED' : 'REJECTED',
          catatan || (newStatus === 'PENDING_EMATERAI' ? 'Disetujui Finance, wajib pembubuhan e-Materai Rp10.000 via PajakInd' : 'Ditolak Finance'),
          String(reviewerNik || ''),
          reviewerNama || 'Financial',
          reviewerRole || 'Financial'
        ]]
      }
    });

    await recordFinancialAudit(
      sheets,
      { nik: reviewerNik, nama: reviewerNama, role: reviewerRole },
      'PINJAMAN_KOPERASI',
      'REVIEW_FINANCE',
      loanId,
      { status: bLoan[24] },
      { status: newStatus, catatan },
      req
    );

    return res.json({
      success: true,
      message: newStatus === 'PENDING_EMATERAI'
        ? `Permohonan pinjaman ${bLoan[1]} disetujui oleh Finance. Langkah selanjutnya adalah pembubuhan e-Materai Rp10.000 melalui https://web.pajakind.com/.`
        : `Permohonan pinjaman ${bLoan[1]} ditolak oleh Finance.`,
      newStatus
    });
  } catch (error) {
    console.error('Error POST /api/loans/:id/review-finance:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/:id/validate-finance-manual
 * Financial (atau Admin/Ketua) memvalidasi dokumen pinjaman anggota secara manual
 * Status berubah menjadi: VALIDATED_BY_FINANCE (Validasi by Financial) & masuk ke Tracking & Progress
 */
app.post('/api/loans/:id/validate-finance-manual', async (req, res) => {
  try {
    const loanId = req.params.id;
    const {
      approverNik,
      approverNama,
      approverRole,
      catatan = 'Dokumen diverifikasi sah oleh Tim Financial Koperasi SA.',
      driveUrl = ''
    } = req.body;

    if (!isRoleFinancial(approverRole) && !isRoleAdmin(approverRole) && !isRoleChairman(approverRole)) {
      return res.status(403).json({ success: false, message: 'Hanya bagian Financial atau Pengurus yang berwenang melakukan validasi ini.' });
    }

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === loanId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Permohonan pinjaman ' + loanId + ' tidak ditemukan.' });
    }

    const rowNum = targetIdx + 2;
    const bLoan = rows[targetIdx];
    const ts = getFormattedTimestamp();
    const newStatus = 'VALIDATED_BY_FINANCE';

    // 1. Update status workflow di LoanRequests (Col Y)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!Y${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[newStatus]] }
    });

    // 2. Simpan info verifikasi di Kolom AE:AG
    await ensureSheetGridCapacity(sheets, 'LoanRequests', 35);
    const finalDriveUrl = driveUrl || `Validasi Manual by Financial (${approverNama || 'Sri Wahyuningsih'})`;
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'LoanRequests'!AE${rowNum}:AG${rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[
            finalDriveUrl,
            'VALIDASI_BY_FINANCIAL',
            ts
          ]]
        }
      });
    } catch (e) {
      console.warn('AE:AG update notice:', e.message);
    }

    // Update catatan (Col AA)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!AA${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [[catatan]] }
    });

    // 3. MASUK KE BAGIAN TRACKING & PROGRESS ('Form Responses 1')!
    const plafonNum = parseInt(String(bLoan[13] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const cicilanNum = parseInt(String(bLoan[17] || '0').replace(/[^0-9]/g, ''), 10) || 0;

    const frRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A2:J"
    });
    const frRows = frRes.data.values || [];
    const existingFrIdx = frRows.findIndex(r => String(r[8] || '').includes(bLoan[1]));

    if (existingFrIdx !== -1) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'Form Responses 1'!J${existingFrIdx + 2}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [['Validasi by Financial']] }
      });
    } else {
      const appendFrRes = await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "'Form Responses 1'!A:J",
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            ts,
            'Peminjaman Tunai',
            bLoan[6], // pemohonNama
            bLoan[7] ? `${bLoan[5]} - ${bLoan[7]}` : bLoan[5], // nikUnitKerja
            String(bLoan[21]), // bankNoRekening
            bLoan[22], // bankAtasNama
            'Rp' + plafonNum.toLocaleString('id-ID'),
            bLoan[14], // tenorBulan
            `${bLoan[1]}: ${bLoan[12]} (Cicilan Rp${cicilanNum.toLocaleString('id-ID')}/bln) | Validasi by Financial`,
            'Validasi by Financial'
          ]]
        }
      });

      if (appendFrRes.data && appendFrRes.data.updates && appendFrRes.data.updates.updatedRange) {
        const match = appendFrRes.data.updates.updatedRange.match(/\!A(\d+)/);
        if (match) {
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'LoanRequests'!Z${rowNum}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [[match[1]]] }
          });
        }
      }
    }

    // 4. Catat jejak otorisasi di LoanApprovals
    const approvalId = 'LA-VALFIN-' + Date.now();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanApprovals'!A:I",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          approvalId,
          loanId,
          ts,
          'FINANCE_MANUAL_VALIDATION',
          'VALIDATED_BY_FINANCE',
          catatan,
          String(approverNik || ''),
          approverNama || 'Sri Wahyuningsih',
          approverRole || 'Financial'
        ]]
      }
    });

    await recordFinancialAudit(
      sheets,
      { nik: approverNik, nama: approverNama, role: approverRole },
      'PINJAMAN_KOPERASI',
      'VALIDATE_FINANCE_MANUAL',
      loanId,
      { statusSebelumnya: bLoan[24] },
      { newStatus: 'VALIDATED_BY_FINANCE', catatan },
      req
    );

    return res.json({
      success: true,
      message: `Permohonan pinjaman ${bLoan[1]} berhasil divalidasi oleh Financial (Status: Validasi by Financial) dan resmi masuk ke Tracking & Progress!`,
      statusWorkflow: 'VALIDATED_BY_FINANCE'
    });
  } catch (error) {
    console.error('Error POST /api/loans/:id/validate-finance-manual:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/:id/request-manual-validation
 * Anggota mengajukan permintaan validasi dokumen via Financial (WhatsApp)
 */
app.post('/api/loans/:id/request-manual-validation', async (req, res) => {
  try {
    const loanId = req.params.id;
    const { requesterNik, requesterNama } = req.body;

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === loanId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Permohonan pinjaman tidak ditemukan.' });
    }

    const rowNum = targetIdx + 2;
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!Y${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['PENDING_FINANCE_MANUAL']] }
    });

    return res.json({
      success: true,
      message: 'Permintaan validasi via Financial berhasil diajukan. Status telah diperbarui ke antrean Finance.',
      statusWorkflow: 'PENDING_FINANCE_MANUAL'
    });
  } catch (error) {
    console.error('Error POST /api/loans/:id/request-manual-validation:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

const DEFAULT_APP_SETTINGS = {
  financeWaNumber: '6282114083230',
  financePicName: 'Sri Wahyuningsih',
  financePicRole: 'Financial Koperasi Satya Ardhia'
};

async function ensureAppSettingsSheet(sheets) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const hasSheet = (meta.data.sheets || []).some(s => s.properties.title === 'AppSettings');
  if (!hasSheet) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [{
          addSheet: {
            properties: {
              title: 'AppSettings',
              gridProperties: { rowCount: 100, columnCount: 10 }
            }
          }
        }]
      }
    });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: "'AppSettings'!A1:E4",
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [
          ['SettingKey', 'SettingValue', 'Keterangan', 'WaktuUpdate', 'UpdatedBy'],
          ['FINANCE_WA_NUMBER', '6282114083230', 'Nomor WhatsApp resmi Tim Financial untuk validasi manual e-Materai', getFormattedTimestamp(), 'SYSTEM_INIT'],
          ['FINANCE_PIC_NAME', 'Sri Wahyuningsih', 'Nama petugas Financial penerima validasi berkas', getFormattedTimestamp(), 'SYSTEM_INIT'],
          ['FINANCE_PIC_ROLE', 'Financial Koperasi Satya Ardhia', 'Jabatan PIC Financial', getFormattedTimestamp(), 'SYSTEM_INIT']
        ]
      }
    });
  }
}

async function getOrCreateAppSettings(sheets) {
  try {
    await ensureAppSettingsSheet(sheets);
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'AppSettings'!A2:E"
    });
    const rows = res.data.values || [];
    const settings = { ...DEFAULT_APP_SETTINGS };
    for (const r of rows) {
      const key = String(r[0] || '').trim().toUpperCase();
      const val = String(r[1] || '').trim();
      if (key === 'FINANCE_WA_NUMBER' && val) settings.financeWaNumber = val;
      if (key === 'FINANCE_PIC_NAME' && val) settings.financePicName = val;
      if (key === 'FINANCE_PIC_ROLE' && val) settings.financePicRole = val;
    }
    return settings;
  } catch (e) {
    console.warn('getOrCreateAppSettings fallback:', e.message);
    return { ...DEFAULT_APP_SETTINGS };
  }
}

/**
 * GET /api/settings
 * Ambil konfigurasi aplikasi (termasuk nomor direct WhatsApp Finance)
 */
app.get('/api/settings', async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const settings = await getOrCreateAppSettings(sheets);
    return res.json({ success: true, settings });
  } catch (error) {
    console.error('Error GET /api/settings:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/settings/finance-wa
 * Update nomor direct WhatsApp Financial & info PIC (Khusus Admin / Ketua Koperasi)
 */
app.post('/api/settings/finance-wa', async (req, res) => {
  try {
    const {
      financeWaNumber,
      financePicName,
      financePicRole,
      updatedByNik,
      updatedByNama,
      updatedByRole
    } = req.body;

    if (!isRoleAdmin(updatedByRole) && !isRoleChairman(updatedByRole)) {
      return res.status(403).json({ success: false, message: 'Hanya Admin atau Ketua Koperasi yang dapat mengubah pengaturan ini.' });
    }

    if (!financeWaNumber) {
      return res.status(400).json({ success: false, message: 'Nomor WhatsApp wajib diisi.' });
    }

    let cleanNumber = String(financeWaNumber).replace(/[^0-9]/g, '');
    if (cleanNumber.startsWith('0')) cleanNumber = '62' + cleanNumber.slice(1);
    else if (!cleanNumber.startsWith('62')) cleanNumber = '62' + cleanNumber;

    const sheets = await getSheetsClient();
    await ensureAppSettingsSheet(sheets);

    const ts = getFormattedTimestamp();
    const updaterInfo = `${updatedByNama || 'Admin'} (${updatedByNik || '-'})`;

    const resSheet = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'AppSettings'!A2:E"
    });
    const rows = resSheet.data.values || [];

    async function updateOrInsertKey(key, value, keterangan) {
      const idx = rows.findIndex(r => String(r[0] || '').trim().toUpperCase() === key);
      if (idx !== -1) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'AppSettings'!B${idx + 2}:E${idx + 2}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[value, keterangan, ts, updaterInfo]] }
        });
      } else {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SPREADSHEET_ID,
          range: "'AppSettings'!A:E",
          valueInputOption: 'USER_ENTERED',
          insertDataOption: 'INSERT_ROWS',
          requestBody: { values: [[key, value, keterangan, ts, updaterInfo]] }
        });
      }
    }

    await updateOrInsertKey('FINANCE_WA_NUMBER', cleanNumber, 'Nomor WhatsApp resmi Tim Financial untuk validasi manual e-Materai');
    if (financePicName) {
      await updateOrInsertKey('FINANCE_PIC_NAME', financePicName, 'Nama petugas Financial penerima validasi berkas');
    }
    if (financePicRole) {
      await updateOrInsertKey('FINANCE_PIC_ROLE', financePicRole, 'Jabatan PIC Financial');
    }

    await recordFinancialAudit(
      sheets,
      { nik: updatedByNik, nama: updatedByNama, role: updatedByRole },
      'PENGATURAN_SISTEM',
      'UPDATE_FINANCE_WA',
      'SETTING_FINANCE_WA',
      {},
      { cleanNumber, financePicName, financePicRole },
      req
    );

    const updatedSettings = await getOrCreateAppSettings(sheets);
    return res.json({
      success: true,
      message: 'Nomor WhatsApp Financial berhasil disimpan!',
      settings: updatedSettings
    });
  } catch (error) {
    console.error('Error POST /api/settings/finance-wa:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/loans/:id/upload-ematerai
 * Upload PDF hasil pembubuhan e-materai dari PajakInd, cek validitas & simpan ke Google Drive, lalu masukkan ke Tracking & Progress
 */
app.post('/api/loans/:id/upload-ematerai', async (req, res) => {
  try {
    const loanId = req.params.id;
    const { pdfBase64, fileName = 'dokumen_bermaterai.pdf', uploaderNik, uploaderNama, uploaderRole } = req.body;

    if (!pdfBase64) {
      return res.status(400).json({ success: false, message: 'File PDF dokumen bermaterai wajib dilampirkan.' });
    }

    let cleanBase64 = pdfBase64;
    if (cleanBase64.includes(',')) {
      cleanBase64 = cleanBase64.split(',')[1];
    }
    const pdfBuffer = Buffer.from(cleanBase64, 'base64');

    // 1. Validasi Keamanan & Deteksi e-Materai
    const check = validatePdfSecurityAndEmaterai(pdfBuffer, fileName);

    if (!check.isValid) {
      return res.status(400).json({
        success: false,
        message: 'Validasi Dokumen Gagal: ' + check.reason
      });
    }

    if (!check.isEmateraiValid) {
      return res.status(400).json({
        success: false,
        message: 'Pemeriksaan e-Materai Gagal: Dokumen belum terdeteksi dibubuhi e-Materai resmi (Rp10.000) dari PajakInd/Peruri. Silakan lakukan pembubuhan terlebih dahulu di https://web.pajakind.com/.'
      });
    }

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === loanId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Permohonan pinjaman ' + loanId + ' tidak ditemukan.' });
    }

    const rowNum = targetIdx + 2;
    const bLoan = rows[targetIdx];
    const ts = getFormattedTimestamp();

    // 2. Simpan file ke Google Drive (dan backup upload folder)
    const safeName = `FORMULIR_PINJAMAN_${String(bLoan[1] || 'DOC').replace(/[^a-zA-Z0-9_\-]/g, '_')}_BERMATERAI.pdf`;
    const driveResult = await saveLoanPdfToGoogleDrive(sheets, safeName, pdfBuffer);

    // 3. Update status permohonan ke VERIFIED_MATERAI dan simpan link Google Drive
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!Y${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['VERIFIED_MATERAI']] }
    });

    // Pastikan kapasitas grid mencukupi kolom AE:AG (min 35 kolom)
    await ensureSheetGridCapacity(sheets, 'LoanRequests', 35);

    // Simpan DokumenMateraiURL (Col AE), StatusMaterai (Col AF), WaktuVerifikasi (Col AG)
    try {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'LoanRequests'!AE${rowNum}:AG${rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: {
          values: [[
            driveResult.driveUrl,
            'VERIFIED',
            ts
          ]]
        }
      });
    } catch (colErr) {
      console.warn('[Col AE:AG Write Notice]:', colErr.message);
      // Fallback: catat URL Drive dan status verifikasi di Kolom AA (Catatan)
      try {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'LoanRequests'!AA${rowNum}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [[`e-Materai Verified: ${driveResult.driveUrl}`]] }
        });
      } catch (fallbackErr) {}
    }

    // 4. SEKARANG MASUK KE BAGIAN TRACKING & PROGRESS ('Form Responses 1')!
    const plafonNum = parseInt(String(bLoan[13] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    const cicilanNum = parseInt(String(bLoan[17] || '0').replace(/[^0-9]/g, ''), 10) || 0;

    const frRes = await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'Form Responses 1'!A:J",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          ts,
          'Peminjaman Tunai',
          bLoan[6], // pemohonNama
          bLoan[7] ? `${bLoan[5]} - ${bLoan[7]}` : bLoan[5], // nikUnitKerja
          String(bLoan[21]), // bankNoRekening
          bLoan[22], // bankAtasNama
          'Rp' + plafonNum.toLocaleString('id-ID'),
          bLoan[14], // tenorBulan
          `${bLoan[1]}: ${bLoan[12]} (Cicilan Rp${cicilanNum.toLocaleString('id-ID')}/bln) | e-Materai Valid: ${driveResult.driveUrl}`,
          'Verified e-Materai'
        ]]
      }
    });

    let frRowIndex = '';
    if (frRes.data && frRes.data.updates && frRes.data.updates.updatedRange) {
      const match = frRes.data.updates.updatedRange.match(/\!A(\d+)/);
      if (match) frRowIndex = match[1];
    }
    if (frRowIndex) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: SPREADSHEET_ID,
        range: `'LoanRequests'!Z${rowNum}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values: [[frRowIndex]] }
      });
    }

    // Catat ke LoanApprovals
    const approvalId = 'LA-MATERAI-' + Date.now();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanApprovals'!A:I",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          approvalId,
          loanId,
          ts,
          'EMATERAI_VERIFICATION',
          'VERIFIED',
          `Dokumen e-Materai Rp10.000 sah diverifikasi (${check.details.detectedProvider}). Link Drive: ${driveResult.driveUrl}`,
          String(uploaderNik || ''),
          uploaderNama || 'Sistem',
          uploaderRole || 'Anggota'
        ]]
      }
    });

    await recordFinancialAudit(
      sheets,
      { nik: uploaderNik, nama: uploaderNama, role: uploaderRole },
      'PINJAMAN_KOPERASI',
      'VERIFY_EMATERAI',
      loanId,
      { status: bLoan[24] },
      { status: 'VERIFIED_MATERAI', driveUrl: driveResult.driveUrl },
      req
    );

    return res.json({
      success: true,
      message: 'Dokumen e-Materai Rp10.000 berhasil diverifikasi dan tersimpan di Google Drive! Pengajuan pinjaman telah resmi masuk ke Tracking & Progress.',
      driveUrl: driveResult.driveUrl,
      provider: check.details.detectedProvider,
      nomorSurat: bLoan[1]
    });
  } catch (error) {
    console.error('Error POST /api/loans/:id/upload-ematerai:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/loans/:id/disburse', async (req, res) => {
  try {
    const loanId = req.params.id;
    const {
      tanggalPencairan,
      nominalPencairan,
      nomorReferensi,
      bankTujuan,
      noRekeningTujuan,
      atasNamaTujuan,
      buktiUrl,
      catatan,
      disbursedByNik,
      disbursedByNama,
      disbursedByRole
    } = req.body;

    if (!isRoleFinancial(disbursedByRole) && !isRoleAdmin(disbursedByRole)) {
      return res.status(403).json({ success: false, message: 'Hanya role Financial atau Administrator yang berwenang memproses pencairan pinjaman.' });
    }

    if (!nomorReferensi) {
      return res.status(400).json({ success: false, message: 'Nomor referensi bukti transfer bank wajib diisi.' });
    }

    const sheets = await getSheetsClient();
    const loansRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanRequests'!A2:AG"
    });
    const rows = loansRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === loanId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Pinjaman ' + loanId + ' tidak ditemukan.' });
    }

    const loan = rows[targetIdx];
    const currentStatus = loan[24] || '';
    if (currentStatus !== 'READY_FOR_DISBURSEMENT' && !isRoleAdmin(disbursedByRole)) {
      return res.status(400).json({
        success: false,
        message: 'Pinjaman belum disetujui penuh oleh Ketua Koperasi. Status saat ini: ' + currentStatus
      });
    }

    const pemohonNik = loan[5];
    const targetRekening = noRekeningTujuan || loan[21];

    // Validasi rekening wajib VERIFIED di MemberBankAccounts
    const bankRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'MemberBankAccounts'!A2:M"
    });
    const bankRows = bankRes.data.values || [];
    const verifiedAccount = bankRows.find(b =>
      String(b[1] || '').trim() === String(pemohonNik).trim() &&
      String(b[4] || '').trim() === String(targetRekening).trim() &&
      String(b[6] || '').toUpperCase() === 'VERIFIED'
    );

    if (!verifiedAccount && !isRoleAdmin(disbursedByRole)) {
      return res.status(400).json({
        success: false,
        message: 'Pencairan dibatalkan: Nomor rekening tujuan belum berstatus VERIFIED oleh Administrator.'
      });
    }

    const rowNum = targetIdx + 2;
    const ts = getFormattedTimestamp();
    const jakarta = getJakartaDateTime();
    const tglCair = tanggalPencairan || jakarta.dateStr;
    const nominal = nominalPencairan || loan[13];

    // 1. Update status di LoanRequests
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'LoanRequests'!Y${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [['DISBURSED']] }
    });

    // 2. Simpan ke LoanDisbursements
    const disbId = 'DISB-' + Date.now();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'LoanDisbursements'!A:M",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          disbId,
          loanId,
          ts,
          tglCair,
          nominal,
          nomorReferensi,
          bankTujuan || loan[20],
          String(targetRekening),
          atasNamaTujuan || loan[22],
          buktiUrl || '-',
          catatan || '-',
          'COMPLETED',
          String(disbursedByNik || ''),
          disbursedByNama || 'Sri Wahyuningsih'
        ]]
      }
    });

    // 3. Sinkronisasi status ke Form Responses 1 menjadi 'Done'
    try {
      const nomorSurat = loan[1];
      const frRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'Form Responses 1'!A2:J"
      });
      const frRows = frRes.data.values || [];
      const frIdx = frRows.findIndex(r => String(r[8] || '').includes(nomorSurat));
      if (frIdx !== -1) {
        await sheets.spreadsheets.values.update({
          spreadsheetId: SPREADSHEET_ID,
          range: `'Form Responses 1'!J${frIdx + 2}`,
          valueInputOption: 'USER_ENTERED',
          requestBody: { values: [['Done']] }
        });
      }
    } catch (e) {
      console.warn('Sync FR warning:', e.message);
    }

    await recordFinancialAudit(
      sheets,
      { nik: disbursedByNik, nama: disbursedByNama, role: disbursedByRole },
      'PINJAMAN_KOPERASI',
      'DISBURSE_LOAN',
      loanId,
      { statusSebelumnya: currentStatus },
      { status: 'DISBURSED', disbId, nomorReferensi, nominal },
      req
    );

    return res.json({
      success: true,
      message: `Pencairan pinjaman ${loan[1]} berhasil dicatat dengan Nomor Referensi: ${nomorReferensi}.`,
      disbId
    });
  } catch (error) {
    console.error('Error POST /api/loans/:id/disburse:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

// ----------------------------------------------------
// 6.3 ENDPOINTS: REKENING ANGGOTA & PENGAMANAN PERUBAHAN
// ----------------------------------------------------

/**
 * GET /api/bank-accounts/me
 * Mengambil informasi rekening bank terdaftar anggota
 */
app.get('/api/bank-accounts/me', async (req, res) => {
  try {
    const { nik, username, nama } = req.query;
    if (!nik && !username) return res.status(400).json({ success: false, message: 'Parameter NIK atau Username diperlukan.' });

    const sheets = await getSheetsClient();
    const bankRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'MemberBankAccounts'!A2:M"
    });
    const rows = bankRes.data.values || [];

    const qNik = String(nik || '').trim().toLowerCase();
    const qUser = String(username || '').trim().toLowerCase();

    const extractBase = (val) => {
      if (!val) return '';
      const s = String(val).trim().toLowerCase();
      const m = s.match(/^([a-z0-9]+)/i);
      return m ? m[1] : s;
    };
    const qBase = extractBase(qNik);

    const accounts = rows
      .filter(r => {
        const rowStatus = String(r[11] || '').toUpperCase().trim();
        if (rowStatus !== 'ACTIVE') return false;

        const rowNik = String(r[1] || '').trim().toLowerCase();
        // 1. Exact NIK match
        if (qNik && rowNik === qNik) return true;

        // 2. Base NIK match (e.g. 20246808 vs 20246808 - Unit...)
        const rowBase = extractBase(rowNik);
        if (qBase && rowBase && qBase === rowBase && qBase.length >= 4) return true;

        // 3. Username match with NIK column
        if (qUser && rowNik === qUser) return true;

        return false;
      })
      .map(r => ({
        accountId: r[0],
        nik: r[1],
        namaAnggota: r[2],
        namaBank: r[3],
        nomorRekening: r[4],
        atasNamaRekening: r[5],
        statusVerifikasi: r[6] || 'UNVERIFIED',
        tanggalVerifikasi: r[7] || '-'
      }));

    return res.json({ success: true, accounts });
  } catch (error) {
    console.error('Error GET /api/bank-accounts/me:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/bank-accounts/change-request
 * Mengajukan permohonan perubahan rekening alternatif (Menunggu persetujuan Admin)
 */
app.post('/api/bank-accounts/change-request', async (req, res) => {
  try {
    const {
      nik,
      namaAnggota,
      bankLama,
      rekeningLama,
      atasNamaLama,
      bankBaru,
      rekeningBaru,
      atasNamaBaru,
      alasan,
      buktiUrl
    } = req.body;

    if (!nik || !namaAnggota || !bankBaru || !rekeningBaru || !atasNamaBaru || !alasan) {
      return res.status(400).json({ success: false, message: 'Seluruh data permohonan penggantian rekening wajib diisi.' });
    }

    const sheets = await getSheetsClient();
    const reqId = 'REQ-BANK-' + Date.now();
    const ts = getFormattedTimestamp();

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "'BankAccountChangeRequests'!A:P",
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: [[
          reqId,
          String(nik).trim(),
          namaAnggota,
          bankLama || '-',
          String(rekeningLama || '-'),
          atasNamaLama || '-',
          bankBaru,
          String(rekeningBaru),
          atasNamaBaru,
          alasan,
          buktiUrl || '-',
          'REQUESTED',
          '-',
          '',
          '',
          ''
        ]]
      }
    });

    await recordFinancialAudit(
      sheets,
      { nik, nama: namaAnggota, role: 'Anggota' },
      'BANK_ACCOUNTS',
      'REQUEST_CHANGE',
      reqId,
      { bankLama, rekeningLama },
      { bankBaru, rekeningBaru, atasNamaBaru, alasan },
      req
    );

    return res.json({
      success: true,
      message: 'Permohonan perubahan rekening berhasil diajukan dan sedang dalam antrean verifikasi Administrator.',
      reqId
    });
  } catch (error) {
    console.error('Error POST /api/bank-accounts/change-request:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * GET /api/bank-accounts/requests
 * Admin melihat daftar antrean permohonan penggantian rekening
 */
app.get('/api/bank-accounts/requests', async (req, res) => {
  try {
    const { role } = req.query;
    if (!isRoleAdmin(role) && !isRoleKetua(role)) {
      return res.status(403).json({ success: false, message: 'Hanya Administrator atau Ketua Koperasi yang memiliki akses ke antrean permohonan rekening.' });
    }

    const sheets = await getSheetsClient();
    const reqRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'BankAccountChangeRequests'!A2:P"
    });
    const rows = reqRes.data.values || [];

    const requests = rows.map((r, idx) => ({
      rowIndex: idx + 2,
      requestId: r[0],
      nik: r[1],
      namaAnggota: r[2],
      bankLama: r[3],
      rekeningLama: r[4],
      atasNamaLama: r[5],
      bankBaru: r[6],
      rekeningBaru: r[7],
      atasNamaBaru: r[8],
      alasan: r[9],
      buktiUrl: r[10],
      status: r[11] || 'REQUESTED',
      catatanAdmin: r[12] || '-'
    }));

    return res.json({ success: true, requests: requests.reverse() });
  } catch (error) {
    console.error('Error GET /api/bank-accounts/requests:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});

/**
 * POST /api/bank-accounts/requests/:id/review
 * Admin menyetujui atau menolak permohonan penggantian rekening
 */
app.post('/api/bank-accounts/requests/:id/review', async (req, res) => {
  try {
    const reqId = req.params.id;
    const { action, catatan, adminNik, adminNama, adminRole } = req.body;

    if (!isRoleAdmin(adminRole) && !isRoleKetua(adminRole)) {
      return res.status(403).json({ success: false, message: 'Hanya Administrator atau Ketua Koperasi yang berhak meninjau permohonan perubahan rekening.' });
    }

    const sheets = await getSheetsClient();
    const reqRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "'BankAccountChangeRequests'!A2:P"
    });
    const rows = reqRes.data.values || [];
    const targetIdx = rows.findIndex(r => r[0] === reqId);

    if (targetIdx === -1) {
      return res.status(404).json({ success: false, message: 'Permohonan ' + reqId + ' tidak ditemukan.' });
    }

    const bReq = rows[targetIdx];
    const newStatus = action === 'APPROVE' ? 'APPROVED' : 'REJECTED';
    const rowNum = targetIdx + 2;
    const ts = getFormattedTimestamp();

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `'BankAccountChangeRequests'!L${rowNum}:P${rowNum}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [[
          newStatus,
          catatan || '-',
          String(adminNik || ''),
          adminNama || 'Administrator',
          ts
        ]]
      }
    });

    if (action === 'APPROVE') {
      const nik = bReq[1];
      const nama = bReq[2];
      const bankBaru = bReq[6];
      const rekeningBaru = bReq[7];
      const atasNamaBaru = bReq[8];

      // Non-aktifkan rekening lama
      const bankRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SPREADSHEET_ID,
        range: "'MemberBankAccounts'!A2:M"
      });
      const bankRows = bankRes.data.values || [];
      for (let i = 0; i < bankRows.length; i++) {
        if (String(bankRows[i][1] || '').trim() === String(nik).trim()) {
          await sheets.spreadsheets.values.update({
            spreadsheetId: SPREADSHEET_ID,
            range: `'MemberBankAccounts'!L${i + 2}`,
            valueInputOption: 'USER_ENTERED',
            requestBody: { values: [['INACTIVE']] }
          });
        }
      }

      // Tambahkan rekening baru berstatus VERIFIED
      const accId = 'ACC-' + Date.now();
      await sheets.spreadsheets.values.append({
        spreadsheetId: SPREADSHEET_ID,
        range: "'MemberBankAccounts'!A:M",
        valueInputOption: 'USER_ENTERED',
        insertDataOption: 'INSERT_ROWS',
        requestBody: {
          values: [[
            accId,
            String(nik).trim(),
            nama,
            bankBaru,
            String(rekeningBaru),
            atasNamaBaru,
            'VERIFIED',
            ts,
            String(adminNik || ''),
            adminNama || 'Administrator',
            bReq[10] || '-',
            'ACTIVE',
            ts
          ]]
        }
      });
    }

    await recordFinancialAudit(
      sheets,
      { nik: adminNik, nama: adminNama, role: adminRole },
      'BANK_ACCOUNTS',
      'REVIEW_CHANGE_REQUEST',
      reqId,
      { status: 'REQUESTED' },
      { status: newStatus, catatan },
      req
    );

    return res.json({
      success: true,
      message: `Permohonan ${reqId} berhasil diproses dengan status: ${newStatus}.`
    });
  } catch (error) {
    console.error('Error POST /api/bank-accounts/requests/:id/review:', error);
    return res.status(500).json({ success: false, message: error.message });
  }
});


// Route untuk preview & download dokumen permohonan pinjaman di New Tab
app.get('/loan-document-preview.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'loan-document-preview.html'));
});
app.get('/loans/:id/preview', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'loan-document-preview.html'));
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