function doGet(e) {
  setupAllSheets();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Koperasi Satya Ardhia - Sistem Terpadu & Presensi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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

// ================= AUTHENTICATION & USER MANAGEMENT =================

// 1. Login User dengan Validasi Status Lengkap (Active, Pending, Deactive, Rejected)
function loginUser(username, password) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { status: 'error', message: 'Sheet Users belum dibuat!' };

  const data = userSheet.getDataRange().getValues();
  const rows = data.slice(1);

  const u = String(username).trim().toLowerCase();
  const p = String(password).trim();

  for (let i = 0; i < rows.length; i++) {
    const rUsername = String(rows[i][0]).trim().toLowerCase();
    const rPassword = String(rows[i][1]).trim();
    const rRole = String(rows[i][2]).trim();
    const rNama = String(rows[i][3]).trim();
    const rNik = String(rows[i][4]).trim();
    const rEmail = String(rows[i][5]).trim();
    const rAccountStatus = rows[i][7] ? String(rows[i][7]).trim() : (rRole === 'Admin' ? 'Active' : 'Pending');

    if (rUsername === u && rPassword === p) {
      // Pengecekan Status Akun
      if (rAccountStatus === 'Pending') {
        return { 
          status: 'error', 
          message: 'Akun Anda sedang menunggu persetujuan dari Admin Koperasi.' 
        };
      }
      if (rAccountStatus === 'Rejected') {
        return { 
          status: 'error', 
          message: 'Akun Anda telah ditolak/dinonaktifkan oleh Admin. Silakan hubungi pengurus koperasi.' 
        };
      }
      if (rAccountStatus === 'Deactive') {
        return { 
          status: 'error', 
          message: 'Akun Anda sedang dinonaktifkan (Deactive). Silakan hubungi Admin untuk mengaktifkannya kembali.' 
        };
      }

      return {
        status: 'success',
        user: {
          username: rows[i][0],
          role: rRole,
          nama: rNama,
          namaLengkap: rNama,
          nik: rNik,
          email: rEmail,
          accountStatus: rAccountStatus
        }
      };
    }
  }
  return { status: 'error', message: 'Username atau password salah!' };
}

// 2. Registrasi Anggota Baru (Default Status = Pending)
function registerUser(userData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let userSheet = ss.getSheetByName('Users');
    
    if (!userSheet) {
      userSheet = ss.insertSheet('Users');
      userSheet.appendRow(['Username', 'Password', 'Role', 'Nama Lengkap', 'NIK', 'Email', 'ResetToken', 'AccountStatus']);
    }

    const data = userSheet.getDataRange().getValues();
    const inputUsername = String(userData.username || '').trim().toLowerCase();
    const inputEmail = String(userData.email || '').trim().toLowerCase();
    const inputNik = String(userData.nik || '').trim();
    const inputNama = String(userData.nama || '').trim();
    const inputPassword = String(userData.password || '').trim();

    if (!inputUsername || !inputPassword || !inputNama || !inputNik || !inputEmail) {
      return { status: 'error', message: 'Semua kolom registrasi wajib diisi!' };
    }

    // Cek Username atau Email ganda
    for (let i = 1; i < data.length; i++) {
      const existingUsername = String(data[i][0]).trim().toLowerCase();
      const existingEmail = String(data[i][5]).trim().toLowerCase();

      if (existingUsername === inputUsername) {
        return { status: 'error', message: 'Username sudah digunakan, silakan pilih username lain.' };
      }
      if (existingEmail === inputEmail) {
        return { status: 'error', message: 'Email sudah terdaftar pada akun lain.' };
      }
    }

    // Append baris baru dengan Status = Pending
    userSheet.appendRow([
      userData.username.trim(),
      inputPassword,
      'User',           // Role
      inputNama,
      inputNik,
      inputEmail,
      '',               // ResetToken
      'Pending'         // AccountStatus
    ]);

    return { 
      status: 'success', 
      message: 'Registrasi berhasil! Akun Anda sedang menunggu persetujuan Admin sebelum dapat digunakan untuk login.' 
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal melakukan registrasi: ' + err.toString() };
  }
}

// 3. Khusus Admin: Ambil Seluruh Data User
function getAllUsersList(requestingUserRole) {
  if (!isRoleAdmin(requestingUserRole) && !isRoleKetua(requestingUserRole) && !isRolePengurus(requestingUserRole)) {
    return []; // Khusus Admin, Ketua Koperasi, & Pengurus
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return [];

  const data = userSheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  return data.slice(1).map((row, index) => {
    return {
      rowIndex: index + 2,
      username: row[0],
      role: row[2] || 'User',
      namaLengkap: row[3] || '-',
      nik: row[4] || '-',
      email: row[5] || '-',
      accountStatus: row[7] ? String(row[7]).trim() : (row[2] === 'Admin' ? 'Active' : 'Pending')
    };
  });
}

// 4. Update Status Akun User oleh Admin
function updateUserAccountStatus(requestingUserRole, rowIndex, newStatus) {
  if (!isRoleAdmin(requestingUserRole) && !isRoleKetua(requestingUserRole) && !isRolePengurus(requestingUserRole)) {
    return { status: 'error', message: 'Akses ditolak! Hanya Admin, Ketua Koperasi, dan Pengurus yang memiliki hak ini.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    userSheet.getRange(Number(rowIndex), 8).setValue(newStatus);
    
    let msg = 'Status akun berhasil diubah menjadi ' + newStatus;
    if (newStatus === 'Active') msg = 'Akun berhasil disetujui & diaktifkan!';
    if (newStatus === 'Rejected') msg = 'Akun berhasil ditolak (Rejected)!';
    if (newStatus === 'Deactive') msg = 'Akun berhasil dinonaktifkan (Deactive)!';

    return { status: 'success', message: msg };
  } catch (err) {
    return { status: 'error', message: 'Gagal update status: ' + err.toString() };
  }
}


// Update Role Akun User oleh Admin / Ketua Koperasi
function updateUserRole(requestingUserRole, rowIndex, newRole) {
  if (!isRoleAdmin(requestingUserRole) && !isRoleKetua(requestingUserRole)) {
    return { status: 'error', message: 'Akses ditolak! Hanya Admin dan Ketua Koperasi yang dapat mengubah role.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const userSheet = ss.getSheetByName('Users');
    userSheet.getRange(Number(rowIndex), 3).setValue(newRole);
    return { status: 'success', message: 'Role akun berhasil diubah menjadi ' + newRole };
  } catch (err) {
    return { status: 'error', message: 'Gagal update role: ' + err.toString() };
  }
}

// 5. Request Lupa Password
function requestPasswordReset(identifier) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  if (!userSheet) return { status: 'error', message: 'Sheet Users tidak ditemukan.' };

  const data = userSheet.getDataRange().getValues();
  const search = String(identifier).trim().toLowerCase();
  let foundRow = -1;
  let userEmail = '';
  let userName = '';

  for (let i = 1; i < data.length; i++) {
    const uName = String(data[i][0]).trim().toLowerCase();
    const uEmail = String(data[i][5]).trim().toLowerCase();

    if (uName === search || uEmail === search) {
      foundRow = i + 1;
      userEmail = String(data[i][5]).trim();
      userName = String(data[i][3]).trim() || data[i][0];
      break;
    }
  }

  if (foundRow === -1 || !userEmail) {
    return { status: 'error', message: 'Akun atau Email tidak ditemukan!' };
  }

  const token = Math.floor(100000 + Math.random() * 900000).toString();
  userSheet.getRange(foundRow, 7).setValue(token);

  try {
    MailApp.sendEmail({
      to: userEmail,
      subject: 'Kode Reset Password - Koperasi Satya Ardhia - HLP',
      htmlBody: `
        <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6; color: #333;">
          <h2 style="color: #059669;">Permintaan Reset Password</h2>
          <p>Halo <b>${userName}</b>,</p>
          <p>Gunakan kode verifikasi (OTP) berikut untuk mengatur ulang password akun Anda:</p>
          <div style="font-size: 24px; font-weight: bold; letter-spacing: 4px; color: #059669; background: #ecfdf5; padding: 12px 24px; display: inline-block; border-radius: 8px; margin: 10px 0;">
            ${token}
          </div>
          <p style="font-size: 12px; color: #777;">Jika Anda tidak merasa meminta reset password, silakan abaikan email ini.</p>
        </div>
      `
    });
    return { status: 'success', message: 'Kode OTP telah dikirim ke email terdaftar.', email: userEmail };
  } catch (err) {
    return { status: 'error', message: 'Gagal mengirim email: ' + err.toString() };
  }
}

// 6. Reset Password Baru
function verifyAndResetPassword(emailOrUsername, token, newPassword) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const userSheet = ss.getSheetByName('Users');
  const data = userSheet.getDataRange().getValues();
  const search = String(emailOrUsername).trim().toLowerCase();
  const inputToken = String(token).trim();

  for (let i = 1; i < data.length; i++) {
    const uName = String(data[i][0]).trim().toLowerCase();
    const uEmail = String(data[i][5]).trim().toLowerCase();
    const savedToken = String(data[i][6]).trim();

    if ((uName === search || uEmail === search) && savedToken === inputToken && inputToken !== '') {
      userSheet.getRange(i + 1, 2).setValue(String(newPassword).trim());
      userSheet.getRange(i + 1, 7).setValue('');
      return { status: 'success', message: 'Password berhasil diubah! Silakan login kembali.' };
    }
  }
  return { status: 'error', message: 'Kode OTP salah atau sudah kedaluwarsa!' };
}

// ================= DATA PENGAJUAN BIAYA =================

function getDataPengajuan(userRole, userNik, userName, userUsername) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Form Responses 1') || ss.getSheets()[0];
  const lastRow = sheet.getLastRow();
  
  if (lastRow <= 1) return [];
  
  const values = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  
  let list = values.map((row, index) => {
    let formattedDate = row[0];
    if (row[0] instanceof Date) {
      formattedDate = Utilities.formatDate(row[0], "GMT+7", "dd/MM/yyyy HH:mm");
    }
    
    return {
      rowIndex: index + 2,
      datetime: formattedDate || '-',
      timestamp: formattedDate || '-',
      tipeFormulir: row[1] || 'Pengajuan Biaya',
      namaKaryawan: row[2] || '-',
      nikUnit: row[3] || '-',
      nikUnitKerja: row[3] || '-',
      noRekening: row[4] ? String(row[4]) : '-',
      atasNamaRekening: row[5] || '-',
      totalDitransfer: Number(row[6]) || 0,
      cicilan: row[7] || '1',
      keterangan: row[8] || '-',
      status: row[9] ? String(row[9]).trim() : 'Created'
    };
  }).reverse();

  // Hak Akses Tracking & Progress:
  // Yang bisa melihat semua data adalah role Administrator, Ketua Koperasi, dan Financial.
  // Selain itu (role lainnya), hanya bisa melihat riwayat tracking and progress milik dirinya sendiri
  // sesuai data berdasarkan username, nama lengkap, atau NIK.
  const canSeeAll = isRoleAdmin(userRole) || isRoleKetua(userRole) || isRoleFinancial(userRole);

  if (!canSeeAll) {
    let cleanNik = String(userNik || '').trim().toLowerCase();
    let cleanName = String(userName || '').trim().toLowerCase().replace(/\s+/g, ' ');
    let cleanUname = String(userUsername || '').trim().toLowerCase();

    // Auto-resolve dari Users sheet jika cleanName kosong tetapi username ada
    if ((!cleanName || !cleanNik) && cleanUname) {
      try {
        const uSheet = ss.getSheetByName('Users');
        if (uSheet) {
          const uData = uSheet.getDataRange().getValues();
          for (let uIdx = 1; uIdx < uData.length; uIdx++) {
            if (String(uData[uIdx][0] || '').trim().toLowerCase() === cleanUname) {
              if (!cleanName) cleanName = String(uData[uIdx][3] || '').trim().toLowerCase().replace(/\s+/g, ' ');
              if (!cleanNik) cleanNik = String(uData[uIdx][4] || '').trim().toLowerCase();
              break;
            }
          }
        }
      } catch (errU) {}
    }

    // Ekstrak angka unik jika NIK berupa nomor identitas (misal: 20246808)
    const nikMatch = cleanNik.match(/\d{3,}/);
    const nikOnly = nikMatch ? nikMatch[0] : '';
    const cleanUnameWords = cleanUname.replace(/[^a-z0-9]/g, ' ').trim();

    list = list.filter(item => {
      const rowNama = String(item.namaKaryawan || '').trim().toLowerCase().replace(/\s+/g, ' ');
      const rowNikUnit = String(item.nikUnit || item.nikUnitKerja || '').trim().toLowerCase();
      const rowAtasNama = String(item.atasNamaRekening || '').trim().toLowerCase().replace(/\s+/g, ' ');

      // 1. Prioritaskan kecocokan Nama Lengkap (Nama Karyawan atau Atas Nama Rekening)
      if (cleanName) {
        if (rowNama === cleanName || rowNama.includes(cleanName) || cleanName.includes(rowNama)) {
          return true;
        }
        if (rowAtasNama && (rowAtasNama === cleanName || rowAtasNama.includes(cleanName) || cleanName.includes(rowAtasNama))) {
          return true;
        }
      }

      // 2. Kecocokan Username
      if (cleanUname) {
        if (rowNama === cleanUname || rowNama.includes(cleanUname)) {
          return true;
        }
        if (rowNikUnit === cleanUname || rowNikUnit.includes(cleanUname)) {
          return true;
        }
        if (cleanUnameWords && (rowNama.includes(cleanUnameWords) || cleanUnameWords.includes(rowNama))) {
          return true;
        }
      }

      // 3. Kecocokan NIK
      if (nikOnly && rowNikUnit.includes(nikOnly)) {
        return true;
      }
      if (cleanNik && !['karyawan koperasi', 'koperasi', 'unit', 'staff', 'avsec', 'arff', '-'].includes(cleanNik)) {
        if (rowNikUnit === cleanNik || rowNikUnit.startsWith(cleanNik + ' ') || rowNikUnit.endsWith(' ' + cleanNik) || rowNikUnit.includes(cleanNik)) {
          return true;
        }
      }

      return false;
    });
  }

  return list;
}

function simpanPengajuan(form) {
  try {
    const role = form.userRole || form.role;
    if (role && !isRoleAdmin(role) && !isRoleKetua(role) && !isRoleFinancial(role)) {
      return { status: 'error', message: 'Akses ditolak! Hanya Admin, Ketua Koperasi, dan Financial yang dapat menginput pengajuan pinjaman/biaya.' };
    }
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    
    sheet.appendRow([
      timestamp,
      form.tipeFormulir,
      form.namaKaryawan,
      form.nikUnit,
      form.noRekening,
      form.atasNamaRekening,
      Number(form.totalDitransfer),
      form.cicilan,
      form.keterangan,
      'Created'
    ]);
    
    return { status: 'success', message: 'Pengajuan berhasil disimpan!' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function updateStatusPengajuan(userRole, rowIndex, newStatus) {
  if (isRoleAnggota(userRole) || isRolePegawai(userRole)) {
    return { status: 'error', message: 'Akses ditolak! Anggota dan Pegawai tidak memiliki wewenang mengubah status pengajuan.' };
  }
  if (isRolePengurus(userRole) || isRoleKetua(userRole)) {
    if (!['On Process', 'Rejected'].includes(newStatus)) {
      return { status: 'error', message: 'Pengurus dan Ketua Koperasi hanya memiliki otoritas untuk Persetujuan (On Process) atau Penolakan (Rejected).' };
    }
  }
  if (isRoleFinancial(userRole)) {
    if (!['Done', 'On Process'].includes(newStatus)) {
      return { status: 'error', message: 'Financial hanya dapat memproses pembayaran (Done) atau status diproses (On Process).' };
    }
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    // Kolom 10 adalah kolom J (Status)
    sheet.getRange(Number(rowIndex), 10).setValue(newStatus);
    
    let msg = 'Status pengajuan berhasil diubah menjadi ' + newStatus;
    if (newStatus === 'On Process') msg = 'Pengajuan disetujui (Approved) & siap diproses transfer oleh Finance.';
    if (newStatus === 'Rejected') msg = 'Pengajuan telah ditolak oleh Pengurus.';
    if (newStatus === 'Done') msg = 'Pengajuan selesai (Dana berhasil ditransfer).';

    return { status: 'success', message: msg };
  } catch (err) {
    return { status: 'error', message: 'Gagal update status: ' + err.toString() };
  }
}

// Khusus Role Admin: Edit Data Pengajuan
function editPengajuan(userRole, rowIndex, form) {
  if (userRole !== 'Admin') {
    return { status: 'error', message: 'Akses ditolak! Hanya Admin yang dapat mengedit data pengajuan.' };
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const row = Number(rowIndex);

    sheet.getRange(row, 2, 1, 9).setValues([[
      form.tipeFormulir || '',
      form.namaKaryawan || '',
      form.nikUnitKerja || form.nikUnit || '',
      String(form.noRekening || ''),
      form.atasNamaRekening || '',
      Number(form.totalDitransfer) || 0,
      form.cicilan || '-',
      form.keterangan || '-',
      form.status || 'Created'
    ]]);

    return { status: 'success', message: 'Data pengajuan baris ke-' + row + ' berhasil diperbarui!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal mengedit pengajuan: ' + err.toString() };
  }
}

// Khusus Role Admin: Hapus Data Pengajuan
function hapusPengajuan(userRole, rowIndex) {
  if (userRole !== 'Admin') {
    return { status: 'error', message: 'Akses ditolak! Hanya Admin yang dapat menghapus data pengajuan.' };
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Form Responses 1') || SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    const row = Number(rowIndex);
    if (row < 2) {
      return { status: 'error', message: 'Baris header tidak dapat dihapus!' };
    }

    sheet.deleteRow(row);
    return { status: 'success', message: 'Data pengajuan baris ke-' + row + ' berhasil dihapus!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal menghapus pengajuan: ' + err.toString() };
  }
}

// =========================================================================
// AUTO-SETUP SPREADSHEET & FITUR ABSENSI FACE RECOGNITION + GEOFENCING HLP
// =========================================================================

// Titik Koordinat Terminal Bandara Internasional Halim Perdanakusuma (HLP)
const HALIM_TERMINAL_LAT = -6.2652;
const HALIM_TERMINAL_LNG = 106.8851;
const GEOFENCE_RADIUS_METERS = 1000; // Radius lingkaran absensi 1 km (1.000 meter)

/**
 * 1. Otomatis Membuat Sheet & Header Lengkap Jika Belum Ada
 */
function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // A. Sheet Users
  let userSheet = ss.getSheetByName('Users');
  if (!userSheet) {
    userSheet = ss.insertSheet('Users');
    userSheet.appendRow([
      'Username', 'Password', 'Role', 'Nama Lengkap', 'NIK', 'Email', 'ResetToken', 'AccountStatus'
    ]);
    userSheet.getRange('A1:H1').setFontWeight('bold').setBackground('#059669').setFontColor('#ffffff');
    userSheet.setFrozenRows(1);
    
    // Akun default Superadmin
    userSheet.appendRow([
      'admin', 'admin123', 'Admin', 'Super Admin Koperasi', '000000 - Kantor Pusat', 'admin@koperasi.co.id', '', 'Active'
    ]);
  }

  // B. Sheet Form Responses 1 (Pengajuan Pinjaman)
  let pengajuanSheet = ss.getSheetByName('Form Responses 1');
  if (!pengajuanSheet) {
    pengajuanSheet = ss.insertSheet('Form Responses 1');
    pengajuanSheet.appendRow([
      'Timestamp', 'Tipe Formulir', 'Nama Karyawan / Anggota', 'NIK - Unit Kerja', 'No. Rekening',
      'Atas Nama Rekening', 'Total Ditransfer (Rp)', 'Cicilan', 'Keterangan / Keperluan', 'Status'
    ]);
    pengajuanSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    pengajuanSheet.setFrozenRows(1);
  }

  // C. Sheet absensi-harian
  let absensiSheet = ss.getSheetByName('absensi-harian');
  if (!absensiSheet) {
    absensiSheet = ss.insertSheet('absensi-harian');
    absensiSheet.appendRow([
      'Timestamp', 'Tanggal', 'NIK', 'Nama Lengkap', 'Unit Kerja', 
      'Clock-In', 'Clock-Out', 'Status Kehadiran', 'Lokasi Lat/Long', 
      'Jarak Terminal Halim (meter)', 'Status Verifikasi Wajah'
    ]);
    absensiSheet.getRange('A1:K1').setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    absensiSheet.setFrozenRows(1);
  }

  // D. Sheet wajah-karyawan (Descriptor 128D Face-API)
  let wajahSheet = ss.getSheetByName('wajah-karyawan');
  if (!wajahSheet) {
    wajahSheet = ss.insertSheet('wajah-karyawan');
    wajahSheet.appendRow([
      'NIK', 'Nama Lengkap', 'FaceDescriptor', 'Waktu Pendaftaran'
    ]);
    wajahSheet.getRange('A1:D1').setFontWeight('bold').setBackground('#6366f1').setFontColor('#ffffff');
    wajahSheet.setFrozenRows(1);
  }

  // E. Sheet operasional-executive-car
  let opCarSheet = ss.getSheetByName('operasional-executive-car');
  if (!opCarSheet) {
    opCarSheet = ss.insertSheet('operasional-executive-car');
    opCarSheet.appendRow([
      'Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Pengguna', 'Detail Pengguna', 'Jumlah Pax', 'Catatan'
    ]);
    opCarSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#d97706').setFontColor('#ffffff');
    opCarSheet.setFrozenRows(1);
  }

  // F. Sheet operasional-air-minum
  let opAirSheet = ss.getSheetByName('operasional-air-minum');
  if (!opAirSheet) {
    opAirSheet = ss.insertSheet('operasional-air-minum');
    opAirSheet.appendRow([
      'Timestamp', 'Tanggal', 'Waktu', 'NIK', 'Nama Petugas', 'Kategori Penerima', 'Nama Tenant / Stakeholder', 'Jumlah Galon', 'Metode Pembayaran', 'Catatan'
    ]);
    opAirSheet.getRange('A1:J1').setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    opAirSheet.setFrozenRows(1);
  }

  return { status: 'success', message: 'Semua sheet berhasil dicek & disiapkan otomatis!' };
}

/**
 * 2. Hitung Jarak Menggunakan Rumus Haversine (Satuan Meter)
 */
function calculateHaversineDistance(lat1, lon1, lat2, lon2) {
  const R = 6371000; // Radius bumi dalam meter
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLon = (lon2 - lon1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
            Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
            Math.sin(dLon / 2) * Math.sin(dLon / 2);
  const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  return Math.round(R * c);
}

/**
 * 3. Simpan / Registrasi Wajah Karyawan
 */
function registerFaceDescriptor(nik, nama, descriptorJson, userRole) {
  try {
    if (userRole && !isRoleAdmin(userRole) && !isRoleFinancial(userRole) && !isRolePegawai(userRole) && !isRoleExecutiveCar(userRole) && !isRoleAirMinumSender(userRole)) {
      return { status: 'error', message: 'Akses ditolak! Pendaftaran wajah hanya diperuntukkan bagi Pegawai, Financial, Executive Car, dan Admin.' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('wajah-karyawan');
    if (!sheet) {
      setupAllSheets();
      sheet = ss.getSheetByName('wajah-karyawan');
    }

    const data = sheet.getDataRange().getValues();
    const cleanNik = String(nik).trim();
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === cleanNik) {
        foundRow = i + 1;
        break;
      }
    }

    const timestamp = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm:ss');

    if (foundRow > 0) {
      sheet.getRange(foundRow, 2).setValue(nama);
      sheet.getRange(foundRow, 3).setValue(descriptorJson);
      sheet.getRange(foundRow, 4).setValue(timestamp);
    } else {
      sheet.appendRow([cleanNik, nama, descriptorJson, timestamp]);
    }

    return { status: 'success', message: 'Wajah karyawan ' + nama + ' berhasil didaftarkan!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal mendaftarkan wajah: ' + err.toString() };
  }
}

/**
 * 4. Ambil Semua Descriptor Wajah Karyawan untuk Face-API Matching
 */
function getAllFaceDescriptors() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('wajah-karyawan');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    const list = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][2]) {
        list.push({
          nik: String(data[i][0]).trim(),
          nama: String(data[i][1] || '').trim(),
          descriptor: data[i][2],
          updatedAt: data[i][3] || ''
        });
      }
    }
    return list;
  } catch (err) {
    return [];
  }
}

/**
 * 5. Clock-In Absensi Karyawan
 */
function submitClockIn(payload) {
  try {
    const role = payload.userRole || payload.role;
    if (role && !isRoleAdmin(role) && !isRoleFinancial(role) && !isRolePegawai(role) && !isRoleExecutiveCar(role) && !isRoleAirMinumSender(role)) {
      return { status: 'error', message: 'Akses ditolak: Fitur presensi clock-in hanya dapat digunakan oleh Pegawai, Financial, Executive Car, dan Admin.' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('absensi-harian');
    if (!sheet) {
      setupAllSheets();
      sheet = ss.getSheetByName('absensi-harian');
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, 'GMT+7', 'HH:mm:ss');
    const timestampStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy HH:mm:ss');

    const cleanNik = String(payload.nik || '').trim();
    const nama = String(payload.nama || '').trim();
    const unitKerja = String(payload.unitKerja || '-').trim();
    const userLat = parseFloat(payload.latitude);
    const userLng = parseFloat(payload.longitude);
    const faceVerified = payload.faceVerified ? 'Terverifikasi Face-API ✅' : 'Wajah Tidak Dikenali ❌';

    // Validasi Geofencing 1 KM dari Terminal Bandara Halim Perdanakusuma
    const distanceMeters = calculateHaversineDistance(userLat, userLng, HALIM_TERMINAL_LAT, HALIM_TERMINAL_LNG);
    if (distanceMeters > GEOFENCE_RADIUS_METERS) {
      return {
        status: 'error',
        message: 'Di luar radius absensi! Jarak Anda ' + distanceMeters + ' meter dari Terminal Bandara Halim Perdanakusuma (Maksimal radius 1.000 meter / 1 KM).'
      };
    }

    // Cek apakah sudah Clock-In hari ini
    const data = sheet.getDataRange().getValues();
    // Best Practice: Validasi Wajib Biometrik Wajah Lolos
    if (!payload.faceVerified || payload.faceVerified === false || payload.faceVerified === 'false') {
      return {
        status: 'error',
        message: 'Presensi ditolak: Verifikasi biometrik wajah tidak valid atau wajah tidak cocok dengan data biometrik akun Anda!'
      };
    }

    // Best Practice: Validasi akun wajib sudah mendaftarkan wajah di sheet wajah-karyawan
    let wajahSheet = ss.getSheetByName('wajah-karyawan');
    if (wajahSheet) {
      const wData = wajahSheet.getDataRange().getValues();
      const isEnrolled = wData.slice(1).some(r => {
        const rNik = String(r[0] || '').trim().toLowerCase();
        const rNama = String(r[1] || '').trim().toLowerCase();
        return (cleanNik && rNik === cleanNik.toLowerCase()) || (nama && rNama === nama.toLowerCase());
      });
      if (!isEnrolled) {
        return {
          status: 'error',
          message: 'Presensi ditolak: Akun Anda belum mendaftarkan biometrik wajah! Silakan lakukan pendaftaran wajah terlebih dahulu.'
        };
      }
    }
    for (let i = 1; i < data.length; i++) {
      const rowTanggal = String(data[i][1]).trim();
      const rowNik = String(data[i][2]).trim();

      if (rowTanggal === dateStr && rowNik === cleanNik) {
        return {
          status: 'error',
          message: 'Anda sudah melakukan Clock-In hari ini pada pukul ' + (data[i][5] || '-') + ' WIB.'
        };
      }
    }

    // Status kehadiran & pengecekan jam kerja (GMT+7 / WIB)
    const hourStr = Utilities.formatDate(now, 'GMT+7', 'HH');
    const minStr = Utilities.formatDate(now, 'GMT+7', 'mm');
    const hours = parseInt(hourStr, 10);
    const minutes = parseInt(minStr, 10);
    const totalMinutes = hours * 60 + minutes;

    let statusKehadiran = 'Hadir Tepat Waktu';
    if (isRoleExecutiveCar(role)) {
      // Role Executive Car: bisa clockin jam 3 subuh s/d jam 11 malam (03:00 - 23:00 WIB), asumsi kerja 12 jam
      if (hours < 3) {
        return {
          status: 'error',
          message: 'Belum memasuki jam operasional Executive Car! Clock-In dibuka mulai pukul 03:00 Subuh WIB (Waktu saat ini: ' + timeStr + ' WIB).'
        };
      }
      if (hours >= 23) {
        return {
          status: 'error',
          message: 'Batas waktu Clock-In Executive Car telah berakhir (Maksimal pukul 23:00 WIB). Waktu saat ini: ' + timeStr + ' WIB.'
        };
      }
      statusKehadiran = 'Hadir (Executive Car - 12 Jam)';
    } else if (isRoleAirMinumSender(role)) {
      if (totalMinutes <= (8 * 60 + 30)) {
        statusKehadiran = 'Hadir Tepat Waktu (Pengantar Air Minum)';
      } else if (totalMinutes <= (16 * 60)) {
        statusKehadiran = 'Terlambat (Pengantar Air Minum)';
      } else {
        statusKehadiran = 'Terlambat (Lewat Jam Kantor 16:00)';
      }
    } else if (isRolePegawai(role) || isRoleFinancial(role)) {
      // Role Finance / Pegawai: jam 8 pagi sampai jam 4 sore (08:00 - 16:00 WIB)
      if (totalMinutes <= (8 * 60 + 30)) {
        statusKehadiran = 'Hadir Tepat Waktu';
      } else if (totalMinutes <= (16 * 60)) {
        statusKehadiran = 'Terlambat';
      } else {
        statusKehadiran = 'Terlambat (Lewat Jam Kantor 16:00)';
      }
    } else {
      statusKehadiran = totalMinutes <= (8 * 60 + 30) ? 'Hadir Tepat Waktu' : 'Terlambat';
    }

    const lokasiStr = 'Lat: ' + userLat.toFixed(6) + ', Lng: ' + userLng.toFixed(6);

    sheet.appendRow([
      timestampStr,
      dateStr,
      cleanNik,
      nama,
      unitKerja,
      timeStr,
      '-', // Clock-Out belum
      statusKehadiran,
      lokasiStr,
      distanceMeters,
      faceVerified
    ]);

    return {
      status: 'success',
      message: 'Clock-In berhasil dicatat pada pukul ' + timeStr + ' WIB (' + statusKehadiran + ')!',
      data: {
        tanggal: dateStr,
        clockIn: timeStr,
        statusKehadiran: statusKehadiran,
        distanceMeters: distanceMeters,
        faceVerified: faceVerified
      }
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal Clock-In: ' + err.toString() };
  }
}

/**
 * 6. Clock-Out Absensi Karyawan
 */
function submitClockOut(payload) {
  try {
    const role = payload.userRole || payload.role;
    if (role && !isRoleAdmin(role) && !isRoleFinancial(role) && !isRolePegawai(role) && !isRoleExecutiveCar(role) && !isRoleAirMinumSender(role)) {
      return { status: 'error', message: 'Akses ditolak: Fitur presensi clock-out hanya dapat digunakan oleh Pegawai, Financial, Executive Car, dan Admin.' };
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('absensi-harian');
    if (!sheet) {
      setupAllSheets();
      sheet = ss.getSheetByName('absensi-harian');
    }

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, 'GMT+7', 'HH:mm:ss');

    const cleanNik = String(payload.nik || '').trim();
    const userLat = parseFloat(payload.latitude);
    const userLng = parseFloat(payload.longitude);

    // Validasi Geofencing 1 KM
    const distanceMeters = calculateHaversineDistance(userLat, userLng, HALIM_TERMINAL_LAT, HALIM_TERMINAL_LNG);
    if (distanceMeters > GEOFENCE_RADIUS_METERS) {
      return {
        status: 'error',
        message: 'Di luar radius absensi! Jarak Anda ' + distanceMeters + ' meter dari Terminal Bandara Halim Perdanakusuma (Maksimal radius 1.000 meter / 1 KM).'
      }
    if (payload.faceVerified !== undefined && (payload.faceVerified === false || payload.faceVerified === 'false')) {
      return {
        status: 'error',
        message: 'Clock-Out ditolak: Verifikasi biometrik wajah tidak cocok dengan akun Anda!'
      };
    };
    }

    const data = sheet.getDataRange().getValues();
    let foundRow = -1;

    for (let i = 1; i < data.length; i++) {
      const rowTanggal = String(data[i][1]).trim();
      const rowNik = String(data[i][2]).trim();

      if (rowTanggal === dateStr && rowNik === cleanNik) {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      return {
        status: 'error',
        message: 'Anda belum melakukan Clock-In hari ini! Silakan Clock-In terlebih dahulu.'
      };
    }

    // Validasi Wajib Isi Data Operasional Sebelum Clock-Out (GAS)
    if (isRoleExecutiveCar(role)) {
      let carSheet = ss.getSheetByName('operasional-executive-car');
      if (carSheet) {
        const carData = carSheet.getDataRange().getValues();
        const hasCarLog = carData.slice(1).some(function(r) {
          return String(r[1]).trim() === dateStr && String(r[3]).trim().toLowerCase() === cleanNik.toLowerCase();
        });
        if (!hasCarLog) {
          return {
            status: 'error',
            isOperationalRequired: true,
            operationalType: 'executive-car',
            message: 'Clock-Out ditolak: Anda wajib mengisi Data Penggunaan Executive Car Harian sebelum melakukan Clock-Out! Silakan isi data di menu Data Operasional.'
          };
        }
      }
    }

    if (isRoleAirMinumSender(role)) {
      let airSheet = ss.getSheetByName('operasional-air-minum');
      if (airSheet) {
        const airData = airSheet.getDataRange().getValues();
        const hasAirLog = airData.slice(1).some(function(r) {
          return String(r[1]).trim() === dateStr && String(r[3]).trim().toLowerCase() === cleanNik.toLowerCase();
        });
        if (!hasAirLog) {
          return {
            status: 'error',
            isOperationalRequired: true,
            operationalType: 'air-minum',
            message: 'Clock-Out ditolak: Anda wajib mengisi Data Pengantaran Galon Air Minum hari ini sebelum melakukan Clock-Out (pulang)! Silakan isi data di menu Data Operasional.'
          };
        }
      }
    }

    // Hitung durasi kerja dari Clock-In (GMT+7 / WIB)
    const hourStr = Utilities.formatDate(now, 'GMT+7', 'HH');
    const minStr = Utilities.formatDate(now, 'GMT+7', 'mm');
    const hours = parseInt(hourStr, 10);
    const minutes = parseInt(minStr, 10);
    const totalMinutes = hours * 60 + minutes;

    const existingClockIn = String(data[foundRow - 1][5] || '').trim();
    let durationHours = 0;
    let durationMinutes = 0;
    if (existingClockIn && existingClockIn.indexOf(':') !== -1) {
      const parts = existingClockIn.split(':');
      const inH = parseInt(parts[0], 10) || 0;
      const inM = parseInt(parts[1], 10) || 0;
      let diff = totalMinutes - (inH * 60 + inM);
      if (diff < 0) diff = 0;
      durationHours = Math.floor(diff / 60);
      durationMinutes = diff % 60;
    }

    let outMessage = 'Clock-Out berhasil dicatat pada pukul ' + timeStr + ' WIB. Terima kasih atas kerja keras hari ini!';
    if (isRoleExecutiveCar(role)) {
      if (hours > 23 || (hours === 23 && minutes > 0)) {
        outMessage = 'Clock-Out berhasil pada pukul ' + timeStr + ' WIB (Melewati batas reguler 23:00 WIB. Durasi kerja: ' + durationHours + ' Jam ' + durationMinutes + ' Menit).';
      } else if (durationHours >= 12) {
        outMessage = 'Clock-Out berhasil dicatat pada pukul ' + timeStr + ' WIB (Durasi kerja: ' + durationHours + ' Jam ' + durationMinutes + ' Menit - Target 12 Jam Terpenuhi ✅). Terima kasih!';
      } else {
        outMessage = 'Clock-Out berhasil dicatat pada pukul ' + timeStr + ' WIB (Durasi kerja: ' + durationHours + ' Jam ' + durationMinutes + ' Menit dari asumsi 12 Jam). Terima kasih!';
      }
    } else if (isRolePegawai(role) || isRoleFinancial(role)) {
      if (totalMinutes < (16 * 60)) {
        outMessage = 'Clock-Out berhasil dicatat pada pukul ' + timeStr + ' WIB (Catatan: Pulang lebih awal sebelum jam kantor 16:00 WIB). Terima kasih!';
      } else {
        outMessage = 'Clock-Out berhasil dicatat pada pukul ' + timeStr + ' WIB. Jam kantor telah usai (16:00 WIB). Terima kasih!';
      }
    }

    // Update kolom Clock-Out (Kolom 7 / G)
    sheet.getRange(foundRow, 7).setValue(timeStr);

    return {
      status: 'success',
      message: outMessage,
      data: {
        tanggal: dateStr,
        clockOut: timeStr,
        distanceMeters: distanceMeters
      }
    };
  } catch (err) {
    return { status: 'error', message: 'Gagal Clock-Out: ' + err.toString() };
  }
}

/**
 * 7. Status Absensi Hari Ini
 */
function getTodayAbsensiStatus(nik) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('absensi-harian');
    if (!sheet) return null;

    const now = new Date();
    const dateStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const cleanNik = String(nik).trim();

    const data = sheet.getDataRange().getValues();
    for (let i = data.length - 1; i >= 1; i--) {
      const rowTanggal = String(data[i][1]).trim();
      const rowNik = String(data[i][2]).trim();

      if (rowTanggal === dateStr && rowNik === cleanNik) {
        return {
          rowIndex: i + 1,
          tanggal: data[i][1],
          clockIn: data[i][5],
          clockOut: data[i][6],
          statusKehadiran: data[i][7],
          lokasi: data[i][8],
          jarakMeters: data[i][9],
          faceVerified: data[i][10]
        };
      }
    }
    return null;
  } catch (err) {
    return null;
  }
}

/**
 * 8. Riwayat Absensi Karyawan
 */
function getAbsensiHistoryList(userRole, nik, nama, startDate, endDate, search, todayOnly) {
  try {
    if (isRoleAnggota(userRole)) {
      return [];
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('absensi-harian');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];

    let list = data.slice(1).map((r, idx) => ({
      rowIndex: idx + 2,
      timestamp: r[0],
      tanggal: r[1],
      nik: r[2],
      nama: r[3],
      unitKerja: r[4],
      clockIn: r[5],
      clockOut: r[6],
      statusKehadiran: r[7],
      lokasi: r[8],
      jarakMeters: r[9],
      faceVerified: r[10]
    }));

    // 1. Filter hak akses: Pegawai, Financial, dan Executive Car hanya melihat miliknya sendiri
    if (isRolePegawai(userRole) || isRoleFinancial(userRole) || isRoleExecutiveCar(userRole) || isRoleAirMinumSender(userRole)) {
      const cleanNik = String(nik || '').trim().toLowerCase();
      const cleanNama = String(nama || '').trim().toLowerCase();
      list = list.filter(item => {
        const itemNik = String(item.nik || '').toLowerCase();
        const itemNama = String(item.nama || '').toLowerCase();
        return (cleanNik && itemNik.includes(cleanNik)) || (cleanNama && itemNama.includes(cleanNama));
      });
    }

    // 2. Filter Tanggal
    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');

    if (todayOnly === true || todayOnly === 'true') {
      list = list.filter(item => String(item.tanggal).trim() === todayStr);
    } else {
      if (startDate) {
        list = list.filter(item => String(item.tanggal).trim() >= String(startDate).trim());
      }
      if (endDate) {
        list = list.filter(item => String(item.tanggal).trim() <= String(endDate).trim());
      }
    }

    // 3. Filter Search
    if (search && String(search).trim()) {
      const q = String(search).trim().toLowerCase();
      list = list.filter(item => {
        const itemNama = String(item.nama || '').toLowerCase();
        const itemNik = String(item.nik || '').toLowerCase();
        const itemUnit = String(item.unitKerja || '').toLowerCase();
        return itemNama.includes(q) || itemNik.includes(q) || itemUnit.includes(q);
      });
    }

    list.reverse();
    return list;
  } catch (err) {
    return [];
  }
}

/**
 * 9. Monitoring: Rekap Seluruh Pegawai yang Masuk Hari Ini
 * Khusus: Admin, Ketua Koperasi, Pengurus
 */
function getTodayEmployeesAttendance(userRole) {
  try {
    if (isRoleAnggota(userRole) || isRolePegawai(userRole) || isRoleFinancial(userRole) || isRoleExecutiveCar(userRole) || isRoleAirMinumSender(userRole)) {
      return { status: 'error', message: 'Akses monitoring presensi hanya untuk Admin, Ketua Koperasi, dan Pengurus.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('absensi-harian');
    if (!sheet) return { status: 'success', summary: { totalHadir: 0, tepatWaktu: 0, terlambat: 0 }, data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'success', summary: { totalHadir: 0, tepatWaktu: 0, terlambat: 0 }, data: [] };

    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');

    const todayRows = [];
    for (let i = 1; i < data.length; i++) {
      const rowTanggal = String(data[i][1] || '').trim();
      if (rowTanggal === todayStr) {
        todayRows.push({
          rowIndex: i + 1,
          timestamp: data[i][0],
          tanggal: data[i][1],
          nik: data[i][2],
          nama: data[i][3],
          unitKerja: data[i][4],
          clockIn: data[i][5],
          clockOut: data[i][6],
          statusKehadiran: data[i][7],
          lokasi: data[i][8],
          jarakMeters: data[i][9],
          faceVerified: data[i][10]
        });
      }
    }
    todayRows.reverse();

    let tepatWaktu = 0;
    let terlambat = 0;
    todayRows.forEach(item => {
      if (item.statusKehadiran && String(item.statusKehadiran).toLowerCase().includes('tepat waktu')) {
        tepatWaktu++;
      } else {
        terlambat++;
      }
    });

    return {
      status: 'success',
      date: todayStr,
      summary: {
        totalHadir: todayRows.length,
        tepatWaktu: tepatWaktu,
        terlambat: terlambat
      },
      data: todayRows
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}
// ==========================================
// OPERASIONAL EXECUTIVE CAR & AIR MINUM (GAS)
// ==========================================

function submitOperasionalExecutiveCar(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-executive-car');
    if (!sheet) {
      setupAllSheets();
      sheet = ss.getSheetByName('operasional-executive-car');
    }

    const now = new Date();
    const timestampStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy HH:mm:ss');
    const dateStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, 'GMT+7', 'HH:mm:ss');

    const nik = String(payload.nik || '').trim();
    const nama = String(payload.nama || '-').trim();
    const kategori = String(payload.kategoriPengguna || '').trim();
    const detail = String(payload.detailPengguna || '-').trim();
    const pax = (payload.jumlahPax !== undefined && payload.jumlahPax !== null && payload.jumlahPax !== '') ? (parseInt(payload.jumlahPax, 10) || 0) : 0;
    const catatan = String(payload.catatan || '-').trim();

    sheet.appendRow([
      timestampStr, dateStr, timeStr, nik, nama, kategori, detail, pax, catatan
    ]);

    return { status: 'success', message: 'Data penggunaan Executive Car berhasil dicatat!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal mencatat data: ' + err.toString() };
  }
}

function submitOperasionalAirMinum(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-air-minum');
    if (!sheet) {
      setupAllSheets();
      sheet = ss.getSheetByName('operasional-air-minum');
    }

    const now = new Date();
    const timestampStr = Utilities.formatDate(now, 'GMT+7', 'dd/MM/yyyy HH:mm:ss');
    const dateStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const timeStr = Utilities.formatDate(now, 'GMT+7', 'HH:mm:ss');

    const nik = String(payload.nik || '').trim();
    const nama = String(payload.nama || '-').trim();
    const kategori = String(payload.kategoriPenerima || '').trim();
    const namaPenerima = String(payload.namaPenerima || '-').trim();
    const galon = (payload.jumlahGalon !== undefined && payload.jumlahGalon !== null && payload.jumlahGalon !== '') ? (parseInt(payload.jumlahGalon, 10) || 0) : 0;
    const metode = String(payload.metodePembayaran || 'Cash').trim();
    const catatan = String(payload.catatan || '-').trim();

    sheet.appendRow([
      timestampStr, dateStr, timeStr, nik, nama, kategori, namaPenerima, galon, metode, catatan
    ]);

    return { status: 'success', message: 'Data pengantaran galon air minum berhasil dicatat!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal mencatat data: ' + err.toString() };
  }
}

function updateOperasionalExecutiveCar(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-executive-car');
    if (!sheet) return { status: 'error', message: 'Sheet operasional-executive-car tidak ditemukan.' };

    const targetRow = parseInt(payload.rowIndex, 10);
    if (!targetRow || targetRow < 2) return { status: 'error', message: 'Row index tidak valid.' };

    const rowData = sheet.getRange(targetRow, 1, 1, 9).getValues()[0];
    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(payload.userNik || '').trim().toLowerCase();

    if (!isRoleAdmin(payload.role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return { status: 'error', message: 'Akses ditolak: Hanya penginput asli atau Administrator yang berhak mengedit data ini.' };
    }

    const pax = (payload.jumlahPax !== undefined && payload.jumlahPax !== null && payload.jumlahPax !== '') ? (parseInt(payload.jumlahPax, 10) || 0) : 0;
    const kategori = String(payload.kategoriPengguna || rowData[5] || 'UMUM').trim();
    const detail = String(payload.detailPengguna || '-').trim();
    const catatan = String(payload.catatan || '-').trim();

    sheet.getRange(targetRow, 6, 1, 4).setValues([[kategori, detail, pax, catatan]]);
    return { status: 'success', message: 'Data operasional Executive Car berhasil diperbarui!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal memperbarui data: ' + err.toString() };
  }
}

function deleteOperasionalExecutiveCar(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-executive-car');
    if (!sheet) return { status: 'error', message: 'Sheet operasional-executive-car tidak ditemukan.' };

    const targetRow = parseInt(payload.rowIndex, 10);
    if (!targetRow || targetRow < 2) return { status: 'error', message: 'Row index tidak valid.' };

    const rowData = sheet.getRange(targetRow, 1, 1, 9).getValues()[0];
    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(payload.userNik || '').trim().toLowerCase();

    if (!isRoleAdmin(payload.role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return { status: 'error', message: 'Akses ditolak: Hanya penginput asli atau Administrator yang berhak menghapus data ini.' };
    }

    sheet.deleteRow(targetRow);
    return { status: 'success', message: 'Data operasional Executive Car berhasil dihapus!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal menghapus data: ' + err.toString() };
  }
}

function updateOperasionalAirMinum(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-air-minum');
    if (!sheet) return { status: 'error', message: 'Sheet operasional-air-minum tidak ditemukan.' };

    const targetRow = parseInt(payload.rowIndex, 10);
    if (!targetRow || targetRow < 2) return { status: 'error', message: 'Row index tidak valid.' };

    const rowData = sheet.getRange(targetRow, 1, 1, 10).getValues()[0];
    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(payload.userNik || '').trim().toLowerCase();

    if (!isRoleAdmin(payload.role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return { status: 'error', message: 'Akses ditolak: Hanya penginput asli atau Administrator yang berhak mengedit data ini.' };
    }

    const galon = (payload.jumlahGalon !== undefined && payload.jumlahGalon !== null && payload.jumlahGalon !== '') ? (parseInt(payload.jumlahGalon, 10) || 0) : 0;
    const kategori = String(payload.kategoriPenerima || rowData[5] || 'ATS').trim();
    const namaPenerima = String(payload.namaPenerima || '-').trim();
    const metode = String(payload.metodePembayaran || rowData[8] || 'Cash').trim();
    const catatan = String(payload.catatan || '-').trim();

    sheet.getRange(targetRow, 6, 1, 5).setValues([[kategori, namaPenerima, galon, metode, catatan]]);
    return { status: 'success', message: 'Data pengantaran galon berhasil diperbarui!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal memperbarui data: ' + err.toString() };
  }
}

function deleteOperasionalAirMinum(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-air-minum');
    if (!sheet) return { status: 'error', message: 'Sheet operasional-air-minum tidak ditemukan.' };

    const targetRow = parseInt(payload.rowIndex, 10);
    if (!targetRow || targetRow < 2) return { status: 'error', message: 'Row index tidak valid.' };

    const rowData = sheet.getRange(targetRow, 1, 1, 10).getValues()[0];
    const rowNik = String(rowData[3] || '').trim().toLowerCase();
    const cleanUserNik = String(payload.userNik || '').trim().toLowerCase();

    if (!isRoleAdmin(payload.role) && (!cleanUserNik || cleanUserNik !== rowNik)) {
      return { status: 'error', message: 'Akses ditolak: Hanya penginput asli atau Administrator yang berhak menghapus data ini.' };
    }

    sheet.deleteRow(targetRow);
    return { status: 'success', message: 'Data pengantaran galon berhasil dihapus!' };
  } catch (err) {
    return { status: 'error', message: 'Gagal menghapus data: ' + err.toString() };
  }
}
function getOperasionalStatusToday(role, nik) {
  try {
    if (!nik || !role) return { status: 'success', filledToday: false, count: 0 };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cleanNik = String(nik).trim().toLowerCase();
    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');

    if (isRoleExecutiveCar(role)) {
      let sheet = ss.getSheetByName('operasional-executive-car');
      if (!sheet) return { status: 'success', type: 'executive-car', filledToday: false, count: 0, todayDate: todayStr };
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return { status: 'success', type: 'executive-car', filledToday: false, count: 0, todayDate: todayStr };
      let count = 0;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1] || '').trim() === todayStr && String(data[i][3] || '').trim().toLowerCase() === cleanNik) {
          count++;
        }
      }
      return { status: 'success', type: 'executive-car', filledToday: count > 0, count: count, todayDate: todayStr };
    }

    if (isRoleAirMinumSender(role)) {
      let sheet = ss.getSheetByName('operasional-air-minum');
      if (!sheet) return { status: 'success', type: 'air-minum', filledToday: false, count: 0, todayDate: todayStr };
      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return { status: 'success', type: 'air-minum', filledToday: false, count: 0, todayDate: todayStr };
      let count = 0;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][1] || '').trim() === todayStr && String(data[i][3] || '').trim().toLowerCase() === cleanNik) {
          count++;
        }
      }
      return { status: 'success', type: 'air-minum', filledToday: count > 0, count: count, todayDate: todayStr };
    }

    return { status: 'success', filledToday: true, count: 0 };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getOperasionalExecutiveCarList(role, nik, todayOnly) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-executive-car');
    if (!sheet) return { status: 'success', data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'success', data: [] };

    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const cleanNik = String(nik || '').trim().toLowerCase();

    let list = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const rTanggal = String(r[1] || '').trim();
      const rNik = String(r[3] || '').trim().toLowerCase();

      if (isRoleExecutiveCar(role) && rNik !== cleanNik) continue;
      if (todayOnly && rTanggal !== todayStr) continue;

      list.push({
        rowIndex: i + 1,
        timestamp: r[0] || '',
        tanggal: r[1] || '',
        waktu: r[2] || '',
        nik: r[3] || '',
        nama: r[4] || '',
        kategoriPengguna: r[5] || '',
        detailPengguna: r[6] || '-',
        jumlahPax: parseInt(r[7], 10) || 0,
        catatan: r[8] || '-'
      });
    }

    list.reverse();
    return { status: 'success', data: list };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getOperasionalAirMinumList(role, nik, todayOnly) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('operasional-air-minum');
    if (!sheet) return { status: 'success', data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'success', data: [] };

    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const cleanNik = String(nik || '').trim().toLowerCase();

    let list = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const rTanggal = String(r[1] || '').trim();
      const rNik = String(r[3] || '').trim().toLowerCase();

      if (isRoleAirMinumSender(role) && rNik !== cleanNik) continue;
      if (todayOnly && rTanggal !== todayStr) continue;

      list.push({
        rowIndex: i + 1,
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
      });
    }

    list.reverse();
    return { status: 'success', data: list };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getOperasionalSummary(role, period) {
  try {
    if (role && !isRoleAdmin(role) && !isRoleKetua(role) && !isRolePengurus(role)) {
      return { status: 'error', message: 'Akses dashboard hanya untuk Admin, Ketua Koperasi, dan Pengurus.' };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let carSheet = ss.getSheetByName('operasional-executive-car');
    let airSheet = ss.getSheetByName('operasional-air-minum');

    const carData = carSheet ? carSheet.getDataRange().getValues().slice(1) : [];
    const airData = airSheet ? airSheet.getDataRange().getValues().slice(1) : [];

    const now = new Date();
    const todayStr = Utilities.formatDate(now, 'GMT+7', 'yyyy-MM-dd');
    const nowMs = now.getTime();

    const filterByPeriod = function(tglStr) {
      if (!tglStr) return false;
      if (period === 'today') return tglStr === todayStr;
      if (period === '7days') {
        const d = new Date(tglStr);
        return (nowMs - d.getTime()) <= 7 * 24 * 60 * 60 * 1000;
      }
      if (period === 'month') {
        return tglStr.indexOf(todayStr.slice(0, 7)) === 0;
      }
      return true;
    };

    let totalCarPax = 0;
    const carKategoriMap = { 'UMUM': 0, 'DPR-RI': 0, 'FOC': 0, 'FS-PICA': 0, 'PRIVATE JET': 0, 'Lain-lain': 0 };
    const carTrendMap = {};
    let carTrips = 0;

    for (let i = 0; i < carData.length; i++) {
      const r = carData[i];
      const tgl = String(r[1] || '').trim();
      if (!filterByPeriod(tgl)) continue;
      carTrips++;
      const kat = String(r[5] || 'Lain-lain').trim();
      const pax = parseInt(r[7], 10) || 0;
      totalCarPax += pax;
      if (carKategoriMap.hasOwnProperty(kat)) {
        carKategoriMap[kat] += pax;
      } else {
        carKategoriMap['Lain-lain'] += pax;
      }
      carTrendMap[tgl] = (carTrendMap[tgl] || 0) + pax;
    }

    let totalAirGalon = 0;
    const airKategoriMap = { 'ATS': 0, 'API': 0, 'Tenant': 0, 'Stakeholder': 0 };
    const airBayarMap = { 'Cash': 0, 'QRIS': 0, 'Akhir Bulan': 0 };
    const airTrendMap = {};
    let airDeliveries = 0;

    for (let j = 0; j < airData.length; j++) {
      const r = airData[j];
      const tgl = String(r[1] || '').trim();
      if (!filterByPeriod(tgl)) continue;
      airDeliveries++;
      const kat = String(r[5] || 'ATS').trim();
      const galon = parseInt(r[7], 10) || 0;
      const metode = String(r[8] || 'Cash').trim();
      totalAirGalon += galon;

      if (airKategoriMap.hasOwnProperty(kat)) {
        airKategoriMap[kat] += galon;
      } else {
        airKategoriMap['Stakeholder'] += galon;
      }

      if (metode.indexOf('QRIS') !== -1) {
        airBayarMap['QRIS'] += galon;
      } else if (metode.indexOf('Bulan') !== -1 || metode.indexOf('Akhir') !== -1) {
        airBayarMap['Akhir Bulan'] += galon;
      } else {
        airBayarMap['Cash'] += galon;
      }

      airTrendMap[tgl] = (airTrendMap[tgl] || 0) + galon;
    }

    return {
      status: 'success',
      period: period || 'all',
      summary: {
        car: {
          totalPax: totalCarPax,
          totalTrips: carTrips,
          byCategory: carKategoriMap,
          trend: carTrendMap
        },
        air: {
          totalGalon: totalAirGalon,
          totalDeliveries: airDeliveries,
          byCategory: airKategoriMap,
          byPayment: airBayarMap,
          trend: airTrendMap
        }
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}


// ==========================================
// 7. MODUL SIMPANAN SUKARELA & PINJAMAN KOPERASI (GAS BACKEND)
// ==========================================

function getGasJakartaDateTime(d) {
  var dt = d || new Date();
  var ts = Utilities.formatDate(dt, "GMT+7", "dd/MM/yyyy HH:mm");
  var dateStr = Utilities.formatDate(dt, "GMT+7", "yyyy-MM-dd");
  var timeStr = Utilities.formatDate(dt, "GMT+7", "HH:mm:ss");
  var year = Utilities.formatDate(dt, "GMT+7", "yyyy");
  var month = Utilities.formatDate(dt, "GMT+7", "MM");
  var day = Utilities.formatDate(dt, "GMT+7", "dd");
  return { timestampStr: ts, dateStr: dateStr, timeStr: timeStr, year: year, month: month, day: day };
}

function getFormattedTimestamp(d) {
  return getGasJakartaDateTime(d).timestampStr;
}

function getFormattedTimeStamp(d) {
  return getGasJakartaDateTime(d).timestampStr;
}

function getCurrentTimestamp(d) {
  return getGasJakartaDateTime(d).timestampStr;
}

function getGasRomanMonth(monthZeroIndexed) {
  var roman = ['I', 'II', 'III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X', 'XI', 'XII'];
  return roman[monthZeroIndexed] || 'I';
}

function recordGasFinancialAudit(actor, modul, aksi, targetId, detailSebelum, detailSesudah) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('FinancialAuditLog');
    if (!sheet) return;
    var logId = 'AUDIT-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    var ts = Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm");
    sheet.appendRow([
      logId,
      ts,
      modul || '',
      aksi || '',
      String(targetId || ''),
      String((actor && actor.nik) || ''),
      (actor && actor.nama) || '',
      (actor && actor.role) || '',
      typeof detailSebelum === 'object' ? JSON.stringify(detailSebelum) : String(detailSebelum || ''),
      typeof detailSesudah === 'object' ? JSON.stringify(detailSesudah) : String(detailSesudah || ''),
      'GAS_WEB_APP',
      'Google Apps Script Environment'
    ]);
  } catch (err) {
    Logger.log('FinancialAuditLog Error: ' + err.toString());
  }
}

function calculateGasLoanSchedule(plafon, tenor) {
  var p = parseInt(plafon, 10);
  var t = parseInt(tenor, 10);

  if (isNaN(p) || p < 10000000 || p > 50000000 || p % 1000000 !== 0) {
    throw new Error('Plafon pinjaman harus antara Rp 10.000.000 s/d Rp 50.000.000 dengan kelipatan Rp 1.000.000');
  }

  var validTenors = p <= 30000000 ? [10, 12, 20, 24, 36] : [10, 12, 20, 24, 36, 48, 60];
  if (validTenors.indexOf(t) === -1) {
    throw new Error('Tenor ' + t + ' bulan tidak tersedia untuk plafon Rp ' + p);
  }

  var rateMap = { 10: 0.11, 12: 0.14, 20: 0.22, 24: 0.28, 36: 0.42, 48: 0.56, 60: 0.70 };
  var rate = rateMap[t];
  var totalBiaya = Math.round(p * rate);
  var totalKewajiban = p + totalBiaya;
  var cicilanNormal = Math.round(totalKewajiban / t);

  var jadwal = [];
  var akumulasi = 0;
  for (var i = 1; i <= t; i++) {
    if (i === t) {
      var cicilanAkhir = totalKewajiban - akumulasi;
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
    rate: rate,
    ratePersen: Math.round(rate * 100) + '%',
    totalBiaya: totalBiaya,
    totalKewajiban: totalKewajiban,
    cicilanBulanan: cicilanNormal,
    cicilanAkhir: jadwal[jadwal.length - 1].angsuran,
    jadwal: jadwal
  };
}

function calculateGasMemberSimpananSaldo(targetNik) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('SimpananSukarelaLedger');
  if (!sheet) {
    return { nik: targetNik, saldo: 0, totalSetoranBulanIni: 0, totalSetoranAll: 0, totalPenarikanAll: 0, mutasiCount: 0, mutasiList: [] };
  }

  var data = sheet.getDataRange().getValues();
  var jkt = getGasJakartaDateTime();
  var currentMonthYear = '-' + jkt.month + '-' + jkt.year;

  var saldo = 0;
  var totalSetoranBulanIni = 0;
  var totalSetoranAll = 0;
  var totalPenarikanAll = 0;
  var mutasiList = [];

  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    var txId = r[0] || '';
    var timestamp = r[1] || '';
    var tanggal = r[2] || '';
    var nik = String(r[3] || '').trim();
    var nama = r[4] || '';
    var unit = r[5] || '';
    var jenis = String(r[6] || '').trim().toUpperCase();
    var sumber = r[7] || '';
    var nominal = parseInt(String(r[8] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    var saldoSebelumnya = parseInt(String(r[9] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    var saldoSesudahnya = parseInt(String(r[10] || '0').replace(/[^0-9]/g, ''), 10) || 0;
    var keterangan = r[11] || '';
    var refPembayaran = r[12] || '';
    var buktiUrl = r[13] || '';
    var status = String(r[14] || '').trim().toUpperCase();
    var dibuatOlehNama = r[16] || '';
    var reversalOf = r[18] || '';

    if (nik === String(targetNik).trim()) {
      if (status === 'POSTED') {
        if (jenis === 'SETORAN' || jenis === 'OPENING_BALANCE') {
          saldo += nominal;
          totalSetoranAll += nominal;
          if (String(tanggal).indexOf(currentMonthYear) !== -1 || String(timestamp).indexOf(jkt.month + '/' + jkt.year) !== -1) {
            totalSetoranBulanIni += nominal;
          }
        } else if (jenis === 'PENARIKAN') {
          saldo -= nominal;
          totalPenarikanAll += nominal;
        } else if (jenis === 'KOREKSI') {
          saldo += nominal;
        }
      }

      mutasiList.push({
        rowIndex: i + 1,
        txId: txId,
        timestamp: timestamp,
        tanggal: tanggal,
        nik: nik,
        nama: nama,
        unit: unit,
        jenis: jenis,
        sumber: sumber,
        nominal: nominal,
        saldoSebelumnya: saldoSebelumnya,
        saldoSesudahnya: saldoSesudahnya,
        keterangan: keterangan,
        refPembayaran: refPembayaran,
        buktiUrl: buktiUrl,
        status: status,
        dibuatOlehNama: dibuatOlehNama,
        reversalOf: reversalOf
      });
    }
  }

  return {
    nik: targetNik,
    saldo: saldo,
    totalSetoranBulanIni: totalSetoranBulanIni,
    totalSetoranAll: totalSetoranAll,
    totalPenarikanAll: totalPenarikanAll,
    mutasiCount: mutasiList.length,
    mutasiList: mutasiList.reverse()
  };
}

function generateGasLoanNomorSurat() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('LoanRequests');
  var jkt = getGasJakartaDateTime();
  var roman = getGasRomanMonth(parseInt(jkt.month, 10) - 1);
  var year = jkt.year;

  var maxSeq = 80;
  if (sheet && sheet.getLastRow() > 1) {
    var values = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < values.length; i++) {
      var noSurat = String(values[i][0] || '');
      var match = noSurat.match(/^(\d+)\/SK\.SIMPIN/i);
      if (match) {
        var num = parseInt(match[1], 10);
        if (num > maxSeq) maxSeq = num;
      }
    }
  }

  var nextSeq = String(maxSeq + 1);
  while (nextSeq.length < 3) nextSeq = '0' + nextSeq;
  return nextSeq + '/SK.SIMPIN/' + roman + '/' + year;
}

// ----------------------------------------------------
// PUBLIC API FUNCTIONS FOR FRONTEND / GAS BRIDGE
// ----------------------------------------------------

function getSimpananSummary(nik, role, targetNik) {
  try {
    var isOfficer = isRoleFinancial(role) || isRoleAdmin(role) || isRoleKetua(role);
    var lookupNik = (isOfficer && targetNik) ? targetNik : nik;
    var memberData = calculateGasMemberSimpananSaldo(lookupNik);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var rulesSheet = ss.getSheetByName('SimpananSukarelaRules');
    var ruleInfo = null;

    if (rulesSheet && rulesSheet.getLastRow() > 1) {
      var rData = rulesSheet.getDataRange().getValues();
      for (var i = 1; i < rData.length; i++) {
        if (String(rData[i][2] || '').trim() === String(lookupNik).trim() && String(rData[i][10] || '').toUpperCase() === 'ACTIVE') {
          ruleInfo = {
            ruleId: rData[i][0],
            nominalRutin: parseInt(rData[i][5], 10) || 0,
            tanggalEksekusi: rData[i][7] || '25',
            status: rData[i][10],
            terakhirDijalankan: rData[i][14] || '-'
          };
          break;
        }
      }
    }

    var overallSummary = null;
    if (isOfficer) {
      var ledgerSheet = ss.getSheetByName('SimpananSukarelaLedger');
      if (ledgerSheet && ledgerSheet.getLastRow() > 1) {
        var lData = ledgerSheet.getDataRange().getValues();
        var grandTotal = 0;
        var grandSetoranBulanIni = 0;
        var membersMap = {};
        var jkt = getGasJakartaDateTime();
        var currentMonthYear = '-' + jkt.month + '-' + jkt.year;

        for (var j = 1; j < lData.length; j++) {
          var rNik = String(lData[j][3] || '').trim();
          var jns = String(lData[j][6] || '').trim().toUpperCase();
          var nom = parseInt(String(lData[j][8] || '0').replace(/[^0-9]/g, ''), 10) || 0;
          var st = String(lData[j][14] || '').trim().toUpperCase();
          var tgl = lData[j][2] || '';

          if (st === 'POSTED') {
            if (rNik) membersMap[rNik] = true;
            if (jns === 'SETORAN' || jns === 'OPENING_BALANCE') {
              grandTotal += nom;
              if (String(tgl).indexOf(currentMonthYear) !== -1) grandSetoranBulanIni += nom;
            } else if (jns === 'PENARIKAN') {
              grandTotal -= nom;
            }
          }
        }
        overallSummary = {
          grandTotalSaldo: grandTotal,
          grandSetoranBulanIni: grandSetoranBulanIni,
          totalAnggotaMenabung: Object.keys(membersMap).length
        };
      }
    }

    return { status: 'success', success: true, member: memberData, rule: ruleInfo, overall: overallSummary };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getSimpananTransactions(nik, role, targetNik, page, limit, search) {
  try {
    var isOfficer = isRoleFinancial(role) || isRoleAdmin(role) || isRoleKetua(role);
    var lookupNik = (isOfficer && targetNik) ? targetNik : (!isOfficer ? nik : null);

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SimpananSukarelaLedger');
    if (!sheet || sheet.getLastRow() <= 1) {
      return { status: 'success', success: true, data: [], totalCount: 0, page: 1, totalPages: 1 };
    }

    var data = sheet.getDataRange().getValues();
    var transactions = [];

    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var rowNik = String(r[3] || '').trim();
      if (lookupNik && rowNik !== lookupNik) continue;

      var txId = r[0] || '';
      var nama = r[4] || '';
      var keterangan = r[11] || '';
      var refPembayaran = r[12] || '';

      if (search) {
        var q = search.toLowerCase();
        var match = txId.toLowerCase().indexOf(q) !== -1 ||
                    nama.toLowerCase().indexOf(q) !== -1 ||
                    rowNik.toLowerCase().indexOf(q) !== -1 ||
                    keterangan.toLowerCase().indexOf(q) !== -1 ||
                    refPembayaran.toLowerCase().indexOf(q) !== -1;
        if (!match) continue;
      }

      transactions.push({
        rowIndex: i + 1,
        txId: txId,
        timestamp: r[1] || '',
        tanggal: r[2] || '',
        nik: rowNik,
        nama: nama,
        unit: r[5] || '',
        jenis: String(r[6] || '').trim().toUpperCase(),
        sumber: r[7] || '',
        nominal: parseInt(String(r[8] || '0').replace(/[^0-9]/g, ''), 10) || 0,
        saldoSebelumnya: parseInt(String(r[9] || '0').replace(/[^0-9]/g, ''), 10) || 0,
        saldoSesudahnya: parseInt(String(r[10] || '0').replace(/[^0-9]/g, ''), 10) || 0,
        keterangan: keterangan,
        refPembayaran: refPembayaran,
        buktiUrl: r[13] || '',
        status: String(r[14] || '').trim().toUpperCase(),
        dibuatOlehNama: r[16] || '',
        reversalOf: r[18] || ''
      });
    }

    transactions.reverse();
    var p = parseInt(page, 10) || 1;
    var l = parseInt(limit, 10) || 20;
    var paginated = transactions.slice((p - 1) * l, p * l);

    return {
      status: 'success',
      success: true,
      data: paginated,
      totalCount: transactions.length,
      page: p,
      totalPages: Math.ceil(transactions.length / l)
    };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function submitSimpananTransaction(payload) {
  try {
    var role = payload.createdByRole;
    if (!isRoleFinancial(role) && !isRoleAdmin(role)) {
      return { status: 'error', success: false, message: 'Hanya role Financial atau Administrator yang berhak mencatat mutasi simpanan.' };
    }

    var nom = parseInt(payload.nominal, 10);
    if (isNaN(nom) || nom <= 0) {
      return { status: 'error', success: false, message: 'Nominal transaksi harus berupa angka positif.' };
    }

    var jenisUpper = String(payload.jenisTransaksi || '').trim().toUpperCase();
    var currentMember = calculateGasMemberSimpananSaldo(payload.nik);
    var saldoSebelumnya = currentMember.saldo;

    if (jenisUpper === 'PENARIKAN' && nom > saldoSebelumnya) {
      return { status: 'error', success: false, message: 'Saldo tidak mencukupi untuk penarikan. Saldo saat ini: Rp ' + saldoSebelumnya };
    }

    var saldoSesudahnya = saldoSebelumnya;
    if (jenisUpper === 'SETORAN' || jenisUpper === 'OPENING_BALANCE') saldoSesudahnya += nom;
    else if (jenisUpper === 'PENARIKAN') saldoSesudahnya -= nom;
    else if (jenisUpper === 'KOREKSI') saldoSesudahnya += nom;

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SimpananSukarelaLedger');
    var txId = 'TX-SIMP-' + new Date().getTime();
    var jkt = getGasJakartaDateTime();
    var tgl = payload.tanggal || jkt.dateStr;

    sheet.appendRow([
      txId,
      jkt.timestampStr,
      tgl,
      String(payload.nik).trim(),
      payload.nama,
      payload.unitKerja || '',
      jenisUpper,
      'MANUAL_FINANCIAL',
      nom,
      saldoSebelumnya,
      saldoSesudahnya,
      payload.keterangan || '-',
      payload.refPembayaran || '-',
      payload.buktiUrl || '-',
      'POSTED',
      String(payload.createdByNik || ''),
      payload.createdByNama || '',
      payload.createdByRole || '',
      ''
    ]);

    recordGasFinancialAudit(
      { nik: payload.createdByNik, nama: payload.createdByNama, role: payload.createdByRole },
      'SIMPANAN_SUKARELA',
      'INPUT_' + jenisUpper,
      txId,
      { saldoSebelumnya: saldoSebelumnya },
      { saldoSesudahnya: saldoSesudahnya, nominal: nom, nik: payload.nik }
    );

    return {
      status: 'success',
      success: true,
      message: 'Transaksi ' + jenisUpper + ' sebesar Rp ' + nom + ' berhasil dicatat.',
      txId: txId,
      saldoSesudahnya: saldoSesudahnya
    };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function submitSimpananReversal(payload) {
  try {
    var role = payload.createdByRole;
    if (!isRoleFinancial(role) && !isRoleAdmin(role)) {
      return { status: 'error', success: false, message: 'Hanya role Financial atau Administrator yang berhak membatalkan mutasi simpanan.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SimpananSukarelaLedger');
    var data = sheet.getDataRange().getValues();
    var targetRowIdx = -1;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.txId) {
        targetRowIdx = i + 1;
        break;
      }
    }

    if (targetRowIdx === -1) {
      return { status: 'error', success: false, message: 'Transaksi tidak ditemukan.' };
    }

    var oldRow = data[targetRowIdx - 1];
    if (String(oldRow[14]).toUpperCase() === 'REVERSED') {
      return { status: 'error', success: false, message: 'Transaksi sudah dibatalkan sebelumnya.' };
    }

    var nik = oldRow[3];
    var nama = oldRow[4];
    var unit = oldRow[5];
    var oldJenis = String(oldRow[6] || '').toUpperCase();
    var oldNominal = parseInt(String(oldRow[8] || '0').replace(/[^0-9]/g, ''), 10) || 0;

    var cur = calculateGasMemberSimpananSaldo(nik);
    var saldoSebelumnya = cur.saldo;
    var saldoSesudahnya = saldoSebelumnya;
    if (oldJenis === 'SETORAN' || oldJenis === 'OPENING_BALANCE' || oldJenis === 'KOREKSI') {
      saldoSesudahnya = saldoSebelumnya - oldNominal;
    } else if (oldJenis === 'PENARIKAN') {
      saldoSesudahnya = saldoSebelumnya + oldNominal;
    }

    // Set REVERSED pada baris lama
    sheet.getRange(targetRowIdx, 15).setValue('REVERSED');

    var revTxId = 'REV-' + new Date().getTime();
    var jkt = getGasJakartaDateTime();

    sheet.appendRow([
      revTxId,
      jkt.timestampStr,
      jkt.dateStr,
      String(nik),
      nama,
      unit,
      'REVERSAL',
      'MANUAL_REVERSAL',
      oldNominal,
      saldoSebelumnya,
      saldoSesudahnya,
      'Reversal dari ' + payload.txId + ': ' + payload.reason,
      payload.txId,
      '-',
      'POSTED',
      String(payload.createdByNik || ''),
      payload.createdByNama || '',
      payload.createdByRole || '',
      payload.txId
    ]);

    recordGasFinancialAudit(
      { nik: payload.createdByNik, nama: payload.createdByNama, role: payload.createdByRole },
      'SIMPANAN_SUKARELA',
      'REVERSAL_TRANSACTION',
      payload.txId,
      { status: 'POSTED', nominal: oldNominal },
      { status: 'REVERSED', reversalTxId: revTxId }
    );

    return { status: 'success', success: true, message: 'Transaksi berhasil dibatalkan.', reversalTxId: revTxId, saldoSesudahnya: saldoSesudahnya };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getLoanRates() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LoanRateMaster');
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', success: true, rates: [] };

    var data = sheet.getDataRange().getValues();
    var rates = [];
    for (var i = 1; i < data.length; i++) {
      rates.push({
        rateId: data[i][0],
        plafonMin: parseInt(data[i][1], 10) || 0,
        plafonMax: parseInt(data[i][2], 10) || 0,
        tenorBulan: parseInt(data[i][3], 10) || 0,
        persentaseTotal: parseFloat(data[i][4]) || 0,
        keterangan: data[i][5] || ''
      });
    }
    return { status: 'success', success: true, rates: rates };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function simulateLoanCalculation(plafon, tenor) {
  try {
    var sim = calculateGasLoanSchedule(plafon, tenor);
    return { status: 'success', success: true, simulation: sim };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function applyLoanRequest(payload) {
  try {
    var calc = calculateGasLoanSchedule(payload.jumlahPinjaman, payload.tenorBulan);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Pastikan rekening terdaftar
    var bankSheet = ss.getSheetByName('MemberBankAccounts');
    if (bankSheet) {
      var bData = bankSheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < bData.length; i++) {
        if (String(bData[i][1]).trim() === String(payload.pemohonNik).trim() && String(bData[i][4]).trim() === String(payload.bankNoRekening).trim()) {
          found = true;
          break;
        }
      }
      if (!found) {
        var accId = 'ACC-' + new Date().getTime();
        var jktNow = getGasJakartaDateTime();
        bankSheet.appendRow([
          accId,
          String(payload.pemohonNik).trim(),
          payload.pemohonNama,
          payload.bankNama,
          String(payload.bankNoRekening),
          payload.bankAtasNama,
          'VERIFIED',
          jktNow.timestampStr,
          String(payload.createdByNik || ''),
          payload.createdByNama || '',
          '-',
          'ACTIVE',
          jktNow.timestampStr
        ]);
      }
    }

    var nomorSurat = generateGasLoanNomorSurat();
    var loanId = 'LN-' + new Date().getTime();
    var jkt = getGasJakartaDateTime();

    var startMonth = (parseInt(jkt.month, 10) % 12) + 1;
    var startYear = parseInt(jkt.month, 10) === 12 ? parseInt(jkt.year, 10) + 1 : parseInt(jkt.year, 10);
    var tmtMulaiGaji = (startMonth < 10 ? '0' + startMonth : startMonth) + '/' + startYear;
    var endTotalMonths = (startYear * 12 + startMonth - 1) + calc.tenor - 1;
    var endYear = Math.floor(endTotalMonths / 12);
    var endMonth = (endTotalMonths % 12) + 1;
    var tmtSelesaiGaji = (endMonth < 10 ? '0' + endMonth : endMonth) + '/' + endYear;

    var reqSheet = ss.getSheetByName('LoanRequests');
    reqSheet.appendRow([
      loanId,
      nomorSurat,
      jkt.timestampStr,
      jkt.dateStr,
      isRoleFinancial(payload.createdByRole) ? 'FINANCIAL' : 'ANGGOTA',
      String(payload.pemohonNik).trim(),
      payload.pemohonNama,
      payload.pemohonUnit || '',
      payload.kantorCabang || 'Bandara Internasional Halim Perdanakusuma (HLP)',
      payload.penghasilanBulanan || 0,
      payload.alamat || '',
      String(payload.noTelepon || ''),
      payload.keperluan,
      calc.plafon,
      calc.tenor,
      calc.ratePersen,
      calc.totalKewajiban,
      calc.cicilanBulanan,
      tmtMulaiGaji,
      tmtSelesaiGaji,
      payload.bankNama,
      String(payload.bankNoRekening),
      payload.bankAtasNama,
      'VERIFIED',
      'PENDING_PIC_SIMPIN',
      '',
      '',
      String(payload.createdByNik || ''),
      payload.createdByNama || '',
      payload.createdByRole || ''
    ]);

    // Backward compatibility Form Responses 1
    var frSheet = ss.getSheetByName('Form Responses 1');
    if (frSheet) {
      frSheet.appendRow([
        jkt.timestampStr,
        'Peminjaman Tunai',
        payload.pemohonNama,
        payload.pemohonUnit ? (payload.pemohonNik + ' - ' + payload.pemohonUnit) : payload.pemohonNik,
        String(payload.bankNoRekening),
        payload.bankAtasNama,
        'Rp' + calc.plafon,
        calc.tenor,
        nomorSurat + ': ' + payload.keperluan + ' (Cicilan Rp' + calc.cicilanBulanan + '/bln)',
        'Pending PIC Simpin'
      ]);
    }

    recordGasFinancialAudit(
      { nik: payload.createdByNik, nama: payload.createdByNama, role: payload.createdByRole },
      'PINJAMAN_KOPERASI',
      'APPLY_LOAN',
      loanId,
      {},
      { nomorSurat: nomorSurat, plafon: calc.plafon, tenor: calc.tenor }
    );

    return {
      status: 'success',
      success: true,
      message: 'Permohonan pinjaman ' + nomorSurat + ' berhasil diajukan.',
      loanId: loanId,
      nomorSurat: nomorSurat,
      simulation: calc
    };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getAllLoanRequests(role, status, search) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LoanRequests');
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', success: true, loans: [] };

    var data = sheet.getDataRange().getValues();
    var loans = [];

    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      var wf = r[24] || 'PENDING_PIC_SIMPIN';
      if (status && wf.toUpperCase() !== status.toUpperCase()) continue;

      var nomorSurat = r[1] || '';
      var nama = r[6] || '';
      var nik = r[5] || '';
      var kep = r[12] || '';

      if (search) {
        var q = search.toLowerCase();
        var match = nomorSurat.toLowerCase().indexOf(q) !== -1 ||
                    nama.toLowerCase().indexOf(q) !== -1 ||
                    String(nik).toLowerCase().indexOf(q) !== -1 ||
                    kep.toLowerCase().indexOf(q) !== -1;
        if (!match) continue;
      }

      loans.push({
        rowIndex: i + 1,
        loanId: r[0],
        nomorSurat: nomorSurat,
        timestamp: r[2],
        tanggalPermohonan: r[3],
        asalPengajuan: r[4],
        pemohonNik: nik,
        pemohonNama: nama,
        pemohonUnit: r[7],
        kantorCabang: r[8],
        penghasilanBulanan: r[9],
        alamat: r[10],
        noTelepon: r[11],
        keperluan: kep,
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
        statusWorkflow: wf,
        catatan: r[26] || '-'
      });
    }

    loans.reverse();
    return { status: 'success', success: true, loans: loans };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getMyLoanApplications(nik) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LoanRequests');
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', success: true, loans: [] };

    var data = sheet.getDataRange().getValues();
    var myLoans = [];

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][5] || '').trim() === String(nik).trim()) {
        myLoans.push({
          rowIndex: i + 1,
          loanId: data[i][0],
          nomorSurat: data[i][1],
          timestamp: data[i][2],
          pemohonNama: data[i][6],
          pemohonUnit: data[i][7],
          keperluan: data[i][12],
          jumlahPinjaman: parseInt(String(data[i][13] || '0').replace(/[^0-9]/g, ''), 10) || 0,
          tenorBulan: parseInt(data[i][14], 10) || 0,
          cicilanBulanan: parseInt(String(data[i][17] || '0').replace(/[^0-9]/g, ''), 10) || 0,
          statusWorkflow: data[i][24] || 'PENDING_PIC_SIMPIN',
          bankNama: data[i][20],
          bankNoRekening: data[i][21],
          bankAtasNama: data[i][22]
        });
      }
    }

    myLoans.reverse();
    return { status: 'success', success: true, loans: myLoans };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getLoanDocumentData(loanId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LoanRequests');
    if (!sheet) return { status: 'error', success: false, message: 'Sheet LoanRequests tidak ditemukan.' };

    var data = sheet.getDataRange().getValues();
    var loan = null;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === loanId || data[i][1] === loanId) {
        loan = data[i];
        break;
      }
    }

    if (!loan) return { status: 'error', success: false, message: 'Data pinjaman tidak ditemukan.' };

    var appSheet = ss.getSheetByName('LoanApprovals');
    var picApp = null;
    var ketuaApp = null;
    if (appSheet && appSheet.getLastRow() > 1) {
      var appData = appSheet.getDataRange().getValues();
      for (var a = 1; a < appData.length; a++) {
        if (appData[a][1] === loan[0]) {
          if (appData[a][3] === 'PIC_SIMPIN') picApp = appData[a];
          if (appData[a][3] === 'KETUA_KOPERASI') ketuaApp = appData[a];
        }
      }
    }

    var disbSheet = ss.getSheetByName('LoanDisbursements');
    var disb = null;
    if (disbSheet && disbSheet.getLastRow() > 1) {
      var disbData = disbSheet.getDataRange().getValues();
      for (var d = 1; d < disbData.length; d++) {
        if (disbData[d][1] === loan[0]) {
          disb = disbData[d];
          break;
        }
      }
    }

    var docPayload = {
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
          nama: '( ........................................ )',
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
          nama: disb ? disb[13] : 'Sri Wahyuningsih',
          jabatan: 'Bendahara / Keuangan',
          tanggal: disb ? disb[3] : '-',
          status: disb ? 'DISBURSED' : 'PENDING'
        }
      },
      statusWorkflow: loan[24],
      disbursement: disb ? {
        tanggal: disb[3],
        nominal: disb[4],
        noReferensi: disb[5],
        buktiUrl: disb[9]
      } : null
    };

    return { status: 'success', success: true, document: docPayload };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function reviewLoanPic(payload) {
  try {
    var role = payload.reviewerRole;
    if (!isRolePicSimpin(role) && !isRoleAdmin(role)) {
      return { status: 'error', success: false, message: 'Hanya PIC Simpin atau Administrator yang berwenang melakukan verifikasi ini.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LoanRequests');
    var data = sheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var targetRow = null;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.loanId) {
        targetRowIdx = i + 1;
        targetRow = data[i];
        break;
      }
    }

    if (targetRowIdx === -1) return { status: 'error', success: false, message: 'Pinjaman tidak ditemukan.' };

    var newStatus = 'PENDING_KETUA';
    var statusFR = 'Pending Ketua Koperasi';
    if (payload.action === 'REVISION') {
      newStatus = 'REVISION_REQUESTED';
      statusFR = 'Perlu Revisi (PIC Simpin)';
    } else if (payload.action === 'REJECT') {
      newStatus = 'REJECTED';
      statusFR = 'Rejected (PIC Simpin)';
    }

    sheet.getRange(targetRowIdx, 25).setValue(newStatus);
    if (payload.catatan) sheet.getRange(targetRowIdx, 27).setValue(payload.catatan);

    var appSheet = ss.getSheetByName('LoanApprovals');
    var jkt = getGasJakartaDateTime();
    appSheet.appendRow([
      'LA-PIC-' + new Date().getTime(),
      payload.loanId,
      jkt.timestampStr,
      'PIC_SIMPIN',
      payload.action === 'APPROVE' ? 'APPROVED' : (payload.action === 'REVISION' ? 'REVISION' : 'REJECTED'),
      payload.catatan || '-',
      String(payload.reviewerNik || ''),
      payload.reviewerNama || 'Lukman Harun Lubis',
      payload.reviewerRole || 'PIC Simpin'
    ]);

    // Sync FR
    var frSheet = ss.getSheetByName('Form Responses 1');
    if (frSheet && frSheet.getLastRow() > 1) {
      var frData = frSheet.getDataRange().getValues();
      var nomorSurat = targetRow[1];
      for (var f = 1; f < frData.length; f++) {
        if (String(frData[f][8] || '').indexOf(nomorSurat) !== -1) {
          frSheet.getRange(f + 1, 10).setValue(statusFR);
          break;
        }
      }
    }

    recordGasFinancialAudit(
      { nik: payload.reviewerNik, nama: payload.reviewerNama, role: payload.reviewerRole },
      'PINJAMAN_KOPERASI',
      'REVIEW_PIC_SIMPIN',
      payload.loanId,
      { status: 'PENDING_PIC_SIMPIN' },
      { status: newStatus, action: payload.action }
    );

    return { status: 'success', success: true, message: 'Verifikasi PIC Simpin berhasil diproses. Status: ' + newStatus };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function reviewLoanChairman(payload) {
  try {
    var role = payload.reviewerRole;
    if (!isRoleKetua(role) && !isRoleAdmin(role)) {
      return { status: 'error', success: false, message: 'Hanya Ketua Koperasi atau Administrator yang berwenang memberikan persetujuan pinjaman ini.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LoanRequests');
    var data = sheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var targetRow = null;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.loanId) {
        targetRowIdx = i + 1;
        targetRow = data[i];
        break;
      }
    }

    if (targetRowIdx === -1) return { status: 'error', success: false, message: 'Pinjaman tidak ditemukan.' };

    var newStatus = 'READY_FOR_DISBURSEMENT';
    var statusFR = 'Approved Ketua (Siap Cair)';
    if (payload.action === 'REVISION') {
      newStatus = 'REVISION_REQUESTED';
      statusFR = 'Perlu Revisi (Ketua Koperasi)';
    } else if (payload.action === 'REJECT') {
      newStatus = 'REJECTED';
      statusFR = 'Rejected (Ketua Koperasi)';
    }

    sheet.getRange(targetRowIdx, 25).setValue(newStatus);

    var appSheet = ss.getSheetByName('LoanApprovals');
    var jkt = getGasJakartaDateTime();
    appSheet.appendRow([
      'LA-KETUA-' + new Date().getTime(),
      payload.loanId,
      jkt.timestampStr,
      'KETUA_KOPERASI',
      payload.action === 'APPROVE' ? 'APPROVED' : (payload.action === 'REVISION' ? 'REVISION' : 'REJECTED'),
      payload.catatan || '-',
      String(payload.reviewerNik || ''),
      payload.reviewerNama || 'Fathi Mubarak',
      payload.reviewerRole || 'Ketua Koperasi'
    ]);

    // Sync FR
    var frSheet = ss.getSheetByName('Form Responses 1');
    if (frSheet && frSheet.getLastRow() > 1) {
      var frData = frSheet.getDataRange().getValues();
      var nomorSurat = targetRow[1];
      for (var f = 1; f < frData.length; f++) {
        if (String(frData[f][8] || '').indexOf(nomorSurat) !== -1) {
          frSheet.getRange(f + 1, 10).setValue(statusFR);
          break;
        }
      }
    }

    recordGasFinancialAudit(
      { nik: payload.reviewerNik, nama: payload.reviewerNama, role: payload.reviewerRole },
      'PINJAMAN_KOPERASI',
      'REVIEW_CHAIRMAN',
      payload.loanId,
      { status: 'PENDING_KETUA' },
      { status: newStatus, action: payload.action }
    );

    return { status: 'success', success: true, message: 'Persetujuan Ketua Koperasi berhasil diproses. Status: ' + newStatus };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function disburseLoan(payload) {
  try {
    var role = payload.disbursedByRole;
    if (!isRoleFinancial(role) && !isRoleAdmin(role)) {
      return { status: 'error', success: false, message: 'Hanya role Financial atau Administrator yang berwenang memproses pencairan pinjaman.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('LoanRequests');
    var data = sheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var targetRow = null;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.loanId) {
        targetRowIdx = i + 1;
        targetRow = data[i];
        break;
      }
    }

    if (targetRowIdx === -1) return { status: 'error', success: false, message: 'Pinjaman tidak ditemukan.' };

    sheet.getRange(targetRowIdx, 25).setValue('DISBURSED');

    var disbSheet = ss.getSheetByName('LoanDisbursements');
    var jkt = getGasJakartaDateTime();
    var disbId = 'DISB-' + new Date().getTime();

    disbSheet.appendRow([
      disbId,
      payload.loanId,
      jkt.timestampStr,
      payload.tanggalPencairan || jkt.dateStr,
      payload.nominalPencairan || targetRow[13],
      payload.nomorReferensi,
      payload.bankTujuan || targetRow[20],
      String(payload.noRekeningTujuan || targetRow[21]),
      payload.atasNamaTujuan || targetRow[22],
      payload.buktiUrl || '-',
      payload.catatan || '-',
      'COMPLETED',
      String(payload.disbursedByNik || ''),
      payload.disbursedByNama || 'Sri Wahyuningsih'
    ]);

    // Sync FR Done
    var frSheet = ss.getSheetByName('Form Responses 1');
    if (frSheet && frSheet.getLastRow() > 1) {
      var frData = frSheet.getDataRange().getValues();
      var nomorSurat = targetRow[1];
      for (var f = 1; f < frData.length; f++) {
        if (String(frData[f][8] || '').indexOf(nomorSurat) !== -1) {
          frSheet.getRange(f + 1, 10).setValue('Done');
          break;
        }
      }
    }

    recordGasFinancialAudit(
      { nik: payload.disbursedByNik, nama: payload.disbursedByNama, role: payload.disbursedByRole },
      'PINJAMAN_KOPERASI',
      'DISBURSE_LOAN',
      payload.loanId,
      { status: 'READY_FOR_DISBURSEMENT' },
      { status: 'DISBURSED', nomorReferensi: payload.nomorReferensi }
    );

    return { status: 'success', success: true, message: 'Pencairan pinjaman berhasil dicatat dengan Nomor Referensi: ' + payload.nomorReferensi };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getMyBankAccounts(nik, username) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('MemberBankAccounts');
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', success: true, accounts: [] };

    var data = sheet.getDataRange().getValues();
    var accounts = [];

    var qNik = String(nik || '').trim().toLowerCase();
    var qUser = String(username || '').trim().toLowerCase();

    var extractBase = function(val) {
      if (!val) return '';
      var s = String(val).trim().toLowerCase();
      var m = s.match(/^([a-z0-9]+)/i);
      return m ? m[1] : s;
    };
    var qBase = extractBase(qNik);

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowStatus = String(row[11] || '').toUpperCase().trim();
      if (rowStatus !== 'ACTIVE') continue;

      var rowNik = String(row[1] || '').trim().toLowerCase();
      var isMatch = false;

      if (qNik && rowNik === qNik) {
        isMatch = true;
      } else {
        var rowBase = extractBase(rowNik);
        if (qBase && rowBase && qBase === rowBase && qBase.length >= 4) {
          isMatch = true;
        } else if (qUser && rowNik === qUser) {
          isMatch = true;
        }
      }

      if (isMatch) {
        accounts.push({
          accountId: row[0],
          nik: row[1],
          namaAnggota: row[2],
          namaBank: row[3],
          nomorRekening: row[4],
          atasNamaRekening: row[5],
          statusVerifikasi: row[6] || 'UNVERIFIED',
          tanggalVerifikasi: row[7] || '-'
        });
      }
    }

    return { status: 'success', success: true, accounts: accounts };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function requestBankAccountChange(payload) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('BankAccountChangeRequests');
    var reqId = 'REQ-BANK-' + new Date().getTime();
    var jkt = getGasJakartaDateTime();

    sheet.appendRow([
      reqId,
      String(payload.nik).trim(),
      payload.namaAnggota,
      payload.bankLama || '-',
      String(payload.rekeningLama || '-'),
      payload.atasNamaLama || '-',
      payload.bankBaru,
      String(payload.rekeningBaru),
      payload.atasNamaBaru,
      payload.alasan,
      payload.buktiUrl || '-',
      'REQUESTED',
      '-',
      '',
      '',
      ''
    ]);

    recordGasFinancialAudit(
      { nik: payload.nik, nama: payload.namaAnggota, role: 'Anggota' },
      'BANK_ACCOUNTS',
      'REQUEST_CHANGE',
      reqId,
      { rekeningLama: payload.rekeningLama },
      { rekeningBaru: payload.rekeningBaru }
    );

    return { status: 'success', success: true, message: 'Permohonan perubahan rekening berhasil diajukan.', reqId: reqId };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getBankAccountRequests(role) {
  try {
    if (!isRoleAdmin(role) && !isRoleKetua(role)) return { status: 'error', success: false, message: 'Akses ditolak. Khusus Admin atau Ketua Koperasi.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('BankAccountChangeRequests');
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', success: true, requests: [] };

    var data = sheet.getDataRange().getValues();
    var requests = [];

    for (var i = 1; i < data.length; i++) {
      requests.push({
        rowIndex: i + 1,
        requestId: data[i][0],
        nik: data[i][1],
        namaAnggota: data[i][2],
        bankLama: data[i][3],
        rekeningLama: data[i][4],
        atasNamaLama: data[i][5],
        bankBaru: data[i][6],
        rekeningBaru: data[i][7],
        atasNamaBaru: data[i][8],
        alasan: data[i][9],
        buktiUrl: data[i][10],
        status: data[i][11] || 'REQUESTED',
        catatanAdmin: data[i][12] || '-'
      });
    }

    requests.reverse();
    return { status: 'success', success: true, requests: requests };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function reviewBankAccountRequest(payload) {
  try {
    if (!isRoleAdmin(payload.adminRole) && !isRoleKetua(payload.adminRole)) return { status: 'error', success: false, message: 'Hanya Admin atau Ketua Koperasi yang berhak meninjau.' };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('BankAccountChangeRequests');
    var data = sheet.getDataRange().getValues();
    var targetRowIdx = -1;
    var bReq = null;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.reqId) {
        targetRowIdx = i + 1;
        bReq = data[i];
        break;
      }
    }

    if (targetRowIdx === -1) return { status: 'error', success: false, message: 'Request tidak ditemukan.' };

    var newStatus = payload.action === 'APPROVE' ? 'APPROVED' : 'REJECTED';
    var jkt = getGasJakartaDateTime();

    sheet.getRange(targetRowIdx, 12, 1, 5).setValues([[
      newStatus,
      payload.catatan || '-',
      String(payload.adminNik || ''),
      payload.adminNama || 'Administrator',
      jkt.timestampStr
    ]]);

    if (payload.action === 'APPROVE') {
      var bankSheet = ss.getSheetByName('MemberBankAccounts');
      if (bankSheet && bankSheet.getLastRow() > 1) {
        var bData = bankSheet.getDataRange().getValues();
        for (var b = 1; b < bData.length; b++) {
          if (String(bData[b][1]).trim() === String(bReq[1]).trim()) {
            bankSheet.getRange(b + 1, 12).setValue('INACTIVE');
          }
        }
      }
      if (bankSheet) {
        bankSheet.appendRow([
          'ACC-' + new Date().getTime(),
          String(bReq[1]).trim(),
          bReq[2],
          bReq[6],
          String(bReq[7]),
          bReq[8],
          'VERIFIED',
          jkt.timestampStr,
          String(payload.adminNik || ''),
          payload.adminNama || 'Administrator',
          bReq[10] || '-',
          'ACTIVE',
          jkt.timestampStr
        ]);
      }
    }

    recordGasFinancialAudit(
      { nik: payload.adminNik, nama: payload.adminNama, role: payload.adminRole },
      'BANK_ACCOUNTS',
      'REVIEW_CHANGE_REQUEST',
      payload.reqId,
      { status: 'REQUESTED' },
      { status: newStatus }
    );

    return { status: 'success', success: true, message: 'Permohonan perubahan rekening ' + newStatus + '.' };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function getSimpananApprovals() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SimpananSukarelaApprovals');
    if (!sheet || sheet.getLastRow() <= 1) return { status: 'success', success: true, approvals: [] };

    var data = sheet.getDataRange().getValues();
    var approvals = [];
    for (var i = 1; i < data.length; i++) {
      var r = data[i];
      approvals.push({
        rowIndex: i + 1,
        approvalId: r[0],
        ruleId: r[1],
        timestamp: r[2],
        nik: r[3],
        namaAnggota: r[4],
        tipePerubahan: r[5],
        lama: r[6],
        baru: r[7],
        diajukanOlehNama: r[9],
        status: r[10] || 'PENDING',
        catatanReview: r[11] || '-'
      });
    }
    approvals.reverse();
    return { status: 'success', success: true, approvals: approvals };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}

function reviewSimpananApproval(payload) {
  try {
    var role = payload.reviewerRole;
    if (!isRoleKetua(role) && !isRoleAdmin(role)) {
      return { status: 'error', success: false, message: 'Hanya Ketua Koperasi atau Administrator yang berwenang meninjau usulan auto-input.' };
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('SimpananSukarelaApprovals');
    if (!sheet) return { status: 'error', success: false, message: 'Sheet tidak ditemukan.' };

    var data = sheet.getDataRange().getValues();
    var targetIdx = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === payload.approvalId) {
        targetIdx = i + 1;
        break;
      }
    }

    if (targetIdx === -1) return { status: 'error', success: false, message: 'Approval tidak ditemukan.' };

    var appRow = data[targetIdx - 1];
    var ruleId = appRow[1];
    var statusResult = payload.action === 'APPROVE' ? 'APPROVED' : 'REJECTED';
    var jkt = getGasJakartaDateTime();

    sheet.getRange(targetIdx, 11, 1, 5).setValues([[
      statusResult,
      payload.catatan || '-',
      String(payload.reviewerNik || ''),
      payload.reviewerNama || '',
      jkt.timestampStr
    ]]);

    if (payload.action === 'APPROVE') {
      var baruObj = {};
      try { baruObj = JSON.parse(appRow[7] || '{}'); } catch(e){}

      var rulesSheet = ss.getSheetByName('SimpananSukarelaRules');
      if (rulesSheet && rulesSheet.getLastRow() > 1) {
        var rData = rulesSheet.getDataRange().getValues();
        for (var j = 1; j < rData.length; j++) {
          if (rData[j][0] === ruleId) {
            if (baruObj.nominalRutin) {
              rulesSheet.getRange(j + 1, 6).setValue(baruObj.nominalRutin);
            }
            rulesSheet.getRange(j + 1, 11).setValue('ACTIVE');
            break;
          }
        }
      }
    }

    return { status: 'success', success: true, message: 'Usulan auto-input berhasil diproses: ' + statusResult };
  } catch (err) {
    return { status: 'error', success: false, message: err.toString() };
  }
}
