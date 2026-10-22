function doGet(e) {
  setupAllSheets();
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Koperasi Satya Ardhia - Sistem Terpadu & Presensi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

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
  if (requestingUserRole !== 'Admin') {
    return []; // Tolak akses jika bukan admin
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
      nama: row[3] || '-',
      nik: row[4] || '-',
      email: row[5] || '-',
      accountStatus: row[7] ? String(row[7]).trim() : (row[2] === 'Admin' ? 'Active' : 'Pending')
    };
  });
}

// 4. Update Status Akun User oleh Admin
function updateUserAccountStatus(requestingUserRole, rowIndex, newStatus) {
  if (requestingUserRole !== 'Admin') {
    return { status: 'error', message: 'Akses ditolak! Hanya Admin yang memiliki hak ini.' };
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
      tipeFormulir: row[1] || 'Pengajuan Biaya',
      namaKaryawan: row[2] || '-',
      nikUnit: row[3] || '-',
      noRekening: row[4] ? String(row[4]) : '-',
      atasNamaRekening: row[5] || '-',
      totalDitransfer: Number(row[6]) || 0,
      cicilan: row[7] || '1',
      keterangan: row[8] || '-',
      status: row[9] ? String(row[9]).trim() : 'Created'
    };
  }).reverse();

  if (userRole === 'User') {
    const cleanNik = String(userNik || '').trim().toLowerCase();
    const cleanName = String(userName || '').trim().toLowerCase();
    const cleanUname = String(userUsername || '').trim().toLowerCase();

    // Ekstrak angka unik jika NIK berupa nomor identitas (misal: 20246808)
    const nikMatch = cleanNik.match(/\d{4,}/);
    const nikOnly = nikMatch ? nikMatch[0] : '';

    list = list.filter(item => {
      const rowNama = String(item.namaKaryawan || '').trim().toLowerCase();
      const rowNikUnit = String(item.nikUnit || '').trim().toLowerCase();

      // 1. Prioritaskan kecocokan Nama Lengkap
      const isNameMatch = Boolean(
        cleanName && (rowNama === cleanName || rowNama.includes(cleanName) || cleanName.includes(rowNama))
      );

      // 2. Kecocokan Username
      const isUnameMatch = Boolean(
        cleanUname && (rowNama.includes(cleanUname) || rowNikUnit.includes(cleanUname))
      );

      // 3. Kecocokan NIK angka unik
      let isNikMatch = false;
      if (nikOnly) {
        isNikMatch = rowNikUnit.includes(nikOnly);
      } else if (cleanNik && !['karyawan koperasi', 'koperasi', 'unit', 'staff', 'avsec', 'arff', '-'].includes(cleanNik)) {
        isNikMatch = rowNikUnit === cleanNik;
      }

      if (isNameMatch || isUnameMatch) return true;
      if (isNikMatch && nikOnly) return true;
      return false;
    });
  }

  return list;
}

function simpanPengajuan(form) {
  try {
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
  // Validasi: User biasa tidak diizinkan mengubah status
  if (userRole === 'User') {
    return { status: 'error', message: 'Akses ditolak! Anda tidak memiliki wewenang mengubah status.' };
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
function registerFaceDescriptor(nik, nama, descriptorJson) {
  try {
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

    // Status kehadiran (Jam masuk kantor normal 08:30 WIB)
    const hours = now.getHours();
    const minutes = now.getMinutes();
    const totalMinutes = hours * 60 + minutes;
    const statusKehadiran = totalMinutes <= (8 * 60 + 30) ? 'Hadir Tepat Waktu' : 'Terlambat';

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

    // Update kolom Clock-Out (Kolom 7 / G)
    sheet.getRange(foundRow, 7).setValue(timeStr);

    return {
      status: 'success',
      message: 'Clock-Out berhasil dicatat pada pukul ' + timeStr + ' WIB. Terima kasih atas kerja keras hari ini!',
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
function getAbsensiHistoryList(userRole, nik, nama) {
  try {
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
    })).reverse();

    if (userRole === 'User') {
      const cleanNik = String(nik || '').trim().toLowerCase();
      const cleanNama = String(nama || '').trim().toLowerCase();
      list = list.filter(item => {
        const itemNik = String(item.nik || '').toLowerCase();
        const itemNama = String(item.nama || '').toLowerCase();
        return (cleanNik && itemNik.includes(cleanNik)) || (cleanNama && itemNama.includes(cleanNama));
      });
    }

    return list;
  } catch (err) {
    return [];
  }
}
