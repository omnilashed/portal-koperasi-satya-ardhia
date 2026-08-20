function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Sistem Pengajuan & Monitoring Koperasi')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
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

    list = list.filter(item => {
      const rowNama = String(item.namaKaryawan || '').trim().toLowerCase();
      const rowNikUnit = String(item.nikUnit || '').trim().toLowerCase();

      const matchNik = cleanNik && (rowNikUnit.includes(cleanNik) || rowNama.includes(cleanNik));
      const matchUsername = cleanUname && (rowNikUnit.includes(cleanUname) || rowNama.includes(cleanUname));
      const matchName = cleanName && (rowNama.includes(cleanName) || cleanName.includes(rowNama));

      return matchNik || matchUsername || matchName;
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
}             "• `/budget set <kategori> <limit>`\n" +
             "• `/budget list` (Melihat status & progress bar)\n\n" +
             "4️⃣ *Utang & Piutang:*\n" +
             "• `/utang add -m \"<nama> - <ket>\" <nominal> [-tempo YYYY-MM-DD]`\n" +
             "• `/piutang add -m \"<nama> - <ket>\" <nominal> [-tempo YYYY-MM-DD]`\n" +
             "• `/utang list` | `/utang lunas <ID>` | `/piutang lunas <ID>`\n\n" +
             "5️⃣ *Laporan & Saldo:*\n" +
             "• `/summary` | `/summary hari-ini` | `/summary kemarin` | `/summary YYYY-MM` | `/summary -c <kategori>`\n" +
             "• `/saldo` (Lihat saldo semua rekening)\n\n" +
             "6️⃣ *AI Features:*\n" +
             "• `/tanya <pertanyaan>` (Financial Advisor AI)\n" +
             "• Pesan teks bebas (misal: *\"barusan beli kopi 25rb pake gopay\"*) langsung dicatat otomatis oleh AI!";

  sendTelegramMessage(chatId, help, msgId);
}
