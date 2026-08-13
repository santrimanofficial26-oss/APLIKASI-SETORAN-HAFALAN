// Helper Dapatkan Spreadsheet Aktif (Membaca langsung Google Sheet tempat Script ini terpasang)
function getDbSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error("Spreadsheet aktif tidak ditemukan. Pastikan script ini dibuka via menu Ekstensi -> Apps Script di dalam Google Sheets Anda.");
  }
  return ss;
}

function doGet(e) {
  try {
    setupDatabase();
  } catch (err) {
    // Abaikan jika spreadsheet belum terhubung saat pertama kali render
  }
  
  var htmlOutput;
  try {
    htmlOutput = HtmlService.createHtmlOutputFromFile('index');
  } catch (err) {
    htmlOutput = HtmlService.createHtmlOutputFromFile('Index');
  }
  
  return htmlOutput
      .setTitle('PENA-Q - Tahsin & Tahfidz')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Inisialisasi Struktur Sheet Google Spreadsheet & Auto Seed Data Awal
function setupDatabase() {
  var ss = getDbSpreadsheet();
  var sheets = [
    { 
      name: 'Guru', 
      headers: ['ID Guru', 'Nama', 'Username', 'Password', 'Role', 'Foto'],
      defaultData: [
        ['ADM01', 'Administrator Sekolah', 'admin', '@penawa123', 'Admin', ''],
        ['G01', 'Rinna Ardiana, S.Pd', 'rinna', 'rinna123', 'Guru', ''],
        ['G02', 'M. SANDISI, S.Pd', 'sandisi', 'sandisi123', 'Guru', ''],
        ['G03', 'Ustadz Mansur, S.Pd.I', 'mansur', 'mansur123', 'Guru', ''],
        ['G04', 'Ustadzah Fatimah, S.Ag', 'fatimah', 'fatimah123', 'Guru', '']
      ]
    },
    { 
      name: 'Siswa', 
      headers: ['NIS', 'Nama', 'Kelas', 'Kelompok', 'Guru', 'Status', 'Foto'],
      defaultData: [
        ['20267001', 'Salsa Apriliani', '9B', 'Tahsin Al-Qur\'an', 'Rinna Ardiana, S.Pd', 'Aktif', ''],
        ['20267002', 'Syauqi Adnan', '8D', 'Tahsin Al-Qur\'an', 'M. SANDISI, S.Pd', 'Aktif', ''],
        ['20267003', 'Ahmad Zaki Mubarok', '7A', 'Tahsin Iqro', 'Ustadz Mansur, S.Pd.I', 'Aktif', ''],
        ['20267004', 'Fatimah Az-Zahra', '8A', 'Tahfidz', 'Ustadzah Fatimah, S.Ag', 'Aktif', '']
      ]
    },
    { name: 'Tahsin', headers: ['Tanggal', 'NIS', 'Nama', 'Iqro/Al-Quran', 'Halaman', 'Status', 'Catatan', 'Guru'] },
    { name: 'Tahfidz', headers: ['Tanggal', 'NIS', 'Nama', 'Surat', 'Ayat', 'Nilai', 'Catatan', 'Guru'] },
    { 
      name: 'Rekap', 
      headers: ['Tanggal', 'NIS', 'Nama', 'Program', 'Detail', 'Status/Nilai', 'Catatan', 'Guru'],
      defaultData: [
        ['2026-07-21', '20267001', 'Salsa Apriliani', 'Tahsin Al-Qur\'an', 'Al-Qur\'an Juz 1 Hal 10', 'Lulus ⭐⭐⭐⭐⭐', 'Membaca sangat lancar, makhorijul huruf tepat', 'Rinna Ardiana, S.Pd'],
        ['2026-07-21', '20267002', 'Syauqi Adnan', 'Tahsin Al-Qur\'an', 'Al-Qur\'an Juz 1 Hal 12', 'Lulus ⭐⭐⭐⭐⭐', 'Tajwid & mad thabi\'i sudah baik, tingkatkan murojaah', 'M. SANDISI, S.Pd']
      ]
    },
    { 
      name: 'Pengaturan', 
      headers: ['Kunci', 'Nilai'],
      defaultData: [
        ['schoolName', 'SMP Pesantren Nahdlatul Wathon'],
        ['appTitle', 'PENA-Q'],
        ['logoUrl', ''],
        ['academicYear', '2026/2027']
      ]
    },
    {
      name: 'Students',
      headers: ['Student ID', 'Full Name', 'NISN', 'Class', 'Group', 'Pembina ID', 'Status', 'Attendance Percentage', 'Performance Points', 'Contact Number', 'Parent Contact', 'Last Updated']
    },
    {
      name: 'Pembina',
      headers: ['Pembina ID', 'Full Name', 'Assigned Group', 'Email', 'Role'],
      defaultData: [
        ['G01', 'Rinna Ardiana, S.Pd', 'Tahsin Al-Qur\'an', 'rinna@penawa.sch.id', 'Pembina'],
        ['G02', 'M. SANDISI, S.Pd', 'Tahsin Al-Qur\'an', 'sandisi@penawa.sch.id', 'Pembina'],
        ['G03', 'Ustadz Mansur, S.Pd.I', 'Tahsin Iqro', 'mansur@penawa.sch.id', 'Pembina Utama'],
        ['G04', 'Ustadzah Fatimah, S.Ag', 'Tahfidz', 'fatimah@penawa.sch.id', 'Pembina']
      ]
    },
    {
      name: 'Activity_Logs',
      headers: ['Log ID', 'Timestamp', 'Pembina ID', 'Action Taken', 'Target Student ID']
    }
  ];

  sheets.forEach(function(sh) {
    var sheet = ss.getSheetByName(sh.name);
    if (!sheet) {
      sheet = ss.insertSheet(sh.name);
    }
    // Overwrite/update header row (row 1)
    sheet.getRange(1, 1, 1, sh.headers.length).setValues([sh.headers]);
    sheet.getRange(1, 1, 1, sh.headers.length).setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    
    // Isi data default jika sheet masih kosong (hanya header)
    if (sh.defaultData && sheet.getLastRow() <= 1) {
      sh.defaultData.forEach(function(rowVal) {
        sheet.appendRow(rowVal);
      });
    }
  });

  // Pastikan data kelas Salsa Apriliani diperbarui ke 9B di Google Sheets
  var siswaSheet = ss.getSheetByName('Siswa');
  if (siswaSheet && siswaSheet.getLastRow() > 1) {
    var sValues = siswaSheet.getDataRange().getValues();
    for (var sIdx = 1; sIdx < sValues.length; sIdx++) {
      var nisVal = String(sValues[sIdx][0] || '').trim();
      var namaVal = String(sValues[sIdx][1] || '').trim().toLowerCase();
      if (nisVal === '20267001' || namaVal.indexOf('salsa') !== -1) {
        siswaSheet.getRange(sIdx + 1, 3).setValue('9B');
      }
    }
  }

  // Seed 245 data siswa otomatis jika sheet Students masih kosong
  var stdSheet = ss.getSheetByName('Students');
  if (stdSheet && stdSheet.getLastRow() <= 1) {
    var firstNames = ['Ahmad', 'Aisyah', 'Bagas', 'Bintang', 'Citra', 'Daffa', 'Dimas', 'Eka', 'Fadli', 'Fatimah', 'Ghanim', 'Hafiz', 'Ibrahim', 'Indah', 'Jasmine', 'Kafi', 'Laila', 'Maulana', 'Nabila', 'Omar', 'Putri', 'Rafi', 'Rizky', 'Salsa', 'Syauqi', 'Tariq', 'Umar', 'Vina', 'Wahyu', 'Zahra'];
    var lastNames = ['Adnan', 'Akbar', 'Amalia', 'Ardiansyah', 'Az-Zahra', 'Fauzi', 'Hidayat', 'Kusuma', 'Mubarok', 'Nugraha', 'Pratama', 'Putra', 'Ramadhan', 'Saputra', 'Setyawan', 'Utami', 'Wibowo', 'Wijaya', 'Yuliana', 'Zulkarnain'];
    var classesList = ['7A', '7B', '7C', '8A', '8B', '8C', '9A', '9B', '9C'];
    var groupsList = ['Tahsin Iqro', 'Tahsin Al-Qur\'an', 'Tahfidz', 'Hadrah', 'Pramuka', 'English Club'];
    var pembinasList = ['G01', 'G02', 'G03', 'G04'];
    var statusList = ['Active', 'Active', 'Active', 'Active', 'Inactive', 'On Leave'];

    var rowsToInsert = [];
    var baseNisn = 30890001;
    var basePhone = 628123456000;

    for (var i = 1; i <= 245; i++) {
      var fn = firstNames[i % firstNames.length];
      var ln = lastNames[i % lastNames.length];
      var fullName = (fn + ' ' + ln + ' ' + (i % 3 === 0 ? 'M.' : '')).trim();
      var stdId = 'STD' + (1000 + i);
      var nisnVal = String(baseNisn + i);
      var cVal = classesList[i % classesList.length];
      var gVal = groupsList[i % groupsList.length];
      var pVal = pembinasList[i % pembinasList.length];
      var stVal = statusList[i % statusList.length];

      var att = Math.floor(60 + (i * 17) % 41);
      var perf = Math.floor(50 + (i * 23) % 51);
      var phone = '+' + (basePhone + i);
      var parentPhone = '+' + (basePhone + i + 1000);
      var dayStr = (i % 28) + 1;
      var updateDate = '2026-07-' + (dayStr < 10 ? '0' + dayStr : dayStr) + ' 10:00:00';

      rowsToInsert.push([stdId, fullName, nisnVal, cVal, gVal, pVal, stVal, att, perf, phone, parentPhone, updateDate]);
    }

    if (rowsToInsert.length > 0) {
      stdSheet.getRange(2, 1, rowsToInsert.length, 12).setValues(rowsToInsert);
    }
  }
}

// Keep setupSpreadsheet for backward compatibility
function setupSpreadsheet() {
  setupDatabase();
}

// Helper Hashing SHA-256 untuk Password
function hashPassword(password) {
  if (!password) return '';
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(password), Utilities.Charset.UTF_8);
  var hex = '';
  for (var i = 0; i < digest.length; i++) {
    var byteVal = digest[i];
    if (byteVal < 0) byteVal += 256;
    var byteHex = byteVal.toString(16);
    if (byteHex.length === 1) byteHex = '0' + byteHex;
    hex += byteHex;
  }
  return hex;
}

// Helper Pembersihan Nama Guru Backend (Menghapus Gelar & Kata Umun seperti Guru, Pembina, Ustadz)
function cleanTeacherNameForMatching(s) {
  if (!s) return '';
  return String(s).toLowerCase()
    .replace(/s\s*\.?\s*pd\s*\.?\s*i/gi, '')
    .replace(/s\s*\.?\s*pd/gi, '')
    .replace(/s\s*\.?\s*ag/gi, '')
    .replace(/m\s*\.?\s*pd/gi, '')
    .replace(/m\s*\.?\s*ag/gi, '')
    .replace(/s\s*\.?\s*kom/gi, '')
    .replace(/s\s*\.?\s*e/gi, '')
    .replace(/s\s*\.?\s*t/gi, '')
    .replace(/ustadzah?/gi, '')
    .replace(/guru/gi, '')
    .replace(/pembina/gi, '')
    .replace(/pembimbing/gi, '')
    .replace(/pengajar/gi, '')
    .replace(/sekolah/gi, '')
    .replace(/admin(istrator)?/gi, '')
    .replace(/pesantren/gi, '')
    .replace(/nahdlatul/gi, '')
    .replace(/wathon/gi, '')
    .replace(/smp/gi, '')
    .replace(/tahsin/gi, '')
    .replace(/tahfidz/gi, '')
    .replace(/iqro/gi, '')
    .replace(/quran/gi, '')
    .replace(/drs/gi, '')
    .replace(/dra/gi, '')
    .replace(/hj/gi, '')
    .replace(/kh/gi, '')
    .replace(/[^a-z0-9]/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function getTeacherTokensBackend(s) {
  var cleaned = cleanTeacherNameForMatching(s);
  if (!cleaned) return [];
  return cleaned.split(' ').filter(function(t) { return t.length >= 3; });
}

// Fungsi Backend Terisolasi Pengambilan Data Siswa Binaan
function getAssignedStudents(teacherIdentifier) {
  try {
    setupDatabase();
    if (!teacherIdentifier) {
      return serializeForClient({ ok: false, message: 'Identitas guru pembina tidak valid', data: [] });
    }
    
    var ss = getDbSpreadsheet();

    // Resolusi Teacher ID ke Name & Username jika teacherIdentifier berupa ID (misal: G01, G02, P05)
    var resolvedTeacherName = String(teacherIdentifier).trim();
    var resolvedTeacherUser = '';
    var guruSheet = ss.getSheetByName('Guru');
    if (guruSheet) {
      var gData = guruSheet.getDataRange().getValues();
      for (var gIdx = 1; gIdx < gData.length; gIdx++) {
        var gId = String(gData[gIdx][0] || '').trim().toLowerCase();
        if (gId === String(teacherIdentifier).trim().toLowerCase()) {
          resolvedTeacherName = String(gData[gIdx][1] || '').trim();
          resolvedTeacherUser = String(gData[gIdx][2] || '').trim();
          break;
        }
      }
    }

    var sheet = ss.getSheetByName('Siswa');
    if (!sheet) {
      return serializeForClient({ ok: false, message: 'Sheet Siswa tidak ditemukan', data: [] });
    }

    var data = sheet.getDataRange().getValues();
    var assigned = [];

    var targetCleanName = cleanTeacherNameForMatching(resolvedTeacherName);
    var targetCleanId = cleanTeacherNameForMatching(teacherIdentifier);
    var targetCleanUser = cleanTeacherNameForMatching(resolvedTeacherUser);

    var targetTokens = getTeacherTokensBackend(resolvedTeacherName);
    if (resolvedTeacherUser) {
      targetTokens = targetTokens.concat(getTeacherTokensBackend(resolvedTeacherUser));
    }

    for (var i = 1; i < data.length; i++) {
      if (!data[i][0]) continue;
      var sNis = String(data[i][0]).trim();
      var sNama = String(data[i][1]).trim();
      var sKelas = String(data[i][2]).trim();
      var sKelompok = String(data[i][3]).trim();
      var sGuru = String(data[i][4]).trim();
      var sStatus = String(data[i][5] || 'Aktif').trim();
      var sFoto = String(data[i][6] || '').trim();

      var sGuruClean = cleanTeacherNameForMatching(sGuru);
      var isMatch = false;

      // Match presisi: Nama Bersih, ID, Username, atau Token Spesifik
      if (sGuruClean && targetCleanName && sGuruClean === targetCleanName) {
        isMatch = true;
      } else if (sGuruClean && targetCleanId && sGuruClean === targetCleanId) {
        isMatch = true;
      } else if (targetCleanUser && sGuruClean && sGuruClean === targetCleanUser) {
        isMatch = true;
      } else if (sGuru.toLowerCase() === String(teacherIdentifier).toLowerCase() || sGuru.toLowerCase() === String(resolvedTeacherName).toLowerCase()) {
        isMatch = true;
      } else {
        var sTokens = getTeacherTokensBackend(sGuru);
        if (sTokens.length > 0 && targetTokens.length > 0) {
          if (sTokens.some(function(st) { return targetTokens.some(function(tt) { return st === tt; }); })) {
            isMatch = true;
          }
        }
      }

      if (isMatch) {
        assigned.push({
          nis: sNis,
          name: sNama,
          kelas: sKelas,
          kelompok: sKelompok,
          guru: sGuru,
          status: sStatus,
          foto: sFoto,
          level: sKelompok.toLowerCase().indexOf('tahfidz') !== -1 ? 'Tahfidz' : (sKelompok.toLowerCase().indexOf('qur') !== -1 ? 'Tahsin Al-Qur\'an' : 'Tahsin Iqro')
        });
      }
    }

    return serializeForClient({
      ok: true,
      message: 'Berhasil mengambil ' + assigned.length + ' siswa binaan',
      data: assigned
    });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString(), data: [] });
  }
}

// Helper untuk validasi kepemilikan siswa oleh guru di backend
function isStudentAssignedToTeacher(nis, guruIdentifier) {
  if (!nis || !guruIdentifier) return false;
  try {
    var ss = getDbSpreadsheet();

    var targetTeacherName = String(guruIdentifier).trim();
    var targetTeacherUser = '';
    var guruSheet = ss.getSheetByName('Guru');
    if (guruSheet) {
      var gData = guruSheet.getDataRange().getValues();
      for (var gIdx = 1; gIdx < gData.length; gIdx++) {
        var gId = String(gData[gIdx][0] || '').trim().toLowerCase();
        if (gId === String(guruIdentifier).trim().toLowerCase()) {
          targetTeacherName = String(gData[gIdx][1] || '').trim();
          targetTeacherUser = String(gData[gIdx][2] || '').trim();
          break;
        }
      }
    }

    var sheet = ss.getSheetByName('Siswa');
    if (!sheet) return false;
    var data = sheet.getDataRange().getValues();
    var cleanNis = String(nis).trim().toLowerCase();

    var targetGuruClean = cleanTeacherNameForMatching(targetTeacherName);
    var targetIdClean = cleanTeacherNameForMatching(guruIdentifier);
    var targetUserClean = cleanTeacherNameForMatching(targetTeacherUser);
    var tTokens = getTeacherTokensBackend(targetTeacherName);

    for (var i = 1; i < data.length; i++) {
      var sNis = String(data[i][0] || '').trim().toLowerCase();
      if (sNis === cleanNis) {
        var sGuru = String(data[i][4] || '').trim();
        var sGuruClean = cleanTeacherNameForMatching(sGuru);
        if (sGuruClean === targetGuruClean || 
            sGuruClean === targetIdClean || 
            (targetUserClean && sGuruClean === targetUserClean) ||
            sGuru.toLowerCase() === String(targetTeacherName).toLowerCase() ||
            sGuru.toLowerCase() === String(guruIdentifier).toLowerCase()) {
          return true;
        }
        var sTokens = getTeacherTokensBackend(sGuru);
        if (sTokens.length > 0 && tTokens.length > 0) {
          if (sTokens.some(function(st) { return tTokens.some(function(tt) { return st === tt; }); })) {
            return true;
          }
        }
        return false;
      }
    }
  } catch (err) {
    return false;
  }
  return false;
}

// Helper to recursively serialize spreadsheet data for client
function serializeForClient(value) {
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (Array.isArray(value)) {
    return value.map(serializeForClient);
  }
  if (value !== null && typeof value === 'object') {
    var copy = {};
    for (var key in value) {
      if (value.hasOwnProperty(key)) {
        copy[key] = serializeForClient(value[key]);
      }
    }
    return copy;
  }
  return value;
}

// Fungsi Simpan Setoran Tahsin
function saveTahsinToSheet(data) {
  try {
    var ss = getDbSpreadsheet();

    // Validasi Keamanan Backend: Pastikan siswa adalah binaan guru yang bersangkutan
    if (data.guru && data.guru !== 'Administrator Sekolah' && !isStudentAssignedToTeacher(data.nis, data.guru)) {
      return serializeForClient({ ok: false, message: 'Akses Ditolak: Siswa ' + (data.nama || data.nis) + ' bukan merupakan siswa binaan ' + data.guru });
    }
    // Save Tahsin backend validation response
    var sheet = ss.getSheetByName('Tahsin');
    var detailText = data.detail || (data.level + (data.halaman ? ' ' + data.halaman : ''));
    sheet.appendRow([data.tanggal, data.nis, data.nama, data.level, data.halaman || data.detail || '', data.status, data.catatan, data.guru]);
    
    // Sinkron ke Rekap dengan Program Tahsin yang Akurat
    var programName = data.program || 'Tahsin';
    var rekapSheet = ss.getSheetByName('Rekap');
    rekapSheet.appendRow([data.tanggal, data.nis, data.nama, programName, detailText, data.status, data.catatan || '-', data.guru]);
    
    return serializeForClient({ 
      ok: true, 
      message: 'Setoran Tahsin ' + data.nama + ' berhasil disimpan',
      data: { nis: data.nis, guru: data.guru } 
    });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Hapus Seluruh Guru
function deleteAllGuruFromSheet() {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Guru');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Guru tidak ditemukan' });
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return serializeForClient({ ok: true, message: 'Seluruh data guru berhasil dihapus dari Spreadsheet' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Hapus Single Guru
function deleteGuruFromSheet(guruId) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Guru');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Guru tidak ditemukan' });
    
    var data = sheet.getDataRange().getValues();
    var targetId = String(guruId).trim();
    
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === targetId) {
        sheet.deleteRow(i + 1);
        return serializeForClient({ ok: true, message: 'Guru berhasil dihapus dari Spreadsheet' });
      }
    }
    return serializeForClient({ ok: false, message: 'Guru tidak ditemukan di Spreadsheet' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Edit Guru
function updateGuruInSheet(guru) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Guru');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Guru tidak ditemukan' });
    
    var data = sheet.getDataRange().getValues();
    var targetId = String(guru.id).trim();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        sheet.getRange(i + 1, 2).setValue(guru.nama);
        sheet.getRange(i + 1, 3).setValue(guru.username);
        if (guru.password) {
          sheet.getRange(i + 1, 4).setValue(guru.password);
        }
        sheet.getRange(i + 1, 5).setValue(guru.role);
        sheet.getRange(i + 1, 6).setValue(guru.foto || '');
        return serializeForClient({ ok: true, message: 'Data guru berhasil diperbarui di Spreadsheet' });
      }
    }
    return serializeForClient({ ok: false, message: 'Guru tidak ditemukan di Spreadsheet' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Simpan Setoran Tahfidz
function saveTahfidzToSheet(data) {
  try {
    var ss = getDbSpreadsheet();

    // Validasi Keamanan Backend: Pastikan siswa adalah binaan guru yang bersangkutan
    if (data.guru && data.guru !== 'Administrator Sekolah' && !isStudentAssignedToTeacher(data.nis, data.guru)) {
      return serializeForClient({ ok: false, message: 'Akses Ditolak: Siswa ' + (data.nama || data.nis) + ' bukan merupakan siswa binaan ' + data.guru });
    }

    var sheet = ss.getSheetByName('Tahfidz');
    sheet.appendRow([data.tanggal, data.nis, data.nama, data.surat, data.ayat, data.nilai, data.catatan, data.guru]);
    
    // Sinkron ke Rekap dengan Program Tahfidz
    var programName = data.program || 'Tahfidz';
    var rekapSheet = ss.getSheetByName('Rekap');
    rekapSheet.appendRow([data.tanggal, data.nis, data.nama, programName, 'Surat ' + data.surat + ' Ayat ' + data.ayat, 'Nilai: ' + data.nilai, data.catatan || '-', data.guru]);
    return serializeForClient({ ok: true, message: 'Setoran Tahfidz berhasil disimpan' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Otentikasi Pengguna (Guru, Siswa, & Admin)
function authenticateUser(username, password, role) {
  try {
    setupDatabase();
    var u = (username || '').trim().toLowerCase();
    var p = (password || '').trim();
    var r = (role || 'guru').trim().toLowerCase();

    if (r === 'siswa') {
      if (p !== '@tahsinsmp') {
        return serializeForClient({ ok: false, message: 'Password Siswa salah! Gunakan password: @tahsinsmp' });
      }
      var ss = getDbSpreadsheet();
      var sheet = ss.getSheetByName('Siswa');
      if (sheet) {
        var data = sheet.getDataRange().getValues();

        function cleanNameStr(str) {
          return String(str || '').toLowerCase()
            .replace(/s\.pd\.i|s\.pd|s\.ag|m\.pd|ustadzah|ustadz|m\.|h\.|drs\./gi, '')
            .replace(/[^a-z0-9]/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
        }

        var cleanInput = cleanNameStr(u);
        var rawInput = u.toLowerCase();

        // 1. Match Presisi: NIS Persis atau Nama Lengkap Persis
        for (var i = 1; i < data.length; i++) {
          var nisVal = String(data[i][0] || '').trim().toLowerCase();
          var namaVal = String(data[i][1] || '').trim().toLowerCase();
          if (rawInput === nisVal || rawInput === namaVal) {
            return serializeForClient({
              ok: true,
              message: 'Login Siswa Berhasil',
              user: { nis: String(data[i][0]).trim(), name: String(data[i][1]).trim(), kelas: String(data[i][2]).trim(), kelompok: String(data[i][3]).trim(), guru: String(data[i][4]).trim(), status: String(data[i][5] || 'Aktif').trim(), foto: String(data[i][6] || '').trim(), role: 'Siswa' }
            });
          }
        }

        // 2. Match Presisi Nama Bersih
        for (var k = 1; k < data.length; k++) {
          var cleanSNama = cleanNameStr(data[k][1]);
          if (cleanInput === cleanSNama) {
            return serializeForClient({
              ok: true,
              message: 'Login Siswa Berhasil',
              user: { nis: String(data[k][0]).trim(), name: String(data[k][1]).trim(), kelas: String(data[k][2]).trim(), kelompok: String(data[k][3]).trim(), guru: String(data[k][4]).trim(), status: String(data[k][5] || 'Aktif').trim(), foto: String(data[k][6] || '').trim(), role: 'Siswa' }
            });
          }
        }

        // 3. Match Token Cerdas: Kumpulkan semua calon yang cocok
        var inputTokens = cleanInput.split(' ').filter(function(t) { return t.length >= 2; });
        if (inputTokens.length > 0) {
          var candidates = [];
          for (var j = 1; j < data.length; j++) {
            var sTokens = cleanNameStr(data[j][1]).split(' ');
            var allMatch = inputTokens.every(function(it) {
              return sTokens.some(function(st) { return st === it; });
            });
            if (allMatch) {
              candidates.push(data[j]);
            }
          }

          if (candidates.length === 1) {
            var mData = candidates[0];
            return serializeForClient({
              ok: true,
              message: 'Login Siswa Berhasil',
              user: { nis: String(mData[0]).trim(), name: String(mData[1]).trim(), kelas: String(mData[2]).trim(), kelompok: String(mData[3]).trim(), guru: String(mData[4]).trim(), status: String(mData[5] || 'Aktif').trim(), foto: String(mData[6] || '').trim(), role: 'Siswa' }
            });
          } else if (candidates.length > 1) {
            return serializeForClient({
              ok: false,
              message: 'Ditemukan ' + candidates.length + ' siswa dengan nama serupa. Silakan ketik Nama Lengkap atau NIS Anda secara lengkap!'
            });
          }
        }
      }

      return serializeForClient({
        ok: false,
        message: 'Nama atau NIS Siswa "' + username + '" tidak ditemukan. Silakan periksa kembali ejaan nama atau NIS Anda!'
      });
    }

    // Check Guru tab in Spreadsheet first (supports ID Guru, Username, or Name)
    var ssGuru = getDbSpreadsheet();
    var guruSheet = ssGuru.getSheetByName('Guru');
    if (guruSheet) {
      var gData = guruSheet.getDataRange().getValues();
      for (var j = 1; j < gData.length; j++) {
        var gId = String(gData[j][0] || '').trim().toLowerCase();
        var gNama = String(gData[j][1] || '').trim();
        var gUser = String(gData[j][2] || '').trim().toLowerCase();
        var gPass = String(gData[j][3] || '').trim();
        var gRole = String(gData[j][4] || 'Guru').trim();

        var hashedInputPass = hashPassword(p);
        var storedPass = String(gPass).trim();
        var isPassValid = (p === storedPass || hashedInputPass === storedPass);

        if ((u === gUser || u === gId || (u.length >= 3 && gNama.toLowerCase().indexOf(u) !== -1)) && isPassValid) {
          return serializeForClient({
            ok: true,
            message: 'Login ' + gRole + ' Berhasil',
            user: { id: gData[j][0], name: gNama, username: gData[j][2], role: gRole, foto: gData[j][5] || '' }
          });
        }
      }
    }

    // Admin default login fallback
    if ((u === 'smp penawa' || u === 'admin') && (p === '@penawa123' || p === 'admin')) {
      return serializeForClient({ ok: true, message: 'Login Admin Berhasil', user: { id: 'ADM01', name: 'Administrator Sekolah', role: 'Admin' } });
    }

    // Guru demo fallback
    if (u === 'mansur' && (p === 'mansur123' || p === 'mansur')) {
      return serializeForClient({ ok: true, message: 'Login Berhasil', user: { id: 'G01', name: 'Ustadz Mansur, S.Pd.I', role: 'Guru' } });
    }

    return serializeForClient({ ok: false, message: 'Username / ID atau Password salah!' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Ambil Seluruh Data Awal dari Sheet
function getInitialDataFromSheet() {
  try {
    setupDatabase();
    var ss = getDbSpreadsheet();
    
    // Config
    var config = { schoolName: 'SMP Pesantren Nahdlatul Wathon', appTitle: 'PENA-Q', logoUrl: '', academicYear: '2026/2027' };
    var cfgSheet = ss.getSheetByName('Pengaturan');
    if (cfgSheet) {
      var cfgData = cfgSheet.getDataRange().getValues();
      for (var i = 1; i < cfgData.length; i++) {
        var k = String(cfgData[i][0]).trim();
        var v = String(cfgData[i][1]).trim();
        if (k) config[k] = v;
      }
    }

    // Guru
    var guruList = [];
    var gSheet = ss.getSheetByName('Guru');
    if (gSheet) {
      var gData = gSheet.getDataRange().getValues();
      for (var j = 1; j < gData.length; j++) {
        if (gData[j][0]) {
          guruList.push({
            id: String(gData[j][0]),
            name: String(gData[j][1]),
            username: String(gData[j][2]),
            role: String(gData[j][4] || 'Guru'),
            foto: String(gData[j][5] || '')
          });
        }
      }
    }

    // Siswa
    var siswaList = [];
    var sSheet = ss.getSheetByName('Siswa');
    if (sSheet) {
      var sData = sSheet.getDataRange().getValues();
      for (var k = 1; k < sData.length; k++) {
        if (sData[k][0]) {
          siswaList.push({
            nis: String(sData[k][0]),
            name: String(sData[k][1]),
            kelas: String(sData[k][2]),
            kelompok: String(sData[k][3]),
            guru: String(sData[k][4]),
            status: String(sData[k][5] || 'Aktif'),
            foto: String(sData[k][6] || ''),
            level: String(sData[k][3]).toLowerCase().indexOf('tahfidz') !== -1 ? 'Tahfidz' : (String(sData[k][3]).toLowerCase().indexOf('qur') !== -1 ? 'Tahsin Al-Qur\'an' : 'Tahsin Iqro')
          });
        }
      }
    }

    // Rekap / Setoran
    var setoranLogs = [];
    var rSheet = ss.getSheetByName('Rekap');
    if (rSheet) {
      var rData = rSheet.getDataRange().getValues();
      for (var m = rData.length - 1; m >= 1; m--) {
        if (rData[m][0]) {
          var tglVal = rData[m][0];
          var tglStr = (tglVal instanceof Date) ? Utilities.formatDate(tglVal, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(tglVal).split('T')[0];
          var catVal = String(rData[m][6] !== undefined && rData[m][6] !== '' ? rData[m][6] : '-');
          var guruVal = String(rData[m][7] !== undefined ? rData[m][7] : '');
          setoranLogs.push({
            tanggal: tglStr,
            nis: String(rData[m][1]),
            nama: String(rData[m][2]),
            type: String(rData[m][3]),
            detail: String(rData[m][4]),
            status: String(rData[m][5]),
            catatan: catVal,
            guru: guruVal
          });
        }
      }
    }

    return serializeForClient({
      ok: true,
      data: {
        config: config,
        guruList: guruList,
        siswaList: siswaList,
        setoranLogs: setoranLogs
      }
    });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Simpan Pengaturan Sekolah
function saveAppConfigToSheet(config) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Pengaturan');
    if (!sheet) {
      sheet = ss.insertSheet('Pengaturan');
    }
    sheet.clearContents();
    sheet.getRange(1, 1, 1, 2).setValues([['Kunci', 'Nilai']]);
    sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#10b981').setFontColor('#ffffff');
    sheet.appendRow(['schoolName', config.schoolName || '']);
    sheet.appendRow(['appTitle', config.appTitle || '']);
    sheet.appendRow(['logoUrl', config.logoUrl || '']);
    sheet.appendRow(['academicYear', config.academicYear || '2026/2027']);
    return serializeForClient({ ok: true, message: 'Pengaturan sekolah berhasil disimpan' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Ambil Pengaturan Sekolah
function getAppConfigFromSheet() {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Pengaturan');
    var config = {
      schoolName: 'SMP Pesantren Nahdlatul Wathon',
      appTitle: 'PENA-Q',
      logoUrl: '',
      academicYear: '2026/2027'
    };
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var key = String(data[i][0]).trim();
        var val = String(data[i][1]).trim();
        if (key) {
          config[key] = val;
        }
      }
    }
    return serializeForClient({ ok: true, data: config });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Edit Data Siswa
function updateStudentInSheet(student) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Siswa');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Siswa tidak ditemukan' });
    
    var data = sheet.getDataRange().getValues();
    var targetNis = String(student.oldNis || student.nis).trim();
    
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetNis) {
        sheet.getRange(i + 1, 1).setValue(student.nis);
        sheet.getRange(i + 1, 2).setValue(student.nama);
        sheet.getRange(i + 1, 3).setValue(student.kelas);
        sheet.getRange(i + 1, 4).setValue(student.kelompok);
        sheet.getRange(i + 1, 5).setValue(student.guru);
        sheet.getRange(i + 1, 6).setValue(student.status || 'Aktif');
        sheet.getRange(i + 1, 7).setValue(student.foto || '');

        // Juga perbarui NIS & Nama di Sheet Rekap jika NIS/Nama berubah agar riwayat tidak hilang
        var rekapSheet = ss.getSheetByName('Rekap');
        if (rekapSheet && rekapSheet.getLastRow() > 1) {
          var rData = rekapSheet.getDataRange().getValues();
          for (var rIdx = 1; rIdx < rData.length; rIdx++) {
            if (String(rData[rIdx][1]).trim() === targetNis) {
              rekapSheet.getRange(rIdx + 1, 2).setValue(student.nis);
              rekapSheet.getRange(rIdx + 1, 3).setValue(student.nama);
            }
          }
        }

        return serializeForClient({ ok: true, message: 'Data siswa berhasil diperbarui' });
      }
    }
    return serializeForClient({ ok: false, message: 'Siswa tidak ditemukan di Spreadsheet' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Tambah Siswa Baru Manual
function addStudentToSheet(student) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Siswa');
    sheet.appendRow([student.nis, student.nama, student.kelas, student.kelompok, student.guru, student.status, student.foto || '']);
    return serializeForClient({ ok: true, message: 'Siswa berhasil ditambahkan' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Import Banyak Siswa dari CSV
function importStudentsToSheet(studentsArray) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Siswa');
    studentsArray.forEach(function(student) {
      sheet.appendRow([student.nis, student.nama, student.kelas, student.kelompok, student.guru, student.status, student.foto || '']);
    });
    return serializeForClient({ ok: true, message: 'Berhasil mengimpor ' + studentsArray.length + ' siswa' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Tambah Guru Baru Manual
function addGuruToSheet(guru) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Guru');
    sheet.appendRow([guru.id, guru.nama, guru.username, guru.password, guru.role, guru.foto || '']);
    return serializeForClient({ ok: true, message: 'Guru berhasil ditambahkan' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Import Banyak Guru dari CSV
function importGuruToSheet(guruArray) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Guru');
    guruArray.forEach(function(guru) {
      sheet.appendRow([guru.id, guru.nama, guru.username, guru.password, guru.role, guru.foto || '']);
    });
    return serializeForClient({ ok: true, message: 'Berhasil mengimpor ' + guruArray.length + ' guru' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Hapus Siswa Manual Berdasarkan NIS
function deleteStudentFromSheet(nis) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Siswa');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Siswa tidak ditemukan' });
    
    var data = sheet.getDataRange().getValues();
    var targetNis = String(nis).trim();
    
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === targetNis) {
        sheet.deleteRow(i + 1);
        return serializeForClient({ ok: true, message: 'Siswa berhasil dihapus dari Spreadsheet' });
      }
    }
    return serializeForClient({ ok: true, message: 'Siswa tidak ditemukan di Spreadsheet, diperbarui secara lokal' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Hapus Siswa Per Kelas
function deleteClassFromSheet(kelas) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Siswa');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Siswa tidak ditemukan' });
    
    var data = sheet.getDataRange().getValues();
    var targetKelas = String(kelas).trim().toLowerCase();
    var deletedCount = 0;
    
    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][2]).trim().toLowerCase() === targetKelas) {
        sheet.deleteRow(i + 1);
        deletedCount++;
      }
    }
    return serializeForClient({ ok: true, message: 'Berhasil menghapus ' + deletedCount + ' siswa kelas ' + kelas });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Hapus Seluruh Siswa
function deleteAllStudentsFromSheet() {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Siswa');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Siswa tidak ditemukan' });
    
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    return serializeForClient({ ok: true, message: 'Seluruh data siswa berhasil dihapus dari Spreadsheet' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Kirim Laporan Guru ke Admin
function sendReportToAdmin(guruName, note) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Pengaturan');
    if (sheet) {
      var tgl = new Date().toISOString().split('T')[0];
      sheet.appendRow(['laporan_guru_' + guruName, 'Dikirim pada ' + tgl + ': ' + (note || 'Laporan telah dikirim')]);
    }
    return serializeForClient({ ok: true, message: 'Laporan berhasil dikirim ke Admin!' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Edit / Input Ulang Setoran di Spreadsheet
function updateSetoranInSheet(data) {
  try {
    var ss = getDbSpreadsheet();
    var rekapSheet = ss.getSheetByName('Rekap');
    if (!rekapSheet) return serializeForClient({ ok: false, message: 'Sheet Rekap tidak ditemukan' });

    var rData = rekapSheet.getDataRange().getValues();
    var targetNis = String(data.nis).trim();
    var targetTanggal = String(data.tanggal).trim();

    var updated = false;
    for (var i = rData.length - 1; i >= 1; i--) {
      var rawDate = rData[i][0];
      var dateStr = (rawDate instanceof Date) ? rawDate.toISOString().split('T')[0] : String(rawDate).split('T')[0];
      if (String(rData[i][1]).trim() === targetNis && dateStr === targetTanggal) {
        rekapSheet.getRange(i + 1, 4).setValue(data.type);
        rekapSheet.getRange(i + 1, 5).setValue(data.detail);
        rekapSheet.getRange(i + 1, 6).setValue(data.status);
        rekapSheet.getRange(i + 1, 7).setValue(data.catatan || '-');
        rekapSheet.getRange(i + 1, 8).setValue(data.guru);
        updated = true;
        break;
      }
    }

    if (!updated) {
      rekapSheet.appendRow([data.tanggal, data.nis, data.nama, data.type, data.detail, data.status, data.catatan || '-', data.guru]);
    }

    return serializeForClient({ ok: true, message: 'Data setoran berhasil diperbarui di Spreadsheet' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Fungsi Hapus Setoran / Laporan Per Periode (Harian, Mingguan, Bulanan)
function deleteSetoranLogsByPeriod(periodType, startDate, endDate) {
  try {
    var ss = getDbSpreadsheet();
    var sheetNames = ['Rekap', 'Tahsin', 'Tahfidz'];
    var totalDeleted = 0;

    sheetNames.forEach(function(sName) {
      var sheet = ss.getSheetByName(sName);
      if (!sheet) return;
      var data = sheet.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        var rawDate = data[i][0];
        if (!rawDate) continue;
        
        var dateStr = '';
        if (rawDate instanceof Date) {
          dateStr = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          dateStr = String(rawDate).trim().split('T')[0];
        }

        var shouldDelete = false;
        if (periodType === 'harian') {
          if (dateStr === startDate) shouldDelete = true;
        } else if (startDate && endDate) {
          if (dateStr >= startDate && dateStr <= endDate) shouldDelete = true;
        }

        if (shouldDelete) {
          sheet.deleteRow(i + 1);
          if (sName === 'Rekap') {
            totalDeleted++;
          }
        }
      }
    });

    return serializeForClient({ ok: true, message: 'Berhasil menghapus ' + totalDeleted + ' data setoran.', deletedCount: totalDeleted });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

// Backend Handlers untuk Pembina Dashboard (High Performance Query, Filter, Pagination & CRUD)
function getDashboardData(params) {
  try {
    setupDatabase();
    params = params || {};
    var pembinaId = String(params.pembinaId || 'all').trim();
    var search = String(params.search || '').trim().toLowerCase();
    var classGrade = String(params.classGrade || 'all').trim();
    var groupName = String(params.groupName || 'all').trim();
    var status = String(params.status || 'all').trim();
    var page = parseInt(params.page || 1, 10);
    var pageSize = parseInt(params.pageSize || 25, 10);

    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) {
      return serializeForClient({ ok: false, message: 'Sheet Students tidak ditemukan', data: null });
    }

    var rawData = sheet.getDataRange().getValues();
    var allStudents = [];

    for (var i = 1; i < rawData.length; i++) {
      var row = rawData[i];
      if (!row[0]) continue;

      var sId = String(row[0]).trim();
      var name = String(row[1]).trim();
      var nisn = String(row[2]).trim();
      var sClass = String(row[3]).trim();
      var group = String(row[4]).trim();
      var pId = String(row[5]).trim();
      var sStatus = String(row[6] || 'Active').trim();
      var attPct = parseFloat(row[7]) || 0;
      var perfPts = parseFloat(row[8]) || 0;
      var contact = String(row[9] || '').trim();
      var parentContact = String(row[10] || '').trim();
      var lastUpdated = row[11];
      if (lastUpdated instanceof Date) {
        lastUpdated = Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
      } else {
        lastUpdated = String(lastUpdated || '').trim();
      }

      allStudents.push({
        id: sId,
        name: name,
        nisn: nisn,
        classGrade: sClass,
        groupName: group,
        pembinaId: pId,
        status: sStatus,
        attendancePct: attPct,
        performancePts: perfPts,
        contact: contact,
        parentContact: parentContact,
        lastUpdated: lastUpdated
      });
    }

    var assignedStudents = allStudents.filter(function(st) {
      if (pembinaId !== 'all' && st.pembinaId.toLowerCase() !== pembinaId.toLowerCase()) {
        return false;
      }
      return true;
    });

    var totalStudents = assignedStudents.length;
    var activeCount = 0;
    var inactiveCount = 0;
    var onLeaveCount = 0;
    var totalAtt = 0;
    var lowPerformanceCount = 0;

    assignedStudents.forEach(function(st) {
      if (st.status === 'Active') activeCount++;
      else if (st.status === 'Inactive') inactiveCount++;
      else if (st.status === 'On Leave') onLeaveCount++;

      totalAtt += st.attendancePct;
      if (st.performancePts < 65 || st.attendancePct < 75) {
        lowPerformanceCount++;
      }
    });

    var avgAtt = totalStudents > 0 ? Math.round((totalAtt / totalStudents) * 10) / 10 : 0;

    var filtered = assignedStudents.filter(function(st) {
      if (classGrade !== 'all' && st.classGrade.toLowerCase() !== classGrade.toLowerCase()) return false;
      if (groupName !== 'all' && st.groupName.toLowerCase() !== groupName.toLowerCase()) return false;
      if (status !== 'all' && st.status.toLowerCase() !== status.toLowerCase()) return false;

      if (search) {
        var matchId = st.id.toLowerCase().indexOf(search) !== -1;
        var matchName = st.name.toLowerCase().indexOf(search) !== -1;
        var matchNisn = st.nisn.toLowerCase().indexOf(search) !== -1;
        var matchClass = st.classGrade.toLowerCase().indexOf(search) !== -1;
        if (!matchId && !matchName && !matchNisn && !matchClass) return false;
      }

      return true;
    });

    var totalFiltered = filtered.length;
    var totalPages = Math.ceil(totalFiltered / pageSize) || 1;
    if (page > totalPages) page = totalPages;
    if (page < 1) page = 1;

    var startIndex = (page - 1) * pageSize;
    var paginatedStudents = filtered.slice(startIndex, startIndex + pageSize);

    return serializeForClient({
      ok: true,
      data: {
        students: paginatedStudents,
        stats: {
          totalStudents: totalStudents,
          activeStudents: activeCount,
          inactiveStudents: inactiveCount,
          onLeaveStudents: onLeaveCount,
          avgAttendance: avgAtt,
          lowPerformanceCount: lowPerformanceCount
        },
        pagination: {
          total: totalFiltered,
          page: page,
          pageSize: pageSize,
          totalPages: totalPages
        }
      }
    });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString(), data: null });
  }
}

function addStudent(studentData) {
  try {
    setupDatabase();
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Students tidak ditemukan' });

    var sId = studentData.id || ('STD' + (sheet.getLastRow() + 1000));
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    sheet.appendRow([
      sId,
      studentData.name || '',
      studentData.nisn || '',
      studentData.classGrade || '7A',
      studentData.groupName || 'Tahsin Iqro',
      studentData.pembinaId || 'G01',
      studentData.status || 'Active',
      parseFloat(studentData.attendancePct) || 100,
      parseFloat(studentData.performancePts) || 85,
      studentData.contact || '',
      studentData.parentContact || '',
      nowStr
    ]);

    logActivity(studentData.pembinaId || 'ADM01', 'Add Student', sId);

    return serializeForClient({ ok: true, message: 'Data siswa berhasil ditambahkan', studentId: sId });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

function updateStudent(studentData) {
  try {
    setupDatabase();
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Students tidak ditemukan' });

    var data = sheet.getDataRange().getValues();
    var targetId = String(studentData.id).trim();
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        sheet.getRange(i + 1, 2).setValue(studentData.name);
        sheet.getRange(i + 1, 3).setValue(studentData.nisn);
        sheet.getRange(i + 1, 4).setValue(studentData.classGrade);
        sheet.getRange(i + 1, 5).setValue(studentData.groupName);
        sheet.getRange(i + 1, 6).setValue(studentData.pembinaId);
        sheet.getRange(i + 1, 7).setValue(studentData.status);
        sheet.getRange(i + 1, 8).setValue(parseFloat(studentData.attendancePct) || 0);
        sheet.getRange(i + 1, 9).setValue(parseFloat(studentData.performancePts) || 0);
        sheet.getRange(i + 1, 10).setValue(studentData.contact || '');
        sheet.getRange(i + 1, 11).setValue(studentData.parentContact || '');
        sheet.getRange(i + 1, 12).setValue(nowStr);

        logActivity(studentData.pembinaId || 'ADM01', 'Update Student', targetId);
        return serializeForClient({ ok: true, message: 'Data siswa berhasil diperbarui' });
      }
    }
    return serializeForClient({ ok: false, message: 'ID Siswa tidak ditemukan' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

function deleteStudent(studentId) {
  try {
    setupDatabase();
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Students tidak ditemukan' });

    var data = sheet.getDataRange().getValues();
    var targetId = String(studentId).trim();

    for (var i = data.length - 1; i >= 1; i--) {
      if (String(data[i][0]).trim() === targetId) {
        sheet.deleteRow(i + 1);
        logActivity('ADM01', 'Delete Student', targetId);
        return serializeForClient({ ok: true, message: 'Siswa berhasil dihapus' });
      }
    }
    return serializeForClient({ ok: false, message: 'ID Siswa tidak ditemukan' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

function batchUpdateStatus(studentIds, newStatus, pembinaId) {
  try {
    setupDatabase();
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Students tidak ditemukan' });

    if (!Array.isArray(studentIds) || studentIds.length === 0) {
      return serializeForClient({ ok: false, message: 'Daftar ID siswa tidak boleh kosong' });
    }

    var data = sheet.getDataRange().getValues();
    var idSet = {};
    studentIds.forEach(function(id) { idSet[String(id).trim()] = true; });
    var updatedCount = 0;
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    for (var i = 1; i < data.length; i++) {
      var rowId = String(data[i][0]).trim();
      if (idSet[rowId]) {
        sheet.getRange(i + 1, 7).setValue(newStatus);
        sheet.getRange(i + 1, 12).setValue(nowStr);
        updatedCount++;
      }
    }

    logActivity(pembinaId || 'ADM01', 'Batch Status Update to ' + newStatus, studentIds.join(', '));
    return serializeForClient({ ok: true, message: 'Status berhasil diperbarui untuk ' + updatedCount + ' siswa' });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString() });
  }
}

function getStudentDetails(studentId) {
  try {
    setupDatabase();
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Students');
    if (!sheet) return serializeForClient({ ok: false, message: 'Sheet Students tidak ditemukan', data: null });

    var data = sheet.getDataRange().getValues();
    var targetId = String(studentId).trim();
    var studentObj = null;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === targetId) {
        var row = data[i];
        var lastUpdated = row[11];
        if (lastUpdated instanceof Date) {
          lastUpdated = Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
        } else {
          lastUpdated = String(lastUpdated || '').trim();
        }

        studentObj = {
          id: String(row[0]).trim(),
          name: String(row[1]).trim(),
          nisn: String(row[2]).trim(),
          classGrade: String(row[3]).trim(),
          groupName: String(row[4]).trim(),
          pembinaId: String(row[5]).trim(),
          status: String(row[6] || 'Active').trim(),
          attendancePct: parseFloat(row[7]) || 0,
          performancePts: parseFloat(row[8]) || 0,
          contact: String(row[9] || '').trim(),
          parentContact: String(row[10] || '').trim(),
          lastUpdated: lastUpdated
        };
        break;
      }
    }

    if (!studentObj) {
      return serializeForClient({ ok: false, message: 'Siswa tidak ditemukan', data: null });
    }

    var logSheet = ss.getSheetByName('Activity_Logs');
    var studentLogs = [];
    if (logSheet) {
      var logData = logSheet.getDataRange().getValues();
      for (var k = logData.length - 1; k >= 1; k--) {
        var tStudent = String(logData[k][4] || '').trim();
        if (tStudent.indexOf(targetId) !== -1) {
          var logTime = logData[k][1];
          if (logTime instanceof Date) {
            logTime = Utilities.formatDate(logTime, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
          } else {
            logTime = String(logTime || '').trim();
          }
          studentLogs.push({
            logId: String(logData[k][0]),
            timestamp: logTime,
            pembinaId: String(logData[k][2]),
            action: String(logData[k][3])
          });
        }
      }
    }

    return serializeForClient({
      ok: true,
      data: {
        student: studentObj,
        logs: studentLogs
      }
    });
  } catch (err) {
    return serializeForClient({ ok: false, message: err.toString(), data: null });
  }
}

function logActivity(pembinaId, action, targetStudentId) {
  try {
    var ss = getDbSpreadsheet();
    var sheet = ss.getSheetByName('Activity_Logs');
    if (!sheet) return;
    var logId = 'LOG' + new Date().getTime();
    var nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([logId, nowStr, pembinaId || 'SYSTEM', action || '', targetStudentId || '']);
  } catch (err) {
    // Ignore log write errors
  }
}

// ==================== MENU KUSTOM & IMPORT DATA ====================

// ID Spreadsheet Lama (sumber data lengkap)
var OLD_SPREADSHEET_ID = '1yFkDcD5k-G9gDzrvD9kMBKnuJsWaADNzeg_DFPpdRaM';

/**
 * Menu Kustom yang muncul di Google Sheets (sebelah menu Ekstensi)
 * Otomatis tampil setiap kali Spreadsheet dibuka
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ PENA-Q Tools')
    .addItem('🔄 Import Data dari SS Lama', 'importDataFromOldSpreadsheet')
    .addSeparator()
    .addItem('🗄️ Setup Database (Buat Sheet & Header)', 'setupDatabase')
    .addItem('📊 Cek Jumlah Data Semua Sheet', 'cekJumlahDataSemuaSheet')
    .addToUi();
}

/**
 * Import (Clone) seluruh data dari Spreadsheet Lama ke Spreadsheet Aktif
 * Sheet yang di-import: Guru, Siswa, Tahsin, Tahfidz, Rekap, Pengaturan, Students, Pembina, Activity_Logs
 */
function importDataFromOldSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  
  // Konfirmasi sebelum import
  var response = ui.alert(
    '⚠️ Konfirmasi Import Data',
    'Fungsi ini akan MENIMPA seluruh data di Spreadsheet aktif dengan data dari Spreadsheet Lama.\n\n' +
    'Spreadsheet Sumber: ' + OLD_SPREADSHEET_ID + '\n\n' +
    'Sheet yang akan di-clone:\n' +
    '• Guru\n• Siswa\n• Tahsin\n• Tahfidz\n• Rekap\n• Pengaturan\n• Students\n• Pembina\n• Activity_Logs\n\n' +
    'Apakah Anda yakin ingin melanjutkan?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Import dibatalkan.');
    return;
  }
  
  try {
    var oldSs = SpreadsheetApp.openById(OLD_SPREADSHEET_ID);
    var newSs = getDbSpreadsheet();
    
    // Pastikan struktur database sudah siap di SS baru
    setupDatabase();
    
    // Daftar semua sheet yang akan di-clone
    var sheetNames = ['Guru', 'Siswa', 'Tahsin', 'Tahfidz', 'Rekap', 'Pengaturan', 'Students', 'Pembina', 'Activity_Logs'];
    
    var totalImported = 0;
    var importLog = [];
    
    for (var i = 0; i < sheetNames.length; i++) {
      var sheetName = sheetNames[i];
      var sourceSheet = oldSs.getSheetByName(sheetName);
      var targetSheet = newSs.getSheetByName(sheetName);
      
      if (!sourceSheet) {
        importLog.push('⏭️ ' + sheetName + ': Tidak ditemukan di SS Lama (skip)');
        continue;
      }
      
      if (!targetSheet) {
        targetSheet = newSs.insertSheet(sheetName);
        importLog.push('🆕 ' + sheetName + ': Sheet baru dibuat');
      }
      
      var sourceData = sourceSheet.getDataRange().getValues();
      var sourceRows = sourceData.length;
      var sourceCols = sourceData[0] ? sourceData[0].length : 0;
      
      if (sourceRows <= 0 || sourceCols <= 0) {
        importLog.push('⚠️ ' + sheetName + ': Kosong di SS Lama (skip)');
        continue;
      }
      
      // Bersihkan sheet target terlebih dahulu
      targetSheet.clearContents();
      targetSheet.clearFormats();
      
      // Tulis seluruh data (header + isi) dari sumber ke target
      targetSheet.getRange(1, 1, sourceRows, sourceCols).setValues(sourceData);
      
      // Format header row (baris 1) supaya tetap rapi
      targetSheet.getRange(1, 1, 1, sourceCols)
        .setFontWeight('bold')
        .setBackground('#10b981')
        .setFontColor('#ffffff');
      
      var dataRows = sourceRows - 1; // Minus header
      totalImported += dataRows;
      importLog.push('✅ ' + sheetName + ': ' + dataRows + ' baris data berhasil di-import');
    }
    
    // ===== PASTIKAN AKUN ADMIN SELALU ADA DI SHEET GURU =====
    var guruSheet = newSs.getSheetByName('Guru');
    if (guruSheet) {
      var guruData = guruSheet.getDataRange().getValues();
      var adminExists = false;
      for (var a = 1; a < guruData.length; a++) {
        var idGuru = String(guruData[a][0] || '').trim().toUpperCase();
        var roleGuru = String(guruData[a][4] || '').trim().toLowerCase();
        if (idGuru === 'ADM01' || roleGuru === 'admin') {
          adminExists = true;
          break;
        }
      }
      if (!adminExists) {
        // Sisipkan akun Admin di baris 2 (setelah header)
        guruSheet.insertRowAfter(1);
        guruSheet.getRange(2, 1, 1, 6).setValues([['ADM01', 'Administrator Sekolah', 'admin', '@penawa123', 'Admin', '']]);
        // Reset format agar tidak ikut hijau seperti header
        guruSheet.getRange(2, 1, 1, 6)
          .setBackground('#ffffff')
          .setFontColor('#000000')
          .setFontWeight('normal');
        totalImported += 1;
        importLog.push('🔐 Guru: Akun Admin (ADM01) otomatis ditambahkan karena tidak ada di SS Lama');
      } else {
        importLog.push('🔐 Guru: Akun Admin sudah ada ✓');
      }
    }
    
    // Tampilkan laporan hasil import
    var report = '🎉 IMPORT DATA SELESAI!\n\n' +
                 'Total baris data yang berhasil di-import: ' + totalImported + '\n\n' +
                 '📋 Detail per Sheet:\n' +
                 importLog.join('\n') + '\n\n' +
                 'Sumber: Spreadsheet Lama (' + OLD_SPREADSHEET_ID.substring(0, 20) + '...)\n' +
                 'Tujuan: Spreadsheet Aktif (' + newSs.getName() + ')';
    
    ui.alert('✅ Import Berhasil!', report, ui.ButtonSet.OK);
    
  } catch (err) {
    ui.alert(
      '❌ Gagal Import',
      'Terjadi error saat import data:\n\n' + err.toString() + '\n\n' +
      'Pastikan:\n' +
      '1. ID Spreadsheet Lama benar\n' +
      '2. Anda memiliki akses baca ke SS Lama\n' +
      '3. Koneksi internet stabil',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Cek jumlah data di semua sheet (untuk verifikasi setelah import)
 */
function cekJumlahDataSemuaSheet() {
  var ui = SpreadsheetApp.getUi();
  var ss = getDbSpreadsheet();
  var sheetNames = ['Guru', 'Siswa', 'Tahsin', 'Tahfidz', 'Rekap', 'Pengaturan', 'Students', 'Pembina', 'Activity_Logs'];
  
  var report = [];
  var totalRows = 0;
  
  for (var i = 0; i < sheetNames.length; i++) {
    var sheet = ss.getSheetByName(sheetNames[i]);
    if (sheet) {
      var lastRow = sheet.getLastRow();
      var dataRows = Math.max(0, lastRow - 1);
      totalRows += dataRows;
      report.push('📄 ' + sheetNames[i] + ': ' + dataRows + ' baris data (+ 1 header)');
    } else {
      report.push('❌ ' + sheetNames[i] + ': Sheet tidak ditemukan');
    }
  }
  
  ui.alert(
    '📊 Laporan Jumlah Data',
    'Spreadsheet: ' + ss.getName() + '\n\n' +
    report.join('\n') + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    '📌 TOTAL: ' + totalRows + ' baris data',
    ui.ButtonSet.OK
  );
}