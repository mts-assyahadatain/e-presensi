/**
 * ==============================================================================
 * E-ABSENSI SISWA BACKEND SCRIPT (Code.gs)
 * ID Spreadsheet: 1VLfOIWnBR_s_RLtdJN6OGrz8Pw-sKNjEGe8WnrJbxM0
 * ==============================================================================
 */

const SPREADSHEET_ID = '1O7QI418ZW27zAsJOW85xbqDx7pn68BX1IJWaBh-HQzM';

/**
 * Mendapatkan instance Spreadsheet berdasarkan ID
 */
function getSpreadsheet() {
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (error) {
    throw new Error('Gagal membuka Spreadsheet. Pastikan ID Spreadsheet benar dan telah diberi izin akses: ' + error.message);
  }
}

/**
 * Handler utama HTTP GET untuk menampilkan halaman HTML E-Absensi
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('E-Absensi Siswa Modern')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Mengambil seluruh data awal (Siswa, Pengaturan, dan Statistik Kehadiran Hari Ini)
 */
function getAppData() {
  const ss = getSpreadsheet();
  
  // 1. Inisialisasi Sheet Data Siswa
  let sheetSiswa = ss.getSheetByName('Siswa');
  if (!sheetSiswa) {
    sheetSiswa = ss.insertSheet('Siswa');
    sheetSiswa.appendRow(['NIS', 'Nama', 'Kelas', 'Jenis Kelamin']);
    // Data bawaan contoh
    sheetSiswa.appendRow(['1001', 'Ahmad Rizki', 'XII RPL 1', 'Laki-laki']);
    sheetSiswa.appendRow(['1002', 'Siti Aminah', 'XII RPL 1', 'Perempuan']);
    sheetSiswa.appendRow(['1003', 'Budi Santoso', 'XII TKJ 2', 'Laki-laki']);
    sheetSiswa.appendRow(['1004', 'Dewi Lestari', 'XII AKL 1', 'Perempuan']);
  }
  
  // 2. Inisialisasi Sheet Kehadiran Log
  let sheetAbsen = ss.getSheetByName('Kehadiran');
  if (!sheetAbsen) {
    sheetAbsen = ss.insertSheet('Kehadiran');
    sheetAbsen.appendRow(['Tanggal', 'Waktu', 'NIS', 'Nama', 'Kelas', 'Status']);
  }
  
  // 3. Inisialisasi Sheet Pengaturan Sekolah
  let sheetSetting = ss.getSheetByName('Pengaturan');
  if (!sheetSetting) {
    sheetSetting = ss.insertSheet('Pengaturan');
    sheetSetting.appendRow(['Key', 'Value']);
    sheetSetting.appendRow(['namaSekolah', 'SMK Negeri 1 Jakarta']);
    sheetSetting.appendRow(['tagline', 'E-Absensi Digital System']);
    sheetSetting.appendRow(['logoUrl', 'https://placehold.co/80x80/4f46e5/ffffff?text=Logo']);
  }

  // Reading Siswa Data
  const dataSiswaVals = sheetSiswa.getDataRange().getValues();
  const siswaList = [];
  for (let i = 1; i < dataSiswaVals.length; i++) {
    if (dataSiswaVals[i][0]) {
      siswaList.push({
        nis: String(dataSiswaVals[i][0]),
        nama: String(dataSiswaVals[i][1]),
        kelas: String(dataSiswaVals[i][2]),
        jk: String(dataSiswaVals[i][3])
      });
    }
  }

  // Reading Settings Data
  const settingsVals = sheetSetting.getDataRange().getValues();
  const settingsObj = {};
  for (let i = 1; i < settingsVals.length; i++) {
    if (settingsVals[i][0]) {
      settingsObj[settingsVals[i][0]] = settingsVals[i][1];
    }
  }

  // Calculating Today's Attendance Statistics
  const timeZone = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd');
  const absenVals = sheetAbsen.getDataRange().getValues();
  
  let hadirCount = 0;
  let izinCount = 0;
  const recordedNis = new Set();

  for (let i = 1; i < absenVals.length; i++) {
    if (!absenVals[i][0]) continue;
    let recDate = absenVals[i][0];
    if (typeof recDate !== 'string') {
      recDate = Utilities.formatDate(new Date(recDate), timeZone, 'yyyy-MM-dd');
    }
    
    if (recDate === todayStr) {
      const nis = String(absenVals[i][2]);
      recordedNis.add(nis);
      const status = String(absenVals[i][5]);
      if (status === 'Hadir') hadirCount++;
      else if (status === 'Izin' || status === 'Sakit') izinCount++;
    }
  }

  const alpaCount = Math.max(0, siswaList.length - recordedNis.size);

  return {
    settings: settingsObj,
    siswa: siswaList,
    stats: {
      total: siswaList.length,
      hadir: hadirCount,
      izin: izinCount,
      alpa: alpaCount
    }
  };
}

/**
 * Mencatat kehadiran siswa berdasarkan NIS
 */
function saveAbsensi(nis) {
  try {
    const ss = getSpreadsheet();
    const sheetSiswa = ss.getSheetByName('Siswa');
    const sheetAbsen = ss.getSheetByName('Kehadiran');
    
    const siswaVals = sheetSiswa.getDataRange().getValues();
    let foundSiswa = null;

    for (let i = 1; i < siswaVals.length; i++) {
      if (String(siswaVals[i][0]) === String(nis)) {
        foundSiswa = {
          nis: String(siswaVals[i][0]),
          nama: String(siswaVals[i][1]),
          kelas: String(siswaVals[i][2])
        };
        break;
      }
    }

    if (!foundSiswa) {
      return { success: false, message: 'Data NIS (' + nis + ') tidak ditemukan di database!' };
    }

    const timeZone = Session.getScriptTimeZone();
    const now = new Date();
    const tglStr = Utilities.formatDate(now, timeZone, 'yyyy-MM-dd');
    const waktuStr = Utilities.formatDate(now, timeZone, 'HH:mm:ss');
    
    // Validasi pencegahan double scan di hari yang sama
    const absenVals = sheetAbsen.getDataRange().getValues();
    for (let i = 1; i < absenVals.length; i++) {
      if (!absenVals[i][0]) continue;
      let recDate = absenVals[i][0];
      if (typeof recDate !== 'string') {
        recDate = Utilities.formatDate(new Date(recDate), timeZone, 'yyyy-MM-dd');
      }
      
      if (recDate === tglStr && String(absenVals[i][2]) === String(nis)) {
        return { 
          success: false, 
          message: foundSiswa.nama + ' (' + foundSiswa.nis + ') sudah mencatat absensi hari ini!' 
        };
      }
    }

    // Append record baru ke sheet Kehadiran
    sheetAbsen.appendRow([tglStr, waktuStr, foundSiswa.nis, foundSiswa.nama, foundSiswa.kelas, 'Hadir']);
    
    return { 
      success: true, 
      nis: foundSiswa.nis, 
      nama: foundSiswa.nama,
      waktu: waktuStr
    };
  } catch (err) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + err.message };
  }
}

/**
 * Menambah atau memperbarui data siswa
 */
function saveSiswaData(siswaObj, originalNis) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Siswa');
    const values = sheet.getDataRange().getValues();

    if (originalNis) {
      // Update data eksisting
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(originalNis)) {
          sheet.getRange(i + 1, 1, 1, 4).setValues([[
            siswaObj.nis, 
            siswaObj.nama, 
            siswaObj.kelas, 
            siswaObj.jk
          ]]);
          return { success: true, message: 'Data siswa berhasil diperbarui!' };
        }
      }
    } else {
      // Cek duplikasi NIS baru
      for (let i = 1; i < values.length; i++) {
        if (String(values[i][0]) === String(siswaObj.nis)) {
          return { success: false, message: 'NIS ' + siswaObj.nis + ' sudah terdaftar!' };
        }
      }
      sheet.appendRow([siswaObj.nis, siswaObj.nama, siswaObj.kelas, siswaObj.jk]);
      return { success: true, message: 'Data siswa baru berhasil ditambahkan!' };
    }
  } catch (err) {
    return { success: false, message: 'Error: ' + err.message };
  }
}

/**
 * Menghapus data siswa dari Spreadsheet berdasarkan NIS
 */
function deleteSiswaData(nis) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Siswa');
    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]) === String(nis)) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Data siswa NIS ' + nis + ' berhasil dihapus.' };
      }
    }
    return { success: false, message: 'NIS tidak ditemukan.' };
  } catch (err) {
    return { success: false, message: 'Gagal menghapus data: ' + err.message };
  }
}

/**
 * Menyimpan pengaturan identitas sekolah ke Spreadsheet
 */
function saveSettingsData(settingsObj) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Pengaturan');
    if (!sheet) {
      sheet = ss.insertSheet('Pengaturan');
    }
    sheet.clear();
    sheet.appendRow(['Key', 'Value']);
    for (let key in settingsObj) {
      sheet.appendRow([key, settingsObj[key]]);
    }
    return { success: true, message: 'Pengaturan sekolah tersimpan!' };
  } catch (err) {
    return { success: false, message: 'Gagal menyimpan pengaturan: ' + err.message };
  }
}