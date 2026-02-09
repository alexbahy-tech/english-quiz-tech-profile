/**
 * ============================================================
 * GOOGLE APPS SCRIPT - EXAM RESULTS RECEIVER (ENHANCED)
 * SMK Negeri 1 Maluku Tengah - English Test System
 * Version 2.0 - Dengan Fitur Kelas
 * ============================================================
 * 
 * CARA SETUP:
 * 1. Buka Google Sheets baru
 * 2. Extensions > Apps Script
 * 3. Copy semua kode ini ke Apps Script Editor
 * 4. Save (Ctrl+S)
 * 5. Deploy > New Deployment > Web App
 * 6. Execute as: Me
 * 7. Who has access: Anyone
 * 8. Copy URL deployment ke file HTML (CONFIG.GOOGLE_SCRIPT_URL)
 * 
 * ============================================================
 */

// ============================================================
// KONFIGURASI
// ============================================================
const CONFIG = {
  SHEET_NAME: 'Hasil Ujian',              // Nama sheet untuk data ujian
  STATS_SHEET_NAME: 'Statistik',          // Nama sheet untuk statistik
  CLASS_STATS_SHEET_NAME: 'Statistik Per Kelas', // Nama sheet untuk statistik per kelas
  SEND_EMAIL_NOTIFICATION: false,          // Set true untuk aktifkan email
  ADMIN_EMAIL: 'alexbahy@500gb.cloud',    // Email admin (ganti dengan email Anda)
  TIMEZONE: 'Asia/Jakarta',               // Timezone Indonesia
  DATE_FORMAT: 'dd/MM/yyyy HH:mm:ss'      // Format tanggal Indonesia
};

// ============================================================
// FUNGSI UTAMA - TERIMA DATA POST DARI WEB APP
// ============================================================
function doPost(e) {
  try {
    Logger.log('📨 Request diterima');
    
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request tidak valid - tidak ada data');
    }
    
    const data = JSON.parse(e.postData.contents);
    Logger.log('📦 Data parsed: ' + JSON.stringify(data));
    
    validateData(data);
    const rowNumber = saveToSheet(data);
    updateStatistics();
    updateClassStatistics();
    
    if (CONFIG.SEND_EMAIL_NOTIFICATION) {
      sendEmailNotification(data);
    }
    
    Logger.log('✅ Data berhasil disimpan di baris: ' + rowNumber);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Data berhasil disimpan',
        row: rowNumber,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    Logger.log('❌ ERROR: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: error.toString(),
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// VALIDASI DATA
// ============================================================
function validateData(data) {
  const requiredFields = ['nama', 'kelas', 'benar', 'total', 'nilai'];
  
  for (const field of requiredFields) {
    if (data[field] === undefined || data[field] === null || data[field] === '') {
      throw new Error(`Field '${field}' tidak ditemukan atau kosong`);
    }
  }
  
  if (typeof data.nama !== 'string' || data.nama.trim() === '') {
    throw new Error('Nama siswa tidak valid');
  }
  
  if (typeof data.kelas !== 'string' || data.kelas.trim() === '') {
    throw new Error('Kelas tidak valid');
  }
  
  if (typeof data.benar !== 'number' || data.benar < 0) {
    throw new Error('Jumlah jawaban benar tidak valid');
  }
  
  if (typeof data.total !== 'number' || data.total <= 0) {
    throw new Error('Total soal tidak valid');
  }
  
  if (typeof data.nilai !== 'number' || data.nilai < 0 || data.nilai > 100) {
    throw new Error('Nilai tidak valid (harus 0-100)');
  }
  
  Logger.log('✅ Validasi data berhasil');
}

// ============================================================
// SIMPAN DATA KE GOOGLE SHEETS
// ============================================================
function saveToSheet(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      createHeaders(sheet);
      Logger.log('📄 Sheet baru dibuat dengan header');
    }
    
    if (sheet.getLastRow() === 0) {
      createHeaders(sheet);
    }
    
    const timestamp = new Date();
    const row = [
      timestamp,                              // A: Timestamp
      data.nama || 'Unknown',                 // B: Nama Siswa
      data.kelas || 'N/A',                    // C: Kelas
      data.benar || 0,                        // D: Jawaban Benar
      data.total || 0,                        // E: Total Soal
      data.nilai || 0,                        // F: Nilai
      data.waktuPengerjaan || 'N/A',          // G: Waktu Pengerjaan (menit)
      data.pelanggaran || 'Tidak ada',        // H: Pelanggaran
      data.waktuSelesai || 'Manual',          // I: Status Selesai
      getStatusLulus(data.nilai),             // J: Status (LULUS/TIDAK LULUS)
      getPredikat(data.nilai)                 // K: Predikat (A/B/C/D/E)
    ];
    
    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    formatRow(sheet, lastRow);
    
    return lastRow;
    
  } catch (error) {
    throw new Error('Gagal menyimpan ke sheet: ' + error.message);
  }
}

// ============================================================
// BUAT HEADER TABEL
// ============================================================
function createHeaders(sheet) {
  const headers = [
    'Timestamp',
    'Nama Siswa',
    'Kelas',
    'Jawaban Benar',
    'Total Soal',
    'Nilai',
    'Waktu Pengerjaan (menit)',
    'Pelanggaran',
    'Status Selesai',
    'Status',
    'Predikat'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#667eea');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 150);  // Timestamp
  sheet.setColumnWidth(2, 200);  // Nama
  sheet.setColumnWidth(3, 100);  // Kelas
  sheet.setColumnWidth(4, 120);  // Benar
  sheet.setColumnWidth(5, 100);  // Total
  sheet.setColumnWidth(6, 80);   // Nilai
  sheet.setColumnWidth(7, 180);  // Waktu Pengerjaan
  sheet.setColumnWidth(8, 250);  // Pelanggaran
  sheet.setColumnWidth(9, 130);  // Status Selesai
  sheet.setColumnWidth(10, 120); // Status
  sheet.setColumnWidth(11, 100); // Predikat
  
  Logger.log('✅ Header tabel dibuat');
}

// ============================================================
// FORMAT ROW DATA
// ============================================================
function formatRow(sheet, rowNumber) {
  sheet.getRange(rowNumber, 1).setNumberFormat(CONFIG.DATE_FORMAT);
  
  const nilaiRange = sheet.getRange(rowNumber, 6);
  nilaiRange.setFontWeight('bold');
  nilaiRange.setHorizontalAlignment('center');
  
  const nilai = sheet.getRange(rowNumber, 6).getValue();
  if (nilai >= 75) {
    nilaiRange.setBackground('#D5F5E3');
    nilaiRange.setFontColor('#1E8449');
  } else {
    nilaiRange.setBackground('#FADBD8');
    nilaiRange.setFontColor('#922B21');
  }
  
  const statusRange = sheet.getRange(rowNumber, 10);
  statusRange.setHorizontalAlignment('center');
  statusRange.setFontWeight('bold');
  
  if (nilai >= 75) {
    statusRange.setBackground('#4facfe');
    statusRange.setFontColor('#FFFFFF');
  } else {
    statusRange.setBackground('#fa709a');
    statusRange.setFontColor('#FFFFFF');
  }
  
  const predikatRange = sheet.getRange(rowNumber, 11);
  predikatRange.setHorizontalAlignment('center');
  predikatRange.setFontWeight('bold');
  
  // Kelas cell
  const kelasRange = sheet.getRange(rowNumber, 3);
  kelasRange.setHorizontalAlignment('center');
  kelasRange.setFontWeight('bold');
  kelasRange.setBackground('#f0f0f0');
  
  if (rowNumber % 2 === 0) {
    sheet.getRange(rowNumber, 1, 1, 11).setBackground('#F8F9FA');
  }
}

// ============================================================
// HELPER FUNCTIONS
// ============================================================
function getStatusLulus(nilai) {
  return nilai >= 75 ? 'LULUS' : 'TIDAK LULUS';
}

function getPredikat(nilai) {
  if (nilai >= 90) return 'A';
  if (nilai >= 80) return 'B';
  if (nilai >= 70) return 'C';
  if (nilai >= 60) return 'D';
  return 'E';
}

// ============================================================
// UPDATE STATISTIK UMUM
// ============================================================
function updateStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!dataSheet || dataSheet.getLastRow() <= 1) {
      return;
    }
    
    let statsSheet = ss.getSheetByName(CONFIG.STATS_SHEET_NAME);
    
    if (!statsSheet) {
      statsSheet = ss.insertSheet(CONFIG.STATS_SHEET_NAME);
    }
    
    statsSheet.clear();
    
    const dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 11);
    const values = dataRange.getValues();
    
    const totalSiswa = values.length;
    const nilaiArray = values.map(row => row[5]); // Kolom F (Nilai)
    
    const totalLulus = values.filter(row => row[5] >= 75).length;
    const totalTidakLulus = totalSiswa - totalLulus;
    const persenLulus = ((totalLulus / totalSiswa) * 100).toFixed(1);
    
    const nilaiTertinggi = Math.max(...nilaiArray);
    const nilaiTerendah = Math.min(...nilaiArray);
    const rataRata = (nilaiArray.reduce((a, b) => a + b, 0) / totalSiswa).toFixed(2);
    
    const predikatA = values.filter(row => row[5] >= 90).length;
    const predikatB = values.filter(row => row[5] >= 80 && row[5] < 90).length;
    const predikatC = values.filter(row => row[5] >= 70 && row[5] < 80).length;
    const predikatD = values.filter(row => row[5] >= 60 && row[5] < 70).length;
    const predikatE = values.filter(row => row[5] < 60).length;
    
    // Hitung pelanggaran
    const totalPelanggaran = values.filter(row => row[7] !== 'Tidak ada').length;
    
    const stats = [
      ['📊 STATISTIK UJIAN ENGLISH TEST', ''],
      ['Terakhir diupdate:', new Date()],
      ['', ''],
      ['📈 RINGKASAN UMUM', ''],
      ['Total Siswa', totalSiswa],
      ['Siswa Lulus (≥75)', totalLulus],
      ['Siswa Tidak Lulus (<75)', totalTidakLulus],
      ['Persentase Kelulusan', persenLulus + '%'],
      ['', ''],
      ['📊 NILAI', ''],
      ['Nilai Tertinggi', nilaiTertinggi],
      ['Nilai Terendah', nilaiTerendah],
      ['Rata-rata Nilai', rataRata],
      ['', ''],
      ['🏆 DISTRIBUSI PREDIKAT', ''],
      ['Predikat A (90-100)', predikatA],
      ['Predikat B (80-89)', predikatB],
      ['Predikat C (70-79)', predikatC],
      ['Predikat D (60-69)', predikatD],
      ['Predikat E (<60)', predikatE],
      ['', ''],
      ['⚠️ PELANGGARAN', ''],
      ['Total Siswa dengan Pelanggaran', totalPelanggaran]
    ];
    
    statsSheet.getRange(1, 1, stats.length, 2).setValues(stats);
    
    // Format
    statsSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold').setBackground('#667eea').setFontColor('#FFFFFF');
    statsSheet.getRange('A1:B1').merge();
    statsSheet.setColumnWidth(1, 280);
    statsSheet.setColumnWidth(2, 150);
    
    // Bold untuk kategori
    const categoryRows = [4, 10, 15, 22];
    categoryRows.forEach(row => {
      statsSheet.getRange(row, 1, 1, 2).setFontWeight('bold').setBackground('#E8F4F8');
    });
    
    Logger.log('✅ Statistik umum diupdate');
    
  } catch (error) {
    Logger.log('⚠️ Gagal update statistik: ' + error.message);
  }
}

// ============================================================
// UPDATE STATISTIK PER KELAS
// ============================================================
function updateClassStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!dataSheet || dataSheet.getLastRow() <= 1) {
      return;
    }
    
    let classStatsSheet = ss.getSheetByName(CONFIG.CLASS_STATS_SHEET_NAME);
    
    if (!classStatsSheet) {
      classStatsSheet = ss.insertSheet(CONFIG.CLASS_STATS_SHEET_NAME);
    }
    
    classStatsSheet.clear();
    
    const dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 11);
    const values = dataRange.getValues();
    
    // Group by class
    const classSummary = {};
    
    values.forEach(row => {
      const kelas = row[2]; // Kolom C (Kelas)
      const nilai = row[5]; // Kolom F (Nilai)
      
      if (!classSummary[kelas]) {
        classSummary[kelas] = {
          total: 0,
          lulus: 0,
          tidakLulus: 0,
          nilaiArray: [],
          predikatA: 0,
          predikatB: 0,
          predikatC: 0,
          predikatD: 0,
          predikatE: 0
        };
      }
      
      classSummary[kelas].total++;
      classSummary[kelas].nilaiArray.push(nilai);
      
      if (nilai >= 75) classSummary[kelas].lulus++;
      else classSummary[kelas].tidakLulus++;
      
      if (nilai >= 90) classSummary[kelas].predikatA++;
      else if (nilai >= 80) classSummary[kelas].predikatB++;
      else if (nilai >= 70) classSummary[kelas].predikatC++;
      else if (nilai >= 60) classSummary[kelas].predikatD++;
      else classSummary[kelas].predikatE++;
    });
    
    // Create header
    const headers = [
      ['📚 STATISTIK PER KELAS', '', '', '', '', '', '', '', '', '']
    ];
    classStatsSheet.getRange(1, 1, 1, 10).setValues(headers);
    classStatsSheet.getRange('A1:J1').merge();
    classStatsSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold')
      .setBackground('#667eea').setFontColor('#FFFFFF').setHorizontalAlignment('center');
    
    const tableHeaders = [
      'Kelas',
      'Total Siswa',
      'Lulus',
      'Tidak Lulus',
      '% Lulus',
      'Rata-rata',
      'Tertinggi',
      'Terendah',
      'Predikat A',
      'Predikat B-E'
    ];
    
    classStatsSheet.getRange(3, 1, 1, 10).setValues([tableHeaders]);
    classStatsSheet.getRange(3, 1, 1, 10)
      .setFontWeight('bold')
      .setBackground('#764ba2')
      .setFontColor('#FFFFFF')
      .setHorizontalAlignment('center');
    
    // Fill data
    let rowIndex = 4;
    const sortedClasses = Object.keys(classSummary).sort();
    
    sortedClasses.forEach(kelas => {
      const data = classSummary[kelas];
      const avg = (data.nilaiArray.reduce((a, b) => a + b, 0) / data.total).toFixed(2);
      const max = Math.max(...data.nilaiArray);
      const min = Math.min(...data.nilaiArray);
      const persenLulus = ((data.lulus / data.total) * 100).toFixed(1);
      const predikatBtoE = data.predikatB + data.predikatC + data.predikatD + data.predikatE;
      
      const row = [
        kelas,
        data.total,
        data.lulus,
        data.tidakLulus,
        persenLulus + '%',
        avg,
        max,
        min,
        data.predikatA,
        predikatBtoE
      ];
      
      classStatsSheet.getRange(rowIndex, 1, 1, 10).setValues([row]);
      
      // Color coding based on average
      const avgRange = classStatsSheet.getRange(rowIndex, 6);
      if (avg >= 80) {
        avgRange.setBackground('#D5F5E3').setFontColor('#1E8449');
      } else if (avg >= 70) {
        avgRange.setBackground('#FCF3CF').setFontColor('#7D6608');
      } else {
        avgRange.setBackground('#FADBD8').setFontColor('#922B21');
      }
      
      // Center align numeric columns
      classStatsSheet.getRange(rowIndex, 2, 1, 9).setHorizontalAlignment('center');
      
      // Alternate row colors
      if (rowIndex % 2 === 0) {
        classStatsSheet.getRange(rowIndex, 1, 1, 10).setBackground('#F8F9FA');
      }
      
      rowIndex++;
    });
    
    // Set column widths
    classStatsSheet.setColumnWidth(1, 120);  // Kelas
    classStatsSheet.setColumnWidth(2, 100);  // Total Siswa
    classStatsSheet.setColumnWidth(3, 80);   // Lulus
    classStatsSheet.setColumnWidth(4, 100);  // Tidak Lulus
    classStatsSheet.setColumnWidth(5, 80);   // % Lulus
    classStatsSheet.setColumnWidth(6, 90);   // Rata-rata
    classStatsSheet.setColumnWidth(7, 90);   // Tertinggi
    classStatsSheet.setColumnWidth(8, 90);   // Terendah
    classStatsSheet.setColumnWidth(9, 90);   // Predikat A
    classStatsSheet.setColumnWidth(10, 100); // Predikat B-E
    
    Logger.log('✅ Statistik per kelas diupdate');
    
  } catch (error) {
    Logger.log('⚠️ Gagal update statistik kelas: ' + error.message);
  }
}

// ============================================================
// KIRIM EMAIL NOTIFIKASI
// ============================================================
function sendEmailNotification(data) {
  try {
    const subject = `📊 Hasil Ujian Baru - ${data.nama} (${data.kelas})`;
    const status = data.nilai >= 75 ? '✅ LULUS' : '❌ TIDAK LULUS';
    
    const body = `
HASIL UJIAN ENGLISH TEST
SMK Negeri 1 Maluku Tengah
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Nama Siswa      : ${data.nama}
Kelas           : ${data.kelas}
Nilai Akhir     : ${data.nilai}
Status          : ${status}

DETAIL:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Jawaban Benar   : ${data.benar} dari ${data.total} soal
Waktu Pengerjaan: ${data.waktuPengerjaan} menit
Pelanggaran     : ${data.pelanggaran}
Status Selesai  : ${data.waktuSelesai}
Predikat        : ${getPredikat(data.nilai)}

Timestamp       : ${new Date().toLocaleString('id-ID', { timeZone: CONFIG.TIMEZONE })}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Lihat data lengkap di Google Sheets.
    `;
    
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: subject,
      body: body
    });
    
    Logger.log('📧 Email notifikasi terkirim ke: ' + CONFIG.ADMIN_EMAIL);
    
  } catch (error) {
    Logger.log('⚠️ Gagal kirim email: ' + error.message);
  }
}

// ============================================================
// FUNGSI UNTUK TESTING
// ============================================================
function testPost() {
  const testData = {
    postData: {
      contents: JSON.stringify({
        nama: 'Budi Santoso',
        kelas: 'X TJKT 1',
        benar: 20,
        total: 25,
        nilai: 80,
        waktuPengerjaan: '35.5',
        pelanggaran: 'Tidak ada',
        waktuSelesai: 'Selesai manual'
      })
    }
  };
  
  const result = doPost(testData);
  Logger.log('📝 Test Result: ' + result.getContent());
}

function testMultipleEntries() {
  const students = [
    { nama: 'Andi Pratama', kelas: 'X TJKT 1', benar: 23, total: 25, nilai: 92, waktuPengerjaan: '32.5', pelanggaran: 'Tidak ada', waktuSelesai: 'Selesai manual' },
    { nama: 'Siti Rahmawati', kelas: 'X TJKT 1', benar: 19, total: 25, nilai: 76, waktuPengerjaan: '38.2', pelanggaran: 'Tidak ada', waktuSelesai: 'Selesai manual' },
    { nama: 'Joko Widodo', kelas: 'X TJKT 2', benar: 15, total: 25, nilai: 60, waktuPengerjaan: '42.0', pelanggaran: '1 pelanggaran: Tab switch', waktuSelesai: 'Selesai manual' },
    { nama: 'Dewi Lestari', kelas: 'XI TJKT 1', benar: 21, total: 25, nilai: 84, waktuPengerjaan: '30.8', pelanggaran: 'Tidak ada', waktuSelesai: 'Selesai manual' },
    { nama: 'Rudi Hermawan', kelas: 'XI TJKT 2', benar: 17, total: 25, nilai: 68, waktuPengerjaan: '40.5', pelanggaran: 'Tidak ada', waktuSelesai: 'Waktu habis' }
  ];
  
  students.forEach(student => {
    const testData = {
      postData: {
        contents: JSON.stringify(student)
      }
    };
    
    const result = doPost(testData);
    Logger.log(`✅ ${student.nama} (${student.kelas}): ${result.getContent()}`);
    Utilities.sleep(500);
  });
  
  Logger.log('🎉 Test selesai! Cek sheet Anda.');
}

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'PERINGATAN!',
    'Apakah Anda yakin ingin menghapus SEMUA data ujian?',
    ui.ButtonSet.YES_NO
  );
  
  if (response === ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (sheet) {
      sheet.clear();
      createHeaders(sheet);
      Logger.log('🗑️ Semua data telah dihapus');
      ui.alert('Data berhasil dihapus!');
    }
  }
}

function exportToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf';
  Logger.log('📄 PDF Export URL: ' + url);
  return url;
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Exam Tools')
    .addItem('🔄 Update Statistik', 'updateStatistics')
    .addItem('📚 Update Statistik Per Kelas', 'updateClassStatistics')
    .addItem('🧪 Test Data Entry', 'testPost')
    .addItem('🧪 Test Multiple Entries', 'testMultipleEntries')
    .addSeparator()
    .addItem('📄 Export ke PDF', 'exportToPDF')
    .addItem('🗑️ Hapus Semua Data', 'clearAllData')
    .addToUi();
}
