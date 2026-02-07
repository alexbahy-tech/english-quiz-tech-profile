/**
 * ============================================================
 * GOOGLE APPS SCRIPT - EXAM RESULTS RECEIVER
 * SMK Negeri 1 Maluku Tengah - English Test System
 * ============================================================
 * 
 * CARA SETUP:
 * 1. Copy semua kode ini ke Apps Script Editor
 * 2. Save (Ctrl+S)
 * 3. Deploy > New Deployment > Web App
 * 4. Execute as: Me
 * 5. Who has access: Anyone
 * 6. Copy URL deployment ke file HTML
 * 
 * ============================================================
 */

// ============================================================
// KONFIGURASI
// ============================================================
const CONFIG = {
  SHEET_NAME: 'Hasil Ujian',              // Nama sheet untuk data ujian
  STATS_SHEET_NAME: 'Statistik',          // Nama sheet untuk statistik
  SEND_EMAIL_NOTIFICATION: false,          // Set true untuk aktifkan email
  ADMIN_EMAIL: 'alexbahy@500gb.cloud',       // Email admin (ganti dengan email Anda)
  TIMEZONE: 'Asia/Jakarta',               // Timezone Indonesia
  DATE_FORMAT: 'dd/MM/yyyy HH:mm:ss'      // Format tanggal Indonesia
};

// ============================================================
// FUNGSI UTAMA - TERIMA DATA POST DARI WEB APP
// ============================================================
function doPost(e) {
  try {
    // Log untuk debugging
    Logger.log('📨 Request diterima');
    
    // Validasi request
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request tidak valid - tidak ada data');
    }
    
    // Parse data JSON
    const data = JSON.parse(e.postData.contents);
    Logger.log('📦 Data parsed: ' + JSON.stringify(data));
    
    // Validasi field yang diperlukan
    validateData(data);
    
    // Simpan ke Google Sheets
    const rowNumber = saveToSheet(data);
    
    // Update statistik
    updateStatistics();
    
    // Kirim notifikasi email (jika diaktifkan)
    if (CONFIG.SEND_EMAIL_NOTIFICATION) {
      sendEmailNotification(data);
    }
    
    // Log success
    Logger.log('✅ Data berhasil disimpan di baris: ' + rowNumber);
    
    // Return success response
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        message: 'Data berhasil disimpan',
        row: rowNumber,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    // Log error
    Logger.log('❌ ERROR: ' + error.toString());
    Logger.log('Stack trace: ' + error.stack);
    
    // Return error response
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
  const requiredFields = ['nama', 'benar', 'total', 'nilai'];
  
  for (const field of requiredFields) {
    if (data[field] === undefined || data[field] === null) {
      throw new Error(`Field '${field}' tidak ditemukan atau kosong`);
    }
  }
  
  // Validasi tipe data
  if (typeof data.nama !== 'string' || data.nama.trim() === '') {
    throw new Error('Nama siswa tidak valid');
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
    // Buka spreadsheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    // Buat sheet baru jika belum ada
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_NAME);
      createHeaders(sheet);
      Logger.log('📄 Sheet baru dibuat dengan header');
    }
    
    // Cek apakah ada header, jika tidak buat
    if (sheet.getLastRow() === 0) {
      createHeaders(sheet);
    }
    
    // Siapkan data untuk row baru
    const timestamp = new Date();
    const row = [
      timestamp,                              // A: Timestamp
      data.nama || 'Unknown',                 // B: Nama Siswa
      data.benar || 0,                        // C: Jawaban Benar
      data.total || 0,                        // D: Total Soal
      data.nilai || 0,                        // E: Nilai
      data.waktuPengerjaan || 'N/A',          // F: Waktu Pengerjaan (menit)
      data.pelanggaran || 'Tidak ada',        // G: Pelanggaran
      data.waktuSelesai || 'Manual',          // H: Status Selesai
      getStatusLulus(data.nilai),             // I: Status (LULUS/TIDAK LULUS)
      getPredikat(data.nilai)                 // J: Predikat (A/B/C/D/E)
    ];
    
    // Tambahkan row ke sheet
    sheet.appendRow(row);
    const lastRow = sheet.getLastRow();
    
    // Format cells
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
  
  // Format header
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0B3C5D');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // Freeze header row
  sheet.setFrozenRows(1);
  
  // Set column widths
  sheet.setColumnWidth(1, 150);  // Timestamp
  sheet.setColumnWidth(2, 200);  // Nama
  sheet.setColumnWidth(3, 120);  // Benar
  sheet.setColumnWidth(4, 100);  // Total
  sheet.setColumnWidth(5, 80);   // Nilai
  sheet.setColumnWidth(6, 180);  // Waktu Pengerjaan
  sheet.setColumnWidth(7, 200);  // Pelanggaran
  sheet.setColumnWidth(8, 130);  // Status Selesai
  sheet.setColumnWidth(9, 120);  // Status
  sheet.setColumnWidth(10, 100); // Predikat
  
  Logger.log('✅ Header tabel dibuat');
}

// ============================================================
// FORMAT ROW DATA
// ============================================================
function formatRow(sheet, rowNumber) {
  // Format timestamp
  sheet.getRange(rowNumber, 1).setNumberFormat(CONFIG.DATE_FORMAT);
  
  // Format nilai (bold dan center)
  const nilaiRange = sheet.getRange(rowNumber, 5);
  nilaiRange.setFontWeight('bold');
  nilaiRange.setHorizontalAlignment('center');
  
  // Warna berdasarkan nilai
  const nilai = sheet.getRange(rowNumber, 5).getValue();
  if (nilai >= 75) {
    nilaiRange.setBackground('#D5F5E3'); // Hijau muda (lulus)
    nilaiRange.setFontColor('#1E8449');
  } else {
    nilaiRange.setBackground('#FADBD8'); // Merah muda (tidak lulus)
    nilaiRange.setFontColor('#922B21');
  }
  
  // Format status
  const statusRange = sheet.getRange(rowNumber, 9);
  statusRange.setHorizontalAlignment('center');
  statusRange.setFontWeight('bold');
  
  if (nilai >= 75) {
    statusRange.setBackground('#2ECC71');
    statusRange.setFontColor('#FFFFFF');
  } else {
    statusRange.setBackground('#E74C3C');
    statusRange.setFontColor('#FFFFFF');
  }
  
  // Format predikat
  const predikatRange = sheet.getRange(rowNumber, 10);
  predikatRange.setHorizontalAlignment('center');
  predikatRange.setFontWeight('bold');
  
  // Alternate row colors (zebra striping)
  if (rowNumber % 2 === 0) {
    sheet.getRange(rowNumber, 1, 1, 10).setBackground('#F5F5F5');
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
// UPDATE STATISTIK
// ============================================================
function updateStatistics() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dataSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!dataSheet || dataSheet.getLastRow() <= 1) {
      return; // Tidak ada data
    }
    
    let statsSheet = ss.getSheetByName(CONFIG.STATS_SHEET_NAME);
    
    // Buat sheet statistik jika belum ada
    if (!statsSheet) {
      statsSheet = ss.insertSheet(CONFIG.STATS_SHEET_NAME);
    }
    
    statsSheet.clear();
    
    // Get data range
    const dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 10);
    const values = dataRange.getValues();
    
    // Hitung statistik
    const totalSiswa = values.length;
    const nilaiArray = values.map(row => row[4]); // Kolom E (Nilai)
    
    const totalLulus = values.filter(row => row[4] >= 75).length;
    const totalTidakLulus = totalSiswa - totalLulus;
    const persenLulus = ((totalLulus / totalSiswa) * 100).toFixed(1);
    
    const nilaiTertinggi = Math.max(...nilaiArray);
    const nilaiTerendah = Math.min(...nilaiArray);
    const rataRata = (nilaiArray.reduce((a, b) => a + b, 0) / totalSiswa).toFixed(2);
    
    // Hitung predikat
    const predikatA = values.filter(row => row[4] >= 90).length;
    const predikatB = values.filter(row => row[4] >= 80 && row[4] < 90).length;
    const predikatC = values.filter(row => row[4] >= 70 && row[4] < 80).length;
    const predikatD = values.filter(row => row[4] >= 60 && row[4] < 70).length;
    const predikatE = values.filter(row => row[4] < 60).length;
    
    // Tulis statistik
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
      ['Predikat E (<60)', predikatE]
    ];
    
    statsSheet.getRange(1, 1, stats.length, 2).setValues(stats);
    
    // Format statistik sheet
    statsSheet.getRange(1, 1).setFontSize(14).setFontWeight('bold').setBackground('#0B3C5D').setFontColor('#FFFFFF');
    statsSheet.getRange('A1:B1').merge();
    
    statsSheet.setColumnWidth(1, 250);
    statsSheet.setColumnWidth(2, 150);
    
    // Bold untuk kategori
    statsSheet.getRange('A4').setFontWeight('bold').setBackground('#E8F4F8');
    statsSheet.getRange('A10').setFontWeight('bold').setBackground('#E8F4F8');
    statsSheet.getRange('A15').setFontWeight('bold').setBackground('#E8F4F8');
    
    Logger.log('✅ Statistik diupdate');
    
  } catch (error) {
    Logger.log('⚠️ Gagal update statistik: ' + error.message);
  }
}

// ============================================================
// KIRIM EMAIL NOTIFIKASI
// ============================================================
function sendEmailNotification(data) {
  try {
    const subject = `📊 Hasil Ujian Baru - ${data.nama}`;
    const status = data.nilai >= 75 ? '✅ LULUS' : '❌ TIDAK LULUS';
    
    const body = `
HASIL UJIAN ENGLISH TEST
SMK Negeri 1 Maluku Tengah
━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Nama Siswa      : ${data.nama}
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

// Jalankan fungsi ini untuk test multiple entries
function testMultipleEntries() {
  const students = [
    { nama: 'Andi Pratama', benar: 23, total: 25, nilai: 92, waktuPengerjaan: '32.5', pelanggaran: 'Tidak ada', waktuSelesai: 'Selesai manual' },
    { nama: 'Siti Rahmawati', benar: 19, total: 25, nilai: 76, waktuPengerjaan: '38.2', pelanggaran: 'Tidak ada', waktuSelesai: 'Selesai manual' },
    { nama: 'Joko Widodo', benar: 15, total: 25, nilai: 60, waktuPengerjaan: '42.0', pelanggaran: '1 pelanggaran: Tab switch', waktuSelesai: 'Selesai manual' },
    { nama: 'Dewi Lestari', benar: 21, total: 25, nilai: 84, waktuPengerjaan: '30.8', pelanggaran: 'Tidak ada', waktuSelesai: 'Selesai manual' },
    { nama: 'Rudi Hermawan', benar: 17, total: 25, nilai: 68, waktuPengerjaan: '40.5', pelanggaran: 'Tidak ada', waktuSelesai: 'Waktu habis' }
  ];
  
  students.forEach(student => {
    const testData = {
      postData: {
        contents: JSON.stringify(student)
      }
    };
    
    const result = doPost(testData);
    Logger.log(`✅ ${student.nama}: ${result.getContent()}`);
    Utilities.sleep(500); // Delay 500ms antar request
  });
  
  Logger.log('🎉 Test selesai! Cek sheet Anda.');
}

// ============================================================
// FUNGSI UTILITY TAMBAHAN
// ============================================================

/**
 * Fungsi untuk menghapus semua data (HATI-HATI!)
 */
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

/**
 * Fungsi untuk export ke PDF
 */
function exportToPDF() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=pdf';
  
  Logger.log('📄 PDF Export URL: ' + url);
  
  return url;
}

/**
 * Fungsi untuk membuat menu custom di Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Exam Tools')
    .addItem('🔄 Update Statistik', 'updateStatistics')
    .addItem('🧪 Test Data Entry', 'testPost')
    .addItem('🧪 Test Multiple Entries', 'testMultipleEntries')
    .addSeparator()
    .addItem('📄 Export ke PDF', 'exportToPDF')
    .addItem('🗑️ Hapus Semua Data', 'clearAllData')
    .addToUi();
}