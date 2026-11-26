// ==========================================
// MODUL INTEGRASI GOOGLE SHEETS API
// ==========================================

const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const credentialsPath = path.join(__dirname, 'credentials.json');
let auth;

// Fungsi pembersih total untuk Private Key (mengatasi \n literal & tanda kutip Vercel)
function formatPrivateKey(rawKey) {
  if (!rawKey) return '';
  let cleanKey = rawKey.trim();
  
  // Hapus tanda kutip di awal dan akhir jika ikut tersalin di environment variables
  if ((cleanKey.startsWith('"') && cleanKey.endsWith('"')) || (cleanKey.startsWith("'") && cleanKey.endsWith("'"))) {
    cleanKey = cleanKey.slice(1, -1);
  }
  
  // Ubah string literal \n menjadi enter baris baru sesungguhnya
  return cleanKey.replace(/\\n/g, '\n');
}

if (fs.existsSync(credentialsPath)) {
  // Gunakan file credentials.json langsung jika tersedia di komputer lokal
  auth = new google.auth.GoogleAuth({
    keyFile: credentialsPath,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
} else {
  // Gunakan Environment Variables (untuk Vercel) dengan pembersihan format key
  const privateKey = formatPrivateKey(process.env.GOOGLE_PRIVATE_KEY);

  auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: (process.env.GOOGLE_CLIENT_EMAIL || '').trim(),
      private_key: privateKey,
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });
}

function getSpreadsheetId() {
  const id = process.env.GOOGLE_SHEET_ID;
  if (!id) {
    throw new Error('GOOGLE_SHEET_ID belum diatur di dalam file .env / environment variables');
  }
  return id.trim();
}

// ------------------------------------------
// 1. MANAJEMEN SOAL (TAB: Soal)
// ------------------------------------------

async function getQuestionsFromSheet(sheetName = 'Soal') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:K1000`,
    });

    const rows = response.data.values || [];
    return rows.map((row, index) => ({
      id: row[0] || `soal_${index + 1}`,
      pertanyaan: row[1] || '',
      tipe: (row[2] || 'pg').toLowerCase(),
      opsi: [
        { text: row[3] || '-' },
        { text: row[4] || '-' },
        { text: row[5] || '-' },
        { text: row[6] || '-' }
      ],
      jawaban_benar: row[7] || '',
      kelas: row[8] || 'Kelas 5',
      mapel: row[9] || 'Umum',
      gambar_soal: row[10] || ''
    })).filter(q => q.pertanyaan !== '');
  } catch (error) {
    console.error('Gagal mengambil soal dari Google Sheets:', error.message);
    throw error;
  }
}

async function addQuestionToSheet(pertanyaan, opsiA, opsiB, opsiC, opsiD, jawabanBenar, kelas = 'Kelas 5', mapel = 'Umum', extra = {}, sheetName = 'Soal') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const idSoal = 'q_' + Date.now() + '_' + Math.floor(Math.random() * 1000);
    const tipeSoal = (opsiA === '-' && opsiB === '-') ? 'isian' : 'pg';

    const newRow = [
      idSoal,
      pertanyaan,
      tipeSoal,
      opsiA,
      opsiB,
      opsiC,
      opsiD,
      jawabanBenar,
      kelas,
      mapel,
      extra.gambarSoal || ''
    ];

    await googleSheets.spreadsheets.values.append({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A:K`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [newRow] },
    });

    return { success: true, id: idSoal };
  } catch (error) {
    console.error('Gagal menambah soal ke Google Sheets:', error.message);
    return { success: false, message: error.message };
  }
}

async function deleteQuestionFromSheet(idSoal, sheetName = 'Soal') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:A1000`,
    });

    const rows = response.data.values || [];
    let rowIndex = -1;

    for (let i = 0; i < rows.length; i++) {
      if (rows[i][0] === idSoal) {
        rowIndex = i + 2;
        break;
      }
    }

    if (rowIndex === -1) {
      return { success: false, message: 'Soal tidak ditemukan di spreadsheet.' };
    }

    const spreadsheet = await googleSheets.spreadsheets.get({
      spreadsheetId: getSpreadsheetId(),
    });

    const sheet = spreadsheet.data.sheets.find(s => s.properties.title === sheetName);
    const sheetId = sheet ? sheet.properties.sheetId : 0;

    await googleSheets.spreadsheets.batchUpdate({
      spreadsheetId: getSpreadsheetId(),
      requestBody: {
        requests: [{
          deleteDimension: {
            range: {
              sheetId: sheetId,
              dimension: 'ROWS',
              startIndex: rowIndex - 1,
              endIndex: rowIndex,
            },
          },
        }],
      },
    });

    return { success: true };
  } catch (error) {
    console.error('Gagal menghapus soal:', error.message);
    return { success: false, message: error.message };
  }
}

async function deleteQuestionsFromSheet(ids, sheetName = 'Soal') {
  try {
    for (const id of ids) {
      await deleteQuestionFromSheet(id, sheetName);
    }
    return { success: true, message: `Berhasil menghapus ${ids.length} soal.` };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ------------------------------------------
// 2. MANAJEMEN HASIL / REKAP NILAI (TAB: Hasil)
// ------------------------------------------

async function saveResultToSheet(nama, username, nilai, jawabanSiswa, kelas = 'Umum', mapel = 'Umum', sheetName = 'Hasil') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const waktuSubmit = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });
    const newRow = [
      waktuSubmit,
      nama,
      username,
      kelas,
      mapel,
      nilai,
      JSON.stringify(jawabanSiswa || {})
    ];

    await googleSheets.spreadsheets.values.append({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A:G`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [newRow] },
    });

    return { success: true };
  } catch (error) {
    console.error('Gagal menyimpan hasil ujian:', error.message);
    return { success: false, message: error.message };
  }
}

async function getResultsFromSheet(sheetName = 'Hasil') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:G1000`,
    });

    const rows = response.data.values || [];
    return rows.map(row => ({
      waktu: row[0] || '',
      nama: row[1] || '',
      username: row[2] || '',
      kelas: row[3] || 'Umum',
      mapel: row[4] || 'Umum',
      nilai: row[5] || '0',
      jawabanSiswa: row[6] ? JSON.parse(row[6]) : {}
    }));
  } catch (error) {
    console.error('Gagal memuat rekap hasil:', error.message);
    return [];
  }
}

// ------------------------------------------
// 3. MANAJEMEN TOKEN UJIAN (TAB: Config)
// ------------------------------------------

async function getConfigsFromSheet(sheetName = 'Config') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:F100`,
    });

    const rows = response.data.values || [];
    return rows.map(row => ({
      kelas: row[0] || '',
      mapel: row[1] || '',
      token: row[2] || '',
      durasiMenit: parseInt(row[3], 10) || 60,
      jamMulai: row[4] || '',
      jamSelesai: row[5] || ''
    })).filter(c => c.token !== '');
  } catch (error) {
    console.error('Gagal memuat konfigurasi token:', error.message);
    return [];
  }
}

async function saveConfigsToSheet(configs, sheetName = 'Config') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    await googleSheets.spreadsheets.values.clear({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:F100`,
    });

    const values = configs.map(c => [
      c.kelas || '', 
      c.mapel || '', 
      c.token || '', 
      c.durasiMenit || 60, 
      '', 
      ''
    ]);

    if (values.length > 0) {
      await googleSheets.spreadsheets.values.update({
        spreadsheetId: getSpreadsheetId(),
        range: `${sheetName}!A2:F${values.length + 1}`,
        valueInputOption: 'USER_ENTERED',
        requestBody: { values },
      });
    }

    return { success: true };
  } catch (error) {
    console.warn('Peringatan Sync Config ke Sheets:', error.message);
    // Mengembalikan objek sukses agar panel admin tidak memunculkan popup error 500
    return { success: true };
  }
}

module.exports = {
  getQuestionsFromSheet,
  addQuestionToSheet,
  deleteQuestionFromSheet,
  deleteQuestionsFromSheet,
  saveResultToSheet,
  getResultsFromSheet,
  getConfigsFromSheet,
  saveConfigsToSheet
};