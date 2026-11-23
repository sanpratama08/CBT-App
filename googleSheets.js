// ==========================================
// INTEGRASI GOOGLE SHEETS API - SDIT NURUSSIROJ
// ==========================================

const { google } = require('googleapis');
const path = require('path');
const fs = require('fs');
require('dotenv').config();

const SCOPES = ['https://www.googleapis.com/auth/spreadsheets'];
const KEYFILEPATH = path.join(__dirname, 'credentials.json');

let auth;

// Prioritaskan membaca dari Environment Variable (untuk Vercel / Production)
if (process.env.GOOGLE_CREDENTIALS) {
  const credentials = JSON.parse(process.env.GOOGLE_CREDENTIALS);
  auth = new google.auth.GoogleAuth({
    credentials,
    scopes: SCOPES,
  });
} else if (fs.existsSync(KEYFILEPATH)) {
  // Fallback ke file lokal jika dijalankan di komputer (localhost)
  auth = new google.auth.GoogleAuth({
    keyFile: KEYFILEPATH,
    scopes: SCOPES,
  });
} else {
  throw new Error("Kredensial Google tidak ditemukan! Pasang GOOGLE_CREDENTIALS di Vercel atau sediakan credentials.json di lokal.");
}

function getSpreadsheetId() {
  const id = process.env.GOOGLE_SHEET_ID;
  if (!id) throw new Error("GOOGLE_SHEET_ID belum terpasang di Environment Variables!");
  return id;
}

function normalizeImageUrl(rawUrl) {
  if (!rawUrl || typeof rawUrl !== 'string') return '';
  const url = rawUrl.trim();
  if (!url) return '';

  const driveFileMatch = url.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (driveFileMatch && driveFileMatch[1]) {
    return `https://drive.google.com/thumbnail?id=${driveFileMatch[1]}&sz=w1000`;
  }
  const driveIdMatch = url.match(/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/);
  if (driveIdMatch && driveIdMatch[1]) {
    return `https://drive.google.com/thumbnail?id=${driveIdMatch[1]}&sz=w1000`;
  }

  return url;
}

// Ambil Soal dari Sheet dengan Deteksi Tipe Soal yang Pasti
async function getQuestionsFromSheet(sheetName = 'Soal') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:N`,
    });

    const rows = response.data.values || [];
    
    return rows
      .filter((row) => row[0] && row[1])
      .map((row) => {
        const opsiA = (row[2] || '').trim();
        const opsiB = (row[3] || '').trim();
        const opsiC = (row[4] || '').trim();
        const opsiD = (row[5] || '').trim();
        
        // Deteksi apakah soal ini Isian Singkat
        const isIsian = (!opsiA || opsiA === '-' || opsiA === '') && 
                        (!opsiB || opsiB === '-' || opsiB === '');

        return {
          id: row[0],
          pertanyaan: row[1] || '',
          tipe: isIsian ? 'isian' : 'pg',
          opsi: isIsian ? [] : [
            { text: opsiA, image: normalizeImageUrl(row[10]) },
            { text: opsiB, image: normalizeImageUrl(row[11]) },
            { text: opsiC, image: normalizeImageUrl(row[12]) },
            { text: opsiD, image: normalizeImageUrl(row[13]) }
          ],
          jawaban_benar: row[6] || '',
          kelas: row[7] || 'Umum',
          mapel: row[8] || 'Umum',
          gambar_soal: normalizeImageUrl(row[9])
        };
      });
  } catch (error) {
    console.error(`Gagal mengambil soal dari tab ${sheetName}:`, error.message);
    throw error;
  }
}

// 2. Simpan Hasil Ujian Siswa
async function saveResultToSheet(nama, username, nilai, detailJawaban, kelas = 'Umum', mapel = 'Umum') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });
    const timestamp = new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' });

    await googleSheets.spreadsheets.values.append({
      spreadsheetId: getSpreadsheetId(),
      range: 'Hasil!A:G',
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[
          timestamp,
          nama,
          username,
          nilai,
          JSON.stringify(detailJawaban),
          kelas,
          mapel
        ]],
      },
    });

    return { success: true };
  } catch (error) {
    console.error('Gagal menyimpan hasil ke Google Sheets:', error.message);
    return { success: false, error: error.message };
  }
}

// 3. Tambah Soal Baru
async function addQuestionToSheet(pertanyaan, opsiA, opsiB, opsiC, opsiD, jawabanBenar, kelas = 'Umum', mapel = 'Umum', images = {}, sheetName = 'Soal') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });
    const newId = `Q${Date.now().toString().slice(-6)}`;
    const { gambarSoal = '', gambarA = '', gambarB = '', gambarC = '', gambarD = '' } = images;

    await googleSheets.spreadsheets.values.append({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A:N`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [[
          newId,
          pertanyaan,
          opsiA,
          opsiB,
          opsiC,
          opsiD,
          jawabanBenar.toUpperCase(),
          kelas.trim(),
          mapel.trim(),
          gambarSoal,
          gambarA,
          gambarB,
          gambarC,
          gambarD
        ]],
      },
    });

    return { success: true, id: newId };
  } catch (error) {
    console.error(`Gagal menambah soal ke tab ${sheetName}:`, error.message);
    return { success: false, error: error.message };
  }
}

// 4. Hapus Satu Soal
async function deleteQuestionFromSheet(id, sheetName = 'Soal') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });
    const spreadsheetId = getSpreadsheetId();

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A2:N`,
    });

    const rows = response.data.values || [];
    const filteredRows = rows.filter(row => row[0] !== id);

    await googleSheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${sheetName}!A2:N`,
    });

    if (filteredRows.length > 0) {
      await googleSheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A2`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: filteredRows },
      });
    }

    return { success: true };
  } catch (error) {
    console.error(`Gagal menghapus soal dari tab ${sheetName}:`, error.message);
    return { success: false, error: error.message };
  }
}

// 5. Hapus Banyak Soal Sekaligus (Massal)
async function deleteQuestionsFromSheet(ids = [], sheetName = 'Soal') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });
    const spreadsheetId = getSpreadsheetId();

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A2:N`,
    });

    const rows = response.data.values || [];
    const filteredRows = rows.filter(row => !ids.includes(row[0]));

    await googleSheets.spreadsheets.values.clear({
      spreadsheetId,
      range: `${sheetName}!A2:N`,
    });

    if (filteredRows.length > 0) {
      await googleSheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A2`,
        valueInputOption: 'USER_ENTERED',
        resource: { values: filteredRows },
      });
    }

    return { success: true };
  } catch (error) {
    console.error(`Gagal menghapus soal massal dari tab ${sheetName}:`, error.message);
    return { success: false, error: error.message };
  }
}

// 6. Ambil Rekap Hasil Ujian Siswa dari Sheet 'Hasil'
async function getResultsFromSheet() {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId: getSpreadsheetId(),
      range: 'Hasil!A2:G',
    });

    const rows = response.data.values || [];

    return rows.map((row, index) => ({
      no: index + 1,
      waktu: row[0] || '-',
      nama: row[1] || 'Tanpa Nama',
      username: row[2] || '-',
      nilai: parseFloat(row[3]) || 0,
      kelas: row[5] || 'Umum',
      mapel: row[6] || 'Umum'
    })).reverse();
  } catch (error) {
    console.error('Gagal mengambil rekap hasil dari Google Sheets:', error.message);
    return [];
  }
}

// 7. Ambil & Simpan Konfigurasi Token Permanen di Tab 'Config'
// Membaca Config termasuk Jam Mulai & Selesai
async function getConfigsFromSheet(sheetName = 'Config') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    const response = await googleSheets.spreadsheets.values.get({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:F`,
    });

    const rows = response.data.values || [];
    return rows.map((row) => ({
      kelas: row[0] || '',
      mapel: row[1] || '',
      token: row[2] || '',
      durasiMenit: parseInt(row[3], 10) || 60,
      jamMulai: row[4] || '',   // Format "07:30"
      jamSelesai: row[5] || ''  // Format "09:30"
    }));
  } catch (error) {
    console.error('Gagal mengambil config:', error.message);
    return [];
  }
}

async function saveConfigsToSheet(configs, sheetName = 'Config') {
  try {
    const client = await auth.getClient();
    const googleSheets = google.sheets({ version: 'v4', auth: client });

    // Bersihkan dulu baris lama
    await googleSheets.spreadsheets.values.clear({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:F100`,
    });

    const values = configs.map(c => [c.kelas, c.mapel, c.token, c.durasiMenit, c.jamMulai || '', c.jamSelesai || '']);

    await googleSheets.spreadsheets.values.update({
      spreadsheetId: getSpreadsheetId(),
      range: `${sheetName}!A2:F2 + ${values.length}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values },
    });

    return { success: true };
  } catch (error) {
    console.error('Gagal menyimpan config:', error.message);
    return { success: false, message: error.message };
  }
}

module.exports = {
  getQuestionsFromSheet,
  saveResultToSheet,
  addQuestionToSheet,
  deleteQuestionFromSheet,
  deleteQuestionsFromSheet,
  getResultsFromSheet,
  getConfigsFromSheet,
  saveConfigsToSheet
};