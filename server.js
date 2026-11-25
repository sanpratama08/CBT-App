// ==========================================
// SERVER UTAMA APLIKASI CBT SDIT NURUSSIROJ
// ==========================================

const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const multer = require('multer');
const XLSX = require('xlsx');
const mammoth = require('mammoth');
require('dotenv').config();

const { 
  getQuestionsFromSheet, 
  saveResultToSheet, 
  addQuestionToSheet,
  deleteQuestionFromSheet,
  deleteQuestionsFromSheet,
  getResultsFromSheet,
  getConfigsFromSheet,
  saveConfigsToSheet
} = require('./googleSheets');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({ storage: multer.memoryStorage() });

const uploadDir = path.join(__dirname, 'public', 'uploads');
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

const imageStorage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadDir),
  filename: (req, file, cb) => {
    const ext = path.extname(file.originalname);
    cb(null, `img_${Date.now()}_${Math.floor(Math.random() * 1000)}${ext}`);
  }
});
const uploadImage = multer({ storage: imageStorage });

app.use(cors());
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

let cachedQuestions = null;
let lastQuestionsFetch = 0;
const CACHE_TTL_MS = 5 * 60 * 1000;

async function getQuestionsCached(forceRefresh = false) {
  const now = Date.now();
  if (!forceRefresh && cachedQuestions && (now - lastQuestionsFetch < CACHE_TTL_MS)) {
    return cachedQuestions;
  }
  try {
    const freshData = await getQuestionsFromSheet();
    cachedQuestions = freshData;
    lastQuestionsFetch = now;
    return freshData;
  } catch (err) {
    if (cachedQuestions) return cachedQuestions;
    throw err;
  }
}

function invalidateQuestionsCache() {
  cachedQuestions = null;
  lastQuestionsFetch = 0;
}

const submissionQueue = [];
let isProcessingQueue = false;

function enqueueResultSave(payload) {
  return new Promise((resolve) => {
    submissionQueue.push({ payload, resolve });
    processSubmissionQueue();
  });
}

async function processSubmissionQueue() {
  if (isProcessingQueue || submissionQueue.length === 0) return;
  isProcessingQueue = true;

  while (submissionQueue.length > 0) {
    const { payload, resolve } = submissionQueue.shift();
    let result = { success: false };

    for (let attempt = 1; attempt <= 3; attempt++) {
      result = await saveResultToSheet(
        payload.nama,
        payload.username,
        payload.nilai,
        payload.jawabanSiswa,
        payload.kelas,
        payload.mapel
      );
      if (result.success) break;
      await new Promise(r => setTimeout(r, 1200 * attempt));
    }

    resolve(result);
    await new Promise(r => setTimeout(r, 250));
  }

  isProcessingQueue = false;
}

let classConfigs = [];

async function initConfigs() {
  const sheetConfigs = await getConfigsFromSheet();
  if (Array.isArray(sheetConfigs)) {
    classConfigs = sheetConfigs;
  }
}

let activeSessions = {};

function shuffleArray(array) {
  let arr = [...array];
  for (let i = arr.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [arr[i], arr[j]] = [arr[j], arr[i]];
  }
  return arr;
}

// ------------------------------------------
// API ADMIN
// ------------------------------------------

app.get('/api/soal', async (req, res) => {
  try {
    const force = req.query.refresh === 'true';
    const soal = await getQuestionsCached(force);
    res.json({ success: true, total: soal.length, data: soal });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.delete('/api/soal/:id', async (req, res) => {
  try {
    const { id } = req.params;
    const result = await deleteQuestionFromSheet(id);
    if (result.success) invalidateQuestionsCache();
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/soal/bulk-delete', async (req, res) => {
  try {
    const { ids, sheetName } = req.body;
    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ success: false, message: 'Pilih minimal satu soal untuk dihapus!' });
    }
    const result = await deleteQuestionsFromSheet(ids, sheetName || 'Soal');
    if (result.success) invalidateQuestionsCache();
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, message: 'Gagal menghapus soal massal: ' + err.message });
  }
});

// ------------------------------------------
// IMPORT SOAL EXCEL (.XLSX) & WORD (.DOCX)
// ------------------------------------------

app.post('/api/admin/import-excel', upload.single('fileExcel'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'Pilih file Excel (.xlsx) terlebih dahulu!' });
    }

    const kelas = (req.body.kelas || 'Kelas 5').trim();
    const mapel = (req.body.mapel || 'Umum').trim();

    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: '' });

    if (rows.length === 0) {
      return res.status(400).json({ success: false, message: 'File Excel kosong atau format kolom tidak terbaca.' });
    }

    let berhasil = 0;
    for (const row of rows) {
      // Mendukung variasi nama header kolom Excel (besar/kecil)
      const pertanyaan = row.pertanyaan || row.Pertanyaan || row.SOAL || row.Soal || '';
      const opsiA = row.A || row.a || row.opsiA || row['Opsi A'] || '-';
      const opsiB = row.B || row.b || row.opsiB || row['Opsi B'] || '-';
      const opsiC = row.C || row.c || row.opsiC || row['Opsi C'] || '-';
      const opsiD = row.D || row.d || row.opsiD || row['Opsi D'] || '-';
      const jawabanBenar = row.jawaban || row.Jawaban || row['jawaban benar'] || row['Jawaban Benar'] || row.kunci || row.Kunci || '';

      if (pertanyaan && jawabanBenar) {
        await addQuestionToSheet(
          pertanyaan.toString().trim(),
          opsiA.toString().trim(),
          opsiB.toString().trim(),
          opsiC.toString().trim(),
          opsiD.toString().trim(),
          jawabanBenar.toString().trim(),
          kelas,
          mapel
        );
        berhasil++;
      }
    }

    invalidateQuestionsCache();
    res.json({ success: true, message: `Berhasil mengimpor ${berhasil} soal dari Excel untuk ${kelas} (${mapel})!` });
  } catch (err) {
    console.error('Error import Excel:', err);
    res.status(500).json({ success: false, message: 'Gagal memproses file Excel: ' + err.message });
  }
});

app.post('/api/admin/import-word', upload.single('fileWord'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: 'Pilih file Word (.docx) terlebih dahulu!' });
    }

    const kelas = (req.body.kelas || 'Kelas 5').trim();
    const mapel = (req.body.mapel || 'Umum').trim();

    const result = await mammoth.extractRawText({ buffer: req.file.buffer });
    const lines = result.value
      .split(/\r?\n/)
      .map(line => line.trim())
      .filter(line => line.length > 0);

    let questions = [];
    let currentQ = null;

    for (const line of lines) {
      // Lewati header bagian seperti "BAGIAN I..." atau judul dokumen
      if (/^(BAGIAN|CONTOH SOAL|SDIT)/i.test(line)) continue;

      // Deteksi awal nomor soal (contoh: "1. Istilah zakat...")
      const matchSoal = line.match(/^(\d+)[\.\)]\s*(.+)/);
      if (matchSoal) {
        if (currentQ && currentQ.pertanyaan && currentQ.jawabanBenar) {
          questions.push(currentQ);
        }
        currentQ = {
          pertanyaan: matchSoal[2].trim(),
          opsiA: '-',
          opsiB: '-',
          opsiC: '-',
          opsiD: '-',
          jawabanBenar: ''
        };
        continue;
      }

      if (!currentQ) continue;

      // Deteksi opsi A, B, C, D
      const matchOpsi = line.match(/^([A-Da-d])[\.\)]\s*(.+)/);
      if (matchOpsi) {
        const huruf = matchOpsi[1].toUpperCase();
        const teksOpsi = matchOpsi[2].trim();
        if (huruf === 'A') currentQ.opsiA = teksOpsi;
        if (huruf === 'B') currentQ.opsiB = teksOpsi;
        if (huruf === 'C') currentQ.opsiC = teksOpsi;
        if (huruf === 'D') currentQ.opsiD = teksOpsi;
        continue;
      }

      // Deteksi Kunci Jawaban PG ("Kunci: A") atau Isian ("Kunci Teks: Suci")
      const matchKunci = line.match(/^Kunci(?:\s*Teks)?\s*:\s*(.+)/i);
      if (matchKunci) {
        currentQ.jawabanBenar = matchKunci[1].trim();
        continue;
      }

      // Jika pertanyaan lebih dari 1 baris sebelum masuk ke opsi/kunci
      if (currentQ.opsiA === '-' && !currentQ.jawabanBenar) {
        currentQ.pertanyaan += ' ' + line;
      }
    }

    // Simpan soal terakhir
    if (currentQ && currentQ.pertanyaan && currentQ.jawabanBenar) {
      questions.push(currentQ);
    }

    if (questions.length === 0) {
      return res.status(400).json({ 
        success: false, 
        message: 'Tidak ada soal yang terbaca. Pastikan format nomor (1. ...), opsi (A. ...), dan Kunci: ... sudah sesuai.' 
      });
    }

    let berhasil = 0;
    for (const q of questions) {
      await addQuestionToSheet(
        q.pertanyaan,
        q.opsiA,
        q.opsiB,
        q.opsiC,
        q.opsiD,
        q.jawabanBenar,
        kelas,
        mapel
      );
      berhasil++;
    }

    invalidateQuestionsCache();
    res.json({ success: true, message: `Berhasil mengimpor ${berhasil} soal dari Word untuk ${kelas} (${mapel})!` });
  } catch (err) {
    console.error('Error import Word:', err);
    res.status(500).json({ success: false, message: 'Gagal memproses file Word: ' + err.message });
  }
});

app.post('/api/admin/upload-gambar', uploadImage.single('gambar'), (req, res) => {
  if (!req.file) return res.status(400).json({ success: false, message: 'File gambar tidak ditemukan.' });
  res.json({ success: true, url: `/uploads/${req.file.filename}` });
});

app.post('/api/soal', async (req, res) => {
  try {
    const { pertanyaan, opsiA, opsiB, opsiC, opsiD, jawabanBenar, kelas, mapel, gambarSoal } = req.body;
    if (!pertanyaan || !jawabanBenar) {
      return res.status(400).json({ success: false, message: 'Pertanyaan dan kunci jawaban wajib diisi!' });
    }
    const result = await addQuestionToSheet(
      pertanyaan, opsiA || '-', opsiB || '-', opsiC || '-', opsiD || '-',
      jawabanBenar, kelas || 'Kelas 5', mapel || 'Umum', { gambarSoal }
    );
    if (result.success) invalidateQuestionsCache();
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.get('/api/admin/config', async (req, res) => {
  try {
    const sheetConfigs = await getConfigsFromSheet();
    if (Array.isArray(sheetConfigs)) {
      classConfigs = sheetConfigs;
    }
    res.json({ success: true, configs: classConfigs });
  } catch (err) {
    console.error('Error GET config:', err);
    res.status(500).json({ success: false, message: 'Gagal membaca config: ' + err.message, configs: classConfigs });
  }
});

app.post('/api/admin/config', async (req, res) => {
  try {
    const { kelas, mapel, token, durasiMenit, jamMulai, jamSelesai } = req.body;
    if (!kelas || !token || !durasiMenit) {
      return res.status(400).json({ success: false, message: 'Field kelas, token, dan durasi wajib diisi!' });
    }
    
    const existingIdx = classConfigs.findIndex(c => c.token.toUpperCase() === token.trim().toUpperCase());
    const newConfig = { 
      kelas: kelas.trim(), 
      mapel: (mapel || 'Umum').trim(), 
      token: token.trim().toUpperCase(), 
      durasiMenit: parseInt(durasiMenit, 10),
      jamMulai: jamMulai ? jamMulai.trim() : '',
      jamSelesai: jamSelesai ? jamSelesai.trim() : ''
    };

    if (existingIdx >= 0) classConfigs[existingIdx] = newConfig;
    else classConfigs.push(newConfig);

    // Cek hasil penyimpanan ke Google Sheets secara eksplisit
    const saveResult = await saveConfigsToSheet(classConfigs);
    if (saveResult && saveResult.success === false) {
      return res.status(500).json({ 
        success: false, 
        message: 'Gagal menyimpan ke Google Sheets: ' + (saveResult.message || 'Periksa tab Config di Spreadsheet') 
      });
    }

    res.json({ success: true, message: 'Token berkala berhasil disimpan ke Google Sheets!', configs: classConfigs });
  } catch (err) {
    console.error('Error POST config:', err);
    res.status(500).json({ success: false, message: 'Terjadi kesalahan server: ' + err.message });
  }
});
app.delete('/api/admin/config/:token', async (req, res) => {
  const { token } = req.params;
  classConfigs = classConfigs.filter(c => c.token.toUpperCase() !== token.trim().toUpperCase());
  await saveConfigsToSheet(classConfigs);
  res.json({ success: true, message: `Token berhasil dihapus!`, configs: classConfigs });
});

app.get('/api/admin/hasil', async (req, res) => {
  try {
    const hasil = await getResultsFromSheet();
    res.json({ success: true, total: hasil.length, data: hasil });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

// Heartbeat & Anti-Cheat
app.post('/api/ujian/heartbeat', (req, res) => {
  const { username, nama, kelas, mapel, answeredCount, totalQuestions, warnings, status } = req.body;
  if (!username) return res.status(400).json({ success: false });

  const currentWarnings = warnings || 0;
  if (currentWarnings >= 3) {
    activeSessions[username] = { username, nama, kelas, mapel, answeredCount, totalQuestions, warnings: currentWarnings, status: 'Diskualifikasi 3x', lastSeen: Date.now() };
    return res.json({ success: true, forceSubmit: true, message: 'Anda telah keluar halaman sebanyak 3 kali. Ujian otomatis dikumpulkan!' });
  }

  activeSessions[username] = { username, nama, kelas, mapel, answeredCount, totalQuestions, warnings: currentWarnings, status: status || 'Mengerjakan', lastSeen: Date.now() };
  res.json({ success: true, forceSubmit: false });
});

app.get('/api/admin/monitoring', (req, res) => {
  const now = Date.now();
  const list = Object.values(activeSessions).map(s => ({
    ...s,
    status: (s.status === 'Mengerjakan' && (now - s.lastSeen > 45000)) ? 'Terputus / Idle' : s.status,
    waktuUpdate: new Date(s.lastSeen).toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta' })
  }));
  res.json({ success: true, data: list });
});

app.delete('/api/admin/monitoring', (req, res) => {
  activeSessions = {};
  res.json({ success: true, message: 'Monitoring dibersihkan.' });
});

// ------------------------------------------
// API SISWA
// ------------------------------------------

app.post('/api/ujian/start', async (req, res) => {
  try {
    const { token } = req.body;
    if (!token) return res.status(400).json({ success: false, message: 'Token tidak boleh kosong!' });

    const matchedConfig = classConfigs.find(c => c.token === token.trim().toUpperCase());
    if (!matchedConfig) return res.status(401).json({ success: false, message: 'Token Ujian salah atau tidak terdaftar!' });

    const allQuestions = await getQuestionsCached();
    const targetKelas = (matchedConfig.kelas || '').trim().toLowerCase();
    const targetMapel = (matchedConfig.mapel || 'Umum').trim().toLowerCase();

    const filteredQuestions = allQuestions.filter(q => {
      return (q.kelas || '').trim().toLowerCase() === targetKelas && (q.mapel || 'Umum').trim().toLowerCase() === targetMapel;
    });

    if (filteredQuestions.length === 0) {
      return res.json({ success: false, message: `Belum ada soal untuk ${matchedConfig.kelas} (${matchedConfig.mapel}).` });
    }

    let randomized = shuffleArray(filteredQuestions);
    randomized = randomized.map(q => {
      if (q.tipe === 'isian') {
        return {
          id: q.id,
          pertanyaan: q.pertanyaan,
          tipe: 'isian',
          gambar_soal: q.gambar_soal || '',
          opsi: []
        };
      } else {
        // Acak teks pilihan jawaban murni TANPA huruf A, B, C, D
        let shuffledOpsiTexts = shuffleArray(q.opsi.map(o => o.text).filter(t => t && t !== '-'));
        return {
          id: q.id,
          pertanyaan: q.pertanyaan,
          tipe: 'pg',
          gambar_soal: q.gambar_soal || '',
          opsi: shuffledOpsiTexts // Array berisi teks pilihan murni
        };
      }
    });

    res.json({ 
      success: true, 
      total: randomized.length, 
      kelas: matchedConfig.kelas,
      mapel: matchedConfig.mapel || 'Umum',
      durasiMenit: matchedConfig.durasiMenit,
      data: randomized 
    });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/submit', async (req, res) => {
  try {
    const { nama, username, jawabanSiswa, kelas, mapel } = req.body;
    if (!nama || !username) return res.status(400).json({ success: false, message: 'Data tidak lengkap!' });

   allQuestions.forEach(q => {
  const jawabSiswa = (jawabanData[q.id] || '').toString().trim().toLowerCase();
  let rawKunci = (q.jawaban_benar || '').toString().trim();

  if (q.tipe === 'isian') {
    // 1. Pecah kunci jawaban berdasarkan pemisah "|" (atau koma ",")
    // Contoh: "Kairo | Cairo | Al-Qairo" menjadi array ['kairo', 'cairo', 'al-qairo']
    const alternatifKunci = rawKunci
      .split('|') // Bisa diganti atau ditambah .split(/[,|]/) jika ingin support koma
      .map(k => k.trim().toLowerCase());

    // 2. Cek apakah jawaban siswa cocok dengan salah satu alternatif kunci
    if (alternatifKunci.includes(jawabSiswa) && jawabSiswa !== '') {
      benar++;
    }
  } else {
    // Logika untuk Pilihan Ganda (tetap seperti biasa)
    let kunci = rawKunci.toLowerCase();
    if (['a', 'b', 'c', 'd'].includes(kunci)) {
      const idxMap = { a: 0, b: 1, c: 2, d: 3 };
      kunci = (q.opsi[idxMap[kunci]]?.text || '').toLowerCase();
    }
    if (jawabSiswa === kunci && jawabSiswa !== '') {
      benar++;
    }
  }
});

    const nilai = totalSoal > 0 ? parseFloat(((benar / totalSoal) * 100).toFixed(1)) : 0;

    await enqueueResultSave({ nama, username, nilai, jawabanSiswa: jawabanData, kelas: kelas || 'Umum', mapel: mapel || 'Umum' });

    if (activeSessions[username]) {
      activeSessions[username].status = 'Selesai';
    }

    res.json({ success: true, nilai, totalSoal, benar });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

app.post('/api/login', async (req, res) => {
  try {
    const { nama, username, token } = req.body;
    if (!nama || !username || !token) return res.status(400).json({ success: false, message: 'Semua kolom wajib diisi!' });

    if (classConfigs.length === 0) await initConfigs();

    const validConfig = classConfigs.find(item => item.token && item.token.toString().trim().toUpperCase() === token.trim().toUpperCase());
    if (!validConfig) return res.status(400).json({ success: false, message: 'Token Ujian salah atau tidak terdaftar.' });

    // --- VALIDASI WAKTU BERKALA OTOMATIS ---
    if (validConfig.jamMulai && validConfig.jamSelesai) {
      const sekarang = new Date();
      // Ambil waktu jam server dalam format "HH:MM" (WIB / Asia/Jakarta)
      const jamSekarang = sekarang.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit', hour12: false, timeZone: 'Asia/Jakarta' });
      
      if (jamSekarang < validConfig.jamMulai) {
        return res.status(400).json({ 
          success: false, 
          message: `Ujian belum dimulai! Token ini baru aktif pada pukul ${validConfig.jamMulai} WIB.` 
        });
      }
      if (jamSekarang > validConfig.jamSelesai) {
        return res.status(400).json({ 
          success: false, 
          message: `Waktu ujian telah berakhir! Token ini ditutup pada pukul ${validConfig.jamSelesai} WIB.` 
        });
      }
    }
    // ----------------------------------------

    res.json({ 
      success: true, 
      session: { nama: nama.trim(), username: username.trim(), kelas: validConfig.kelas, mapel: validConfig.mapel, durasiMenit: validConfig.durasiMenit, token: validConfig.token } 
    });
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

initConfigs().then(() => {
  app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
});