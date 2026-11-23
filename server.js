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

// ==========================================
// OPTIMASI RATE-LIMIT: CACHE & QUEUE SYSTEM
// ==========================================

let cachedQuestions = null;
let lastQuestionsFetch = 0;
const CACHE_TTL_MS = 5 * 60 * 1000; // Cache soal aktif selama 5 menit

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
    if (cachedQuestions) {
      console.warn('Menggunakan cache soal cadangan karena limit API:', err.message);
      return cachedQuestions;
    }
    throw err;
  }
}

function invalidateQuestionsCache() {
  cachedQuestions = null;
  lastQuestionsFetch = 0;
}

// Antrean penyimpanan hasil ujian
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
    console.log(`Loaded ${classConfigs.length} token ujian dari Google Sheets.`);
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

app.post('/api/admin/upload-gambar', uploadImage.single('gambar'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ success: false, message: 'File gambar tidak ditemukan.' });
  }
  res.json({ success: true, url: `/uploads/${req.file.filename}` });
});

app.post('/api/soal', async (req, res) => {
  try {
    const { 
      pertanyaan, opsiA, opsiB, opsiC, opsiD, jawabanBenar, kelas, mapel,
      gambarSoal, gambarA, gambarB, gambarC, gambarD 
    } = req.body;

    if (!pertanyaan || !jawabanBenar) {
      return res.status(400).json({ success: false, message: 'Pertanyaan dan kunci jawaban wajib diisi!' });
    }

    const result = await addQuestionToSheet(
      pertanyaan,
      opsiA || '-',
      opsiB || '-',
      opsiC || '-',
      opsiD || '-',
      jawabanBenar,
      kelas || 'Kelas 5',
      mapel || 'Umum',
      { gambarSoal, gambarA, gambarB, gambarC, gambarD }
    );
    if (result.success) invalidateQuestionsCache();
    res.json(result);
  } catch (err) {
    res.status(500).json({ success: false, message: err.message });
  }
});

app.post('/api/admin/import-excel', upload.single('fileExcel'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'File Excel tidak ditemukan.' });

    const targetKelas = req.body.kelas || 'Kelas 5';
    const targetMapel = req.body.mapel || 'Umum';
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);

    let countSuccess = 0;
    for (const row of sheetData) {
      const q = row.pertanyaan || row.Pertanyaan || row.soal || row.Soal;
      const a = row.opsi_a || row.opsiA || row.A || row.a;
      const b = row.opsi_b || row.opsiB || row.B || row.b;
      const c = row.opsi_c || row.opsiC || row.C || row.c;
      const d = row.opsi_d || row.opsiD || row.D || row.d;
      const kunci = row.jawaban_benar || row.jawaban || row.kunci || row.Kunci;
      const rowKelas = row.kelas || row.Kelas || targetKelas;
      const rowMapel = row.mapel || row.Mapel || targetMapel;

      if (q && a && b && c && d && kunci) {
        await addQuestionToSheet(
          q, a, b, c, d, kunci.toString().trim(), rowKelas, rowMapel,
          {
            gambarSoal: row.gambar_soal || '',
            gambarA: row.gambar_a || '',
            gambarB: row.gambar_b || '',
            gambarC: row.gambar_c || '',
            gambarD: row.gambar_d || ''
          }
        );
        countSuccess++;
      }
    }

    if (countSuccess > 0) invalidateQuestionsCache();
    res.json({ success: true, message: `Berhasil mengimpor ${countSuccess} soal (${targetKelas} - ${targetMapel})!` });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Gagal memproses file Excel: ' + err.message });
  }
});

app.post('/api/admin/import-word', upload.single('fileWord'), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ success: false, message: 'File Word tidak ditemukan.' });

    const targetKelas = req.body.kelas || 'Kelas 5';
    const targetMapel = req.body.mapel || 'Umum';
    const result = await mammoth.extractRawText({ buffer: req.file.buffer });
    const text = result.value;

    const blocks = text.split(/(?=\d+[\.\)])/g);
    let countSuccess = 0;

    for (const block of blocks) {
      const lines = block.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      if (lines.length < 5) continue;

      let pertanyaan = lines[0].replace(/^\d+[\.\)]\s*/, '');
      let opsiA = '', opsiB = '', opsiC = '', opsiD = '', jawabanBenar = '';

      lines.forEach(line => {
        if (/^[aA][\.\)]\s*/.test(line)) opsiA = line.replace(/^[aA][\.\)]\s*/, '');
        else if (/^[bB][\.\)]\s*/.test(line)) opsiB = line.replace(/^[bB][\.\)]\s*/, '');
        else if (/^[cC][\.\)]\s*/.test(line)) opsiC = line.replace(/^[cC][\.\)]\s*/, '');
        else if (/^[dD][\.\)]\s*/.test(line)) opsiD = line.replace(/^[dD][\.\)]\s*/, '');
        else if (/^(kunci|jawaban|jawaban_benar):\s*/i.test(line)) {
          jawabanBenar = line.replace(/^(kunci|jawaban|jawaban_benar):\s*/i, '').trim();
        }
      });

      if (pertanyaan && opsiA && opsiB && opsiC && opsiD && jawabanBenar) {
        await addQuestionToSheet(pertanyaan, opsiA, opsiB, opsiC, opsiD, jawabanBenar, targetKelas, targetMapel);
        countSuccess++;
      }
    }

    if (countSuccess > 0) invalidateQuestionsCache();
    res.json({ success: true, message: `Berhasil mengimpor ${countSuccess} soal (${targetKelas} - ${targetMapel})!` });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Gagal memproses file Word: ' + err.message });
  }
});

app.get('/api/admin/config', async (req, res) => {
  const sheetConfigs = await getConfigsFromSheet();
  if (Array.isArray(sheetConfigs)) {
    classConfigs = sheetConfigs;
  }
  res.json({ success: true, configs: classConfigs });
});

app.post('/api/admin/config', async (req, res) => {
  const { kelas, mapel, token, durasiMenit } = req.body;
  if (!kelas || !token || !durasiMenit) {
    return res.status(400).json({ success: false, message: 'Field kelas, token, dan durasi wajib diisi!' });
  }

  const existingIdx = classConfigs.findIndex(c => c.token.toUpperCase() === token.trim().toUpperCase());
  const newConfig = { 
    kelas: kelas.trim(), 
    mapel: (mapel || 'Umum').trim(), 
    token: token.trim().toUpperCase(), 
    durasiMenit: parseInt(durasiMenit, 10) 
  };

  if (existingIdx >= 0) {
    classConfigs[existingIdx] = newConfig;
  } else {
    classConfigs.push(newConfig);
  }

  const saveResult = await saveConfigsToSheet(classConfigs);
  if (!saveResult.success) {
    return res.status(500).json({ 
      success: false, 
      message: 'Gagal menyimpan ke tab Config Google Sheets: ' + saveResult.error 
    });
  }

  res.json({ success: true, message: `Token ${token.toUpperCase()} berhasil disimpan permanen!`, configs: classConfigs });
});

app.delete('/api/admin/config/:token', async (req, res) => {
  const { token } = req.params;
  const initialLength = classConfigs.length;
  classConfigs = classConfigs.filter(c => c.token.toUpperCase() !== token.trim().toUpperCase());

  if (classConfigs.length < initialLength) {
    await saveConfigsToSheet(classConfigs);
    res.json({ success: true, message: `Token "${token.toUpperCase()}" berhasil dihapus!`, configs: classConfigs });
  } else {
    res.status(404).json({ success: false, message: 'Token tidak ditemukan!' });
  }
});

app.get('/api/admin/hasil', async (req, res) => {
  try {
    const hasil = await getResultsFromSheet();
    res.json({ success: true, total: hasil.length, data: hasil });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Gagal memuat rekap nilai: ' + err.message });
  }
});

app.get('/api/admin/export-hasil', async (req, res) => {
  try {
    const { kelas, mapel } = req.query;
    let dataHasil = await getResultsFromSheet();

    if (kelas && kelas !== 'semua') {
      dataHasil = dataHasil.filter(h => h.kelas.toLowerCase() === kelas.toLowerCase());
    }
    if (mapel && mapel !== 'semua') {
      dataHasil = dataHasil.filter(h => h.mapel.toLowerCase() === mapel.toLowerCase());
    }

    const excelRows = dataHasil.map((item, idx) => ({
      'No': idx + 1,
      'Waktu Submit': item.waktu,
      'Nama Siswa': item.nama,
      'NIS / Username': item.username,
      'Kelas': item.kelas,
      'Mata Pelajaran': item.mapel,
      'Nilai Akhir': item.nilai
    }));

    const worksheet = XLSX.utils.json_to_sheet(excelRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Rekap Nilai');

    const buffer = XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' });
    const fileName = `Rekap_Nilai_${kelas || 'Semua'}_${mapel || 'Semua'}.xlsx`.replace(/\s+/g, '_');

    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (err) {
    res.status(500).send('Gagal mengekspor Excel: ' + err.message);
  }
});

// Monitoring & Anti-Cheat Heartbeat (Batas Keluar Maksimal 3 Kali)
app.post('/api/ujian/heartbeat', (req, res) => {
  const { username, nama, kelas, mapel, answeredCount, totalQuestions, warnings, status } = req.body;
  if (!username) return res.status(400).json({ success: false });

  const currentWarnings = warnings || 0;
  
  // Jika peringatan keluar tab sudah mencapai atau melebihi 3 kali
  if (currentWarnings >= 3) {
    activeSessions[username] = {
      username,
      nama: nama || 'Siswa',
      kelas: kelas || '-',
      mapel: mapel || '-',
      answeredCount: answeredCount || 0,
      totalQuestions: totalQuestions || 0,
      warnings: currentWarnings,
      status: 'Diskualifikasi / Keluar 3x',
      lastSeen: Date.now()
    };
    // Mengirim instruksi ke frontend untuk memaksa submit otomatis
    return res.json({ success: true, forceSubmit: true, message: 'Anda telah keluar halaman sebanyak 3 kali. Ujian otomatis dikumpulkan!' });
  }

  activeSessions[username] = {
    username,
    nama: nama || 'Siswa',
    kelas: kelas || '-',
    mapel: mapel || '-',
    answeredCount: answeredCount || 0,
    totalQuestions: totalQuestions || 0,
    warnings: currentWarnings,
    status: status || 'Mengerjakan',
    lastSeen: Date.now()
  };

  res.json({ success: true, forceSubmit: false });
});

app.get('/api/admin/monitoring', (req, res) => {
  const now = Date.now();
  const list = Object.values(activeSessions).map(s => {
    let currentStatus = s.status;
    if (currentStatus === 'Mengerjakan' && (now - s.lastSeen > 45000)) {
      currentStatus = 'Terputus / Idle';
    }
    return {
      ...s,
      status: currentStatus,
      waktuUpdate: new Date(s.lastSeen).toLocaleTimeString('id-ID', { timeZone: 'Asia/Jakarta' })
    };
  });

  list.sort((a, b) => b.lastSeen - a.lastSeen);
  res.json({ success: true, data: list });
});

app.delete('/api/admin/monitoring', (req, res) => {
  activeSessions = {};
  res.json({ success: true, message: 'Log monitoring berhasil dibersihkan.' });
});

// ------------------------------------------
// API SISWA
// ------------------------------------------

app.post('/api/ujian/start', async (req, res) => {
  try {
    const { token } = req.body;
    if (!token) {
      return res.status(400).json({ success: false, message: 'Token ujian tidak boleh kosong!' });
    }

    const matchedConfig = classConfigs.find(c => c.token === token.trim().toUpperCase());

    if (!matchedConfig) {
      return res.status(401).json({ success: false, message: 'Token Ujian salah atau tidak terdaftar!' });
    }

    const allQuestions = await getQuestionsCached();

    const targetKelas = (matchedConfig.kelas || '').trim().toLowerCase();
    const targetMapel = (matchedConfig.mapel || 'Umum').trim().toLowerCase();

    const filteredQuestions = allQuestions.filter(q => {
      const qKelas = (q.kelas || '').trim().toLowerCase();
      const qMapel = (q.mapel || 'Umum').trim().toLowerCase();
      return qKelas === targetKelas && qMapel === targetMapel;
    });

    if (filteredQuestions.length === 0) {
      return res.json({ 
        success: false, 
        message: `Belum ada soal tersedia untuk ${matchedConfig.kelas} (${matchedConfig.mapel || 'Umum'}).` 
      });
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
            // 1. Ambil opsi asli, lalu acak posisi teksnya
            let shuffledOpsiTexts = shuffleArray(q.opsi);
            
            // 2. Berikan label A, B, C, D secara berurutan pada hasil acakan tersebut
            let labeledOpsi = shuffledOpsiTexts.map((opt, idx) => ({
              key: ['A', 'B', 'C', 'D'][idx], // Pastikan selalu A, B, C, D berurutan
              text: opt.text,
              image: opt.image
            }));

            return {
              id: q.id,
              pertanyaan: q.pertanyaan,
              tipe: 'pg',
              gambar_soal: q.gambar_soal || '',
              opsi: labeledOpsi
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
    console.error('Error /api/ujian/start:', error);
    res.status(500).json({ 
      success: false, 
      message: 'Terjadi kesalahan server: ' + error.message 
    });
  }
});

app.post('/api/submit', async (req, res) => {
  try {
    const { nama, username, jawabanSiswa, kelas, mapel } = req.body;
    if (!nama || !username) {
      return res.status(400).json({ success: false, message: 'Data tidak lengkap!' });
    }

    const allQuestions = await getQuestionsCached();
    const targetKelas = (kelas || '').trim().toLowerCase();
    const targetMapel = (mapel || '').trim().toLowerCase();

    const classQuestions = allQuestions.filter(q => 
      (q.kelas || '').trim().toLowerCase() === targetKelas &&
      (q.mapel || '').trim().toLowerCase() === targetMapel
    );

    let benar = 0;
    const targetSoal = classQuestions.length > 0 ? classQuestions : allQuestions;
    const totalSoal = targetSoal.length;
    const jawabanData = jawabanSiswa || {};

    targetSoal.forEach(q => {
      const jawabSiswa = (jawabanData[q.id] || '').toString().trim().toLowerCase();
      const kunci = (q.jawaban_benar || '').toString().trim().toLowerCase();

      if (q.tipe === 'isian') {
        if (jawabSiswa === kunci && jawabSiswa !== '') {
          benar++;
        }
      } else {
        if (jawabanData[q.id] && jawabanData[q.id].toString().toUpperCase() === q.jawaban_benar.toString().toUpperCase()) {
          benar++;
        }
      }
    });

    const nilai = totalSoal > 0 ? parseFloat(((benar / totalSoal) * 100).toFixed(1)) : 0;

    const saveStatus = await enqueueResultSave({
      nama,
      username,
      nilai,
      jawabanSiswa: jawabanData,
      kelas: kelas || 'Umum',
      mapel: mapel || 'Umum'
    });

    if (saveStatus.success) {
      if (activeSessions[username]) {
        activeSessions[username].status = 'Selesai';
        activeSessions[username].answeredCount = totalSoal;
        activeSessions[username].lastSeen = Date.now();
      }
      res.json({ success: true, nilai: nilai, totalSoal: totalSoal, benar: benar });
    } else {
      res.status(500).json({ success: false, message: 'Gagal menyimpan hasil.' });
    }
  } catch (error) {
    console.error('Error /api/submit:', error);
    res.status(500).json({ success: false, message: 'Gagal memproses jawaban.' });
  }
});

app.post('/api/login', async (req, res) => {
  try {
    const { nama, username, token } = req.body;
    
    if (!nama || !username || !token) {
      return res.status(400).json({ success: false, message: 'Semua kolom wajib diisi!' });
    }

    if (classConfigs.length === 0) {
      await initConfigs();
    }

    const inputToken = token.trim().toUpperCase();
    const validConfig = classConfigs.find(
      item => item.token && item.token.toString().trim().toUpperCase() === inputToken
    );

    if (!validConfig) {
      return res.status(400).json({ 
        success: false, 
        message: `Token "${inputToken}" tidak valid atau belum diaktifkan oleh admin.` 
      });
    }

    const sessionData = {
      nama: nama.trim(),
      username: username.trim(),
      kelas: validConfig.kelas || 'Kelas',
      mapel: validConfig.mapel || 'Ujian',
      durasiMenit: validConfig.durasiMenit || 60,
      token: inputToken
    };

    return res.json({ success: true, session: sessionData });

  } catch (err) {
    console.error('Error saat login:', err);
    return res.status(500).json({ success: false, message: 'Terjadi kesalahan internal pada server.' });
  }
});

initConfigs().then(() => {
  app.listen(PORT, () => {
    console.log(`=================================`);
    console.log(`Server CBT SDIT NURUSSIROJ Running`);
    console.log(`Akses Ujian Siswa : http://localhost:${PORT}`);
    console.log(`Akses Admin Guru  : http://localhost:${PORT}/admin.html`);
    console.log(`=================================`);
  });
});