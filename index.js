// index.js (ESM)
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import sqlite3pkg from 'sqlite3';

const __dirname = process.cwd(); // Render runs from repo root

const app = express();
app.use(cors());
app.use(express.json());

// --- SQLite setup ---
const sqlite3 = sqlite3pkg.verbose();
const DB_PATH = path.join(__dirname, 'data.db');
const db = new sqlite3.Database(DB_PATH);

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS scans (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      serial   TEXT,
      stage    TEXT,
      operator TEXT,
      loadId   TEXT,
      wagon1   TEXT,
      wagon2   TEXT,
      wagon3   TEXT,
      grade    TEXT,
      railType TEXT,
      spec     TEXT,
      lengthM  TEXT,
      timestamp TEXT
    )
  `);
});

// --- uploads dir ---
const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const upload = multer({ dest: UPLOAD_DIR });

// ---- routes ----

// add scan
app.post('/api/scan', (req, res) => {
  const {
    serial, stage, operator, loadId, wagon1, wagon2, wagon3,
    grade, railType, spec, lengthM, timestamp
  } = req.body || {};

  if (!serial) return res.status(400).json({ error: 'Serial required' });

  const stmt = db.prepare(`
    INSERT INTO scans (serial, stage, operator, loadId, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, timestamp)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  stmt.run(
    serial,
    stage || 'received',
    operator || 'unknown',
    loadId || '',
    wagon1 || '',
    wagon2 || '',
    wagon3 || '',
    grade || '',
    railType || '',
    spec || '',
    lengthM || '',
    timestamp || new Date().toISOString(),
    function (err) {
      if (err) return res.status(500).json({ error: err.message });
      res.json({ ok: true, id: this.lastID });
    }
  );
  stmt.finalize();
});

// list scans
app.get('/api/staged', (_req, res) => {
  db.all(`SELECT * FROM scans ORDER BY id DESC`, (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// clear scans
app.post('/api/staged/clear', (_req, res) => {
  db.run(`DELETE FROM scans`, (err) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

// upload template (.xlsm)
app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// export to Excel (.xlsm) using template
app.post('/api/export-to-excel', (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, 'template.xlsm');
    if (!fs.existsSync(templatePath)) {
      return res.status(400).json({ error: 'template.xlsm not found in uploads/' });
    }

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const existing = XLSX.utils.sheet_to_json(ws, { defval: '' });

    db.all(`SELECT * FROM scans ORDER BY id ASC`, (err, rows) => {
      if (err) return res.status(500).json({ error: err.message });

      const appended = existing.concat(
        rows.map((s) => ({
          'Serial No': s.serial,
          'Stage': s.stage,
          'Operator': s.operator,
          'Load ID': s.loadId,
          'Wagon 1': s.wagon1,
          'Wagon 2': s.wagon2,
          'Wagon 3': s.wagon3,
          'Grade': s.grade,
          'Rail Type': s.railType,
          'Spec': s.spec,
          'Length (m)': s.lengthM,
          'Timestamp': s.timestamp
        }))
      );

      const newWs = XLSX.utils.json_to_sheet(appended, { skipHeader: false });
      wb.Sheets[sheetName] = newWs;

      const outName = `Master_${Date.now()}.xlsm`;
      const outPath = path.join(UPLOAD_DIR, outName);
      XLSX.writeFile(wb, outPath, { bookType: 'xlsm', bookVBA: true });

      res.download(outPath, outName);
    });
  } catch (err) {
    console.error('Export error:', err);
    res.status(500).json({ error: 'Export failed', details: err.message });
  }
});

// health
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// start
const PORT = process.env.PORT || 4000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Backend on :${PORT}`);
});
