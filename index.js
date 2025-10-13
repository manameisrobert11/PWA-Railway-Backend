// backend/index.js
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import sqlite3pkg from 'sqlite3';

const __dirname = process.cwd();

const app = express();
app.use(cors());
app.use(express.json());

// --- SQLite setup ---
const sqlite3 = sqlite3pkg.verbose();
const DB_PATH = path.join(__dirname, 'rail_scans.db');
const db = new sqlite3.Database(DB_PATH);

// Create table if not exists
db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS scans (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      serial TEXT,
      stage TEXT,
      operator TEXT,
      wagon1Id TEXT,
      wagon2Id TEXT,
      wagon3Id TEXT,
      grade TEXT,
      railType TEXT,
      spec TEXT,
      lengthM TEXT,
      timestamp TEXT,
      raw TEXT
    )
  `);
});

// --- File storage setup ---
const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const upload = multer({ dest: UPLOAD_DIR });

// --- API routes ---
// Add new scan
app.post('/api/scan', (req, res) => {
  const {
    serial,
    stage,
    operator,
    wagon1Id,
    wagon2Id,
    wagon3Id,
    grade,
    railType,
    spec,
    lengthM,
    timestamp,
    raw
  } = req.body;

  if (!serial) return res.status(400).json({ error: 'Serial required' });

  const stmt = db.prepare(`
    INSERT INTO scans 
    (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id, grade, railType, spec, lengthM, timestamp, raw)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  stmt.run(
    serial,
    stage || 'received',
    operator || 'unknown',
    wagon1Id || '',
    wagon2Id || '',
    wagon3Id || '',
    grade || '',
    railType || '',
    spec || '',
    lengthM || '',
    timestamp || new Date().toISOString(),
    raw || '',
    function (err) {
      if (err) return res.status(500).json({ error: err.message });
      res.json({ ok: true, id: this.lastID });
    }
  );

  stmt.finalize();
});

// Get staged scans
app.get('/api/staged', (_req, res) => {
  db.all('SELECT * FROM scans ORDER BY id DESC', (_err, rows) => {
    res.json(rows);
  });
});

// Delete scan
app.delete('/api/staged/:id', (req, res) => {
  const { id } = req.params;
  db.run('DELETE FROM scans WHERE id = ?', [id], function(err) {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ ok: true, id });
  });
});

// Clear all scans
app.post('/api/staged/clear', (_req, res) => {
  db.run('DELETE FROM scans', function(err) {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

// Upload template (.xlsm)
app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file.path });
});

// Export to Excel
app.post('/api/export-to-excel', (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, 'template.xlsm');
    if (!fs.existsSync(templatePath)) return res.status(400).json({ error: 'template.xlsm not found' });

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const existing = XLSX.utils.sheet_to_json(ws, { defval: '' });

    db.all('SELECT * FROM scans', (_err, scans) => {
      const appended = existing.concat(
        scans.map(s => ({
          Serial: s.serial,
          Stage: s.stage,
          Operator: s.operator,
          Wagon1: s.wagon1Id,
          Wagon2: s.wagon2Id,
          Wagon3: s.wagon3Id,
          Grade: s.grade,
          RailType: s.railType,
          Spec: s.spec,
          LengthM: s.lengthM,
          Timestamp: s.timestamp
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
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

// --- Start server ---
const PORT = process.env.PORT || 4000;
app.listen(PORT, '0.0.0.0', () => console.log(`Backend running on :${PORT}`));
