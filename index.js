// backend/index.js
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import Database from 'better-sqlite3';

const app = express();
app.use(cors());
app.use(express.json());

// --- Paths ---
const DATA_DIR = process.cwd();
const DB_PATH = path.join(DATA_DIR, 'rail_scans.db');
const UPLOAD_DIR = path.join(DATA_DIR, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

// --- Database setup ---
const db = new Database(DB_PATH);
db.prepare(`
  CREATE TABLE IF NOT EXISTS scans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    serial TEXT,
    operator TEXT,
    wagon1 TEXT,
    wagon2 TEXT,
    wagon3 TEXT,
    grade TEXT,
    railType TEXT,
    spec TEXT,
    lengthM TEXT,
    stage TEXT,
    timestamp TEXT
  )
`).run();

// --- File upload setup ---
const upload = multer({ dest: UPLOAD_DIR });

// --- API routes ---

// Health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// Add a new scan
app.post('/api/scan', (req, res) => {
  const { serial, operator, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, stage, timestamp } = req.body;
  if (!serial) return res.status(400).json({ error: 'Serial required' });

  try {
    const stmt = db.prepare(`
      INSERT INTO scans 
      (serial, operator, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, stage, timestamp)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const result = stmt.run(
      serial,
      operator || 'Unknown',
      wagon1 || '',
      wagon2 || '',
      wagon3 || '',
      grade || '',
      railType || '',
      spec || '',
      lengthM || '',
      stage || 'received',
      timestamp || new Date().toISOString()
    );
    res.json({ ok: true, id: result.lastInsertRowid });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Database error' });
  }
});

// Get all staged scans
app.get('/api/staged', (_req, res) => {
  const rows = db.prepare('SELECT * FROM scans ORDER BY id DESC').all();
  res.json(rows);
});

// Delete a scan
app.delete('/api/staged/:id', (req, res) => {
  const { id } = req.params;
  try {
    db.prepare('DELETE FROM scans WHERE id = ?').run(id);
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to remove scan' });
  }
});

// Clear all staged scans
app.post('/api/staged/clear', (_req, res) => {
  try {
    db.prepare('DELETE FROM scans').run();
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to clear scans' });
  }
});

// Upload Excel template
app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
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

    const scans = db.prepare('SELECT * FROM scans').all();
    const appended = existing.concat(
      scans.map(s => ({
        Serial: s.serial,
        Operator: s.operator,
        Wagon1: s.wagon1,
        Wagon2: s.wagon2,
        Wagon3: s.wagon3,
        Grade: s.grade,
        RailType: s.railType,
        Spec: s.spec,
        LengthM: s.lengthM,
        Stage: s.stage,
        Timestamp: s.timestamp
      }))
    );

    const newWs = XLSX.utils.json_to_sheet(appended, { skipHeader: false });
    wb.Sheets[sheetName] = newWs;

    const outName = `Master_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR, outName);
    XLSX.writeFile(wb, outPath, { bookType: 'xlsm', bookVBA: true });

    res.download(outPath, outName);
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: 'Failed to export Excel', details: err.message });
  }
});

// --- Start server ---
const PORT = process.env.PORT || 4000;
app.listen(PORT, '0.0.0.0', () => console.log(`✅ Backend running on port ${PORT}`));
