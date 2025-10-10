// backend/index.js
// index.js (ESM)
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import Database from 'better-sqlite3';
import sqlite3pkg from 'sqlite3';

const __dirname = process.cwd(); // Render runs from repo root

const app = express();
app.use(cors());
app.use(express.json());

// --- DB setup ---
const DB_PATH = path.join(process.cwd(), 'rail_scans.db');
const db = new Database(DB_PATH);

// Create table if not exists
db.prepare(`
  CREATE TABLE IF NOT EXISTS scans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    serial TEXT,
    stage TEXT,
    operator TEXT,
    loadId TEXT,
    wagon1 TEXT,
    wagon2 TEXT,
    wagon3 TEXT,
    timestamp TEXT
  )
`).run();

// --- File storage setup for Excel ---
const UPLOAD_DIR = path.join(process.cwd(), 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR);

// --- SQLite setup ---
const sqlite3 = sqlite3pkg.verbose();
const dbSqlite = new sqlite3.Database(DB_PATH);

dbSqlite.serialize(() => {
  dbSqlite.run(`
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
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const upload = multer({ dest: UPLOAD_DIR });

// --- API Routes ---
// Add new scan
app.post('/api/scan', (req, res) => {
  const {
    serial, stage, operator, loadId, wagon1, wagon2, wagon3, timestamp,
    grade, railType, spec, lengthM
  } = req.body;

  if (!serial) return res.status(400).json({ error: 'Serial required' });

  try {
    const stmt = db.prepare(`
      INSERT INTO scans (serial, stage, operator, loadId, wagon1, wagon2, wagon3, timestamp)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const result = stmt.run(serial, stage, operator, loadId, wagon1, wagon2, wagon3, timestamp);
    res.json({ ok: true, id: result.lastInsertRowid });
  } catch (e) {
    console.error('DB insert error', e);
    res.status(500).json({ error: 'Database error' });
  }
});

// Get all scans
app.get('/api/staged', (req, res) => {
  const rows = db.prepare('SELECT * FROM scans ORDER BY id DESC').all();
  res.json(rows);
});

// Clear scans
app.post('/api/staged/clear', (req, res) => {
  db.prepare('DELETE FROM scans').run();
  res.json({ ok: true });
});

// Upload Excel template
app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file.path });
});

// Export to Excel (.xlsm)
app.post('/api/export-to-excel', async (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, 'template.xlsm');
    if (!fs.existsSync(templatePath)) {
      return res.status(400).json({ error: 'template.xlsm not found in uploads/' });
    }

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const existing = XLSX.utils.sheet_to_json(ws, { defval: '' });

    const scans = db.prepare('SELECT * FROM scans').all();
    const appended = existing.concat(scans.map(s => ({
      Serial: s.serial,
      Stage: s.stage,
      Operator: s.operator,
      LoadID: s.loadId,
      Wagon1: s.wagon1,
      Wagon2: s.wagon2,
      Wagon3: s.wagon3,
      Timestamp: s.timestamp,
    })));

    const newWs = XLSX.utils.json_to_sheet(appended, { skipHeader: false });
    wb.Sheets[sheetName] = newWs;

    const outName = `Master_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR, outName);
    XLSX.writeFile(wb, outPath, { bookType: 'xlsm', bookVBA: true });

    res.download(outPath, outName);
  } catch (err) {
    console.error('Export failed:', err);
    res.status(500).json({ error: err.message });
  }
});

// --- Start server ---
const PORT = process.env.PORT || 4000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Backend on :${PORT}`);
});

// Health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});
