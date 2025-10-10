// backend/index.js
// index.js (ESM)
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import sqlite3pkg from 'sqlite3';  // Using sqlite3 instead of better-sqlite3

const __dirname = process.cwd(); // Render runs from repo root

const app = express();
app.use(cors());
app.use(express.json());

// --- DB setup ---
const DB_PATH = path.join(process.cwd(), 'rail_scans.db');  // Path for the SQLite database file
const db = new sqlite3pkg.Database(DB_PATH);  // Creating SQLite database connection

// Create table if not exists
db.serialize(() => {
  db.run(`
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
  `);
});

// --- File storage setup for Excel ---
const UPLOAD_DIR = path.join(process.cwd(), 'uploads'); // Directory for uploaded files
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR); // Create uploads folder if not exists

const upload = multer({ dest: UPLOAD_DIR }); // Set up file upload middleware

// --- API Routes ---
// Add new scan (POST)
app.post('/api/scan', (req, res) => {
  const { serial, stage, operator, loadId, wagon1, wagon2, wagon3, timestamp, grade, railType, spec, lengthM } = req.body;

  if (!serial) return res.status(400).json({ error: 'Serial required' });

  try {
    const stmt = db.prepare(`
      INSERT INTO scans (serial, stage, operator, loadId, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, timestamp)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const result = stmt.run(
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
      timestamp || new Date().toISOString()
    );
    res.json({ ok: true, id: result.lastInsertRowid });
  } catch (e) {
    console.error('DB insert error', e);
    res.status(500).json({ error: 'Database error' });
  }
});

// Get all scans (GET)
app.get('/api/staged', (req, res) => {
  db.all(`SELECT * FROM scans ORDER BY id DESC`, (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// --- Delete a specific staged scan by ID ---
app.delete('/api/staged/:id', (req, res) => {
  const scanId = req.params.id;

  if (!scanId) {
    return res.status(400).json({ error: 'Scan ID is required' });
  }

  try {
    // Check if the scan exists
    const scanExists = db.prepare('SELECT * FROM scans WHERE id = ?').get(scanId);
    if (!scanExists) {
      return res.status(404).json({ error: 'Scan not found' });
    }

    // Delete the scan
    const stmt = db.prepare('DELETE FROM scans WHERE id = ?');
    const result = stmt.run(scanId);

    if (result.changes === 0) {
      return res.status(404).json({ error: 'Scan not found' });
    }

    res.json({ ok: true, message: 'Scan removed successfully' });
  } catch (e) {
    console.error('Error removing scan', e);
    res.status(500).json({ error: 'Failed to remove scan' });
  }
});

// Clear all scans (POST)
app.post('/api/staged/clear', (req, res) => {
  db.run(`DELETE FROM scans`, (err) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

// Upload Excel template (POST)
app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// Export to Excel (.xlsm) (POST)
app.post('/api/export-to-excel', async (req, res) => {
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

// --- Health check ---
// Check if the server and database are working
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// --- Start server ---
const PORT = process.env.PORT || 4000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Backend on :${PORT}`);
});
