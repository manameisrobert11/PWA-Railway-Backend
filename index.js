// backend/index.js
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import Database from 'better-sqlite3';
import { fileURLToPath } from 'url';

// ESM-safe __dirname
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(cors());
app.use(express.json());

// Allow persistent disk via DATA_DIR; otherwise use current folder
const DATA_DIR = process.env.DATA_DIR || __dirname;

// --- DB setup ---
const DB_PATH = path.join(DATA_DIR, 'rail_scans.db');
const db = new Database(DB_PATH);
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

// --- File storage for Excel ---
const UPLOAD_DIR = path.join(DATA_DIR, 'uploads');
fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const upload = multer({ dest: UPLOAD_DIR });

// ---- Health/root (avoid "Cannot GET /") ----
app.get('/', (_req, res) => res.send('API OK. Try GET /api/staged or POST /api/scan'));
app.get('/health', (_req, res) => res.json({ ok: true }));

// ---- API ----
app.post('/api/scan', (req, res) => {
  const { serial, stage, operator, loadId, wagon1, wagon2, wagon3, timestamp } = req.body || {};
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

app.get('/api/staged', (_req, res) => {
  const rows = db.prepare('SELECT * FROM scans ORDER BY id DESC').all();
  res.json(rows);
});

// New DELETE endpoint for removing staged scans
app.delete('/api/remove-scan/:id', (req, res) => {
  const { id } = req.params;

  // Ensure ID is provided and is a valid number
  if (!id || isNaN(id)) {
    return res.status(400).json({ error: 'Invalid scan ID' });
  }

  try {
    const stmt = db.prepare('DELETE FROM scans WHERE id = ?');
    const result = stmt.run(id);

    // Check if a scan was deleted
    if (result.changes === 0) {
      return res.status(404).json({ error: 'Scan not found' });
    }

    res.json({ ok: true, message: 'Scan removed successfully' });
  } catch (e) {
    console.error('DB delete error', e);
    res.status(500).json({ error: 'Database error' });
  }
});

app.post('/api/staged/clear', (_req, res) => {
  db.prepare('DELETE FROM scans').run();
  res.json({ ok: true });
});

app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file.path });
});

app.post('/api/export-to-excel', async (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, 'template.xlsm');
    if (!fs.existsSync(templatePath)) {
      return res.status(400).json({ error: 'template.xlsm not found' });
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

// --- Start server (Render/Railway supply PORT) ---
const PORT = process.env.PORT || 4000;
app.listen(PORT, () => console.log(`✅ Backend on http://localhost:${PORT}`));
