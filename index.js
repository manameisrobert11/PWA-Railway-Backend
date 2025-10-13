// backend/index.js
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

// --- DB setup ---
const DB_PATH = path.join(__dirname, 'rail_scans.db');
const db = new sqlite3pkg.Database(DB_PATH);

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS scans (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      serial TEXT,
      stage TEXT,
      operator TEXT,
      wagonId TEXT,
      grade TEXT,
      railType TEXT,
      spec TEXT,
      lengthM TEXT,
      timestamp TEXT
    )
  `);
});

// --- File storage setup for Excel ---
const UPLOAD_DIR = path.join(__dirname, 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const upload = multer({ dest: UPLOAD_DIR });

// --- API Routes ---

// Add new scan
app.post('/api/scan', (req, res) => {
  const { serial, stage, operator, wagonId, timestamp, grade, railType, spec, lengthM } = req.body;

  if (!serial) return res.status(400).json({ error: 'Serial required' });

  try {
    const stmt = db.prepare(`
      INSERT INTO scans (serial, stage, operator, wagonId, grade, railType, spec, lengthM, timestamp)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);
    const result = stmt.run(
      serial,
      stage || 'received',
      operator || 'unknown',
      wagonId || '',
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

// Get all scans
app.get('/api/staged', (req, res) => {
  db.all(`SELECT * FROM scans ORDER BY id DESC`, (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// Delete a specific staged scan by ID
app.delete('/api/staged/:id', (req, res) => {
  const scanId = req.params.id;

  if (!scanId) return res.status(400).json({ error: 'Scan ID is required' });

  try {
    const scanExists = db.prepare('SELECT * FROM scans WHERE id = ?').get(scanId);
    if (!scanExists) return res.status(404).json({ error: 'Scan not found' });

    const stmt = db.prepare('DELETE FROM scans WHERE id = ?');
    const result = stmt.run(scanId);

    if (result.changes === 0) return res.status(404).json({ error: 'Scan not found' });

    res.json({ ok: true, message: 'Scan removed successfully' });
  } catch (e) {
    console.error('Error removing scan', e);
    res.status(500).json({ error: 'Failed to remove scan' });
  }
});

// Clear all scans
app.post('/api/staged/clear', (_req, res) => {
  db.run(`DELETE FROM scans`, (err) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

// Upload Excel template
app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// Export to Excel (.xlsm)
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
        rows.map(s => ({
          Serial: s.serial,
          Stage: s.stage,
          Operator: s.operator,
          WagonID: s.wagonId,
          Grade: s.grade,
          RailType: s.railType,
          Spec: s.spec,
          Length: s.lengthM,
          Timestamp: s.timestamp
        }))
      );

      const newWs = XLSX.utils.json_to_sheet(appended, { skipHeader: false });
      wb.Sheets[sheetName] = newWs;

      const outName = `Master_${Date.now()}.xlsm`;
      const outPath = path.join(UPLOAD_DIR, outName);
      XLSX.writeFile(wb, outPath, { bookType: 'xlsm', bookVBA: true });

      res.download(outPath, outName, err => {
        if (err) console.error('Error sending file:', err);
      });
    });

  } catch (err) {
    console.error('Export failed:', err);
    res.status(500).json({ error: err.message });
  }
});

// Health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// Start server
const PORT = process.env.PORT || 4000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Backend on :${PORT}`);
});
