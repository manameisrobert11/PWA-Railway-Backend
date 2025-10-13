// backend/index.js
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import sqlite3 from 'sqlite3';

const app = express();
app.use(cors());
app.use(express.json());

const PORT = process.env.PORT || 4000;

// --- Database setup ---
const DB_PATH = path.join(process.cwd(), 'rail_scans.db');
const db = new sqlite3.Database(DB_PATH);

db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS scans (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      serial TEXT,
      stage TEXT,
      operator TEXT,
      wagon1 TEXT,
      wagon2 TEXT,
      wagon3 TEXT,
      grade TEXT,
      railType TEXT,
      spec TEXT,
      lengthM TEXT,
      timestamp TEXT
    )
  `);
});

// --- File uploads setup ---
const UPLOAD_DIR = path.join(process.cwd(), 'uploads');
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const upload = multer({ dest: UPLOAD_DIR });

// --- API routes ---

// Add new scan
app.post('/api/scan', (req, res) => {
  const { serial, stage, operator, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, timestamp } = req.body;
  const ts = timestamp || new Date().toISOString();

  const stmt = db.prepare(`
    INSERT INTO scans
    (serial, stage, operator, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, timestamp)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);
  stmt.run(serial, stage || 'received', operator || 'unknown', wagon1 || '', wagon2 || '', wagon3 || '', grade || '', railType || '', spec || '', lengthM || '', ts, function(err){
    if(err) return res.status(500).json({ error: err.message });
    res.json({ ok: true, id: this.lastID });
  });
});

// Get all staged scans
app.get('/api/staged', (_req, res) => {
  db.all(`SELECT * FROM scans ORDER BY id DESC`, (err, rows) => {
    if(err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// Delete a scan
app.delete('/api/staged/:id', (req, res) => {
  const { id } = req.params;
  db.run(`DELETE FROM scans WHERE id = ?`, [id], function(err){
    if(err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

// Clear all scans
app.post('/api/staged/clear', (_req, res) => {
  db.run(`DELETE FROM scans`, function(err){
    if(err) return res.status(500).json({ error: err.message });
    res.json({ ok: true });
  });
});

// Upload Excel template
app.post('/api/upload-template', upload.single('template'), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// Export staged scans to Excel (.xlsm)
app.post('/api/export-to-excel', (_req, res) => {
  const templatePath = path.join(UPLOAD_DIR, 'template.xlsm');
  if(!fs.existsSync(templatePath)) return res.status(400).json({ error: 'template.xlsm not found in uploads/' });

  const wb = XLSX.readFile(templatePath, { bookVBA: true, cellDates: true });
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const existing = XLSX.utils.sheet_to_json(ws, { defval: '' });

  db.all(`SELECT * FROM scans ORDER BY id ASC`, (err, rows) => {
    if(err) return res.status(500).json({ error: err.message });

    const appended = existing.concat(rows.map(s => ({
      Serial: s.serial,
      Stage: s.stage,
      Operator: s.operator,
      Wagon1: s.wagon1,
      Wagon2: s.wagon2,
      Wagon3: s.wagon3,
      Grade: s.grade,
      RailType: s.railType,
      Spec: s.spec,
      Length: s.lengthM,
      Timestamp: s.timestamp
    })));

    const newWs = XLSX.utils.json_to_sheet(appended, { skipHeader: false });
    wb.Sheets[sheetName] = newWs;

    const outName = `Master_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR, outName);
    XLSX.writeFile(wb, outPath, { bookType:'xlsm', bookVBA:true });

    res.download(outPath, outName);
  });
});

// Health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// Start server
app.listen(PORT, () => console.log(`Backend running on port ${PORT}`));
