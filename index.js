// index.js
import express from 'express';
import cors from 'cors';
import fs from 'fs';
import path from 'path';
import multer from 'multer';
import XLSX from 'xlsx';
import sqlite3pkg from 'sqlite3';

const app = express();
app.use(cors());
app.use(express.json());

const __dirname = process.cwd();
const DB_PATH = path.join(__dirname, 'rail_scans.db');
const db = new sqlite3pkg.Database(DB_PATH);

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

// Upload dir for Excel
const UPLOAD_DIR = path.join(__dirname,'uploads');
if(!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, {recursive:true});
const upload = multer({ dest: UPLOAD_DIR });

// Add new scan
app.post('/api/scan', (req,res)=>{
  const {serial, stage, operator, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, timestamp} = req.body;
  if(!serial) return res.status(400).json({error:'Serial required'});

  db.run(
    `INSERT INTO scans (serial, stage, operator, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, timestamp)
     VALUES (?,?,?,?,?,?,?,?,?,?,?)`,
    [serial, stage, operator, wagon1, wagon2, wagon3, grade, railType, spec, lengthM, timestamp],
    function(err){
      if(err) return res.status(500).json({error:err.message});
      res.json({ok:true, id:this.lastID});
    }
  );
});

// Get staged scans
app.get('/api/staged', (_req,res)=>{
  db.all('SELECT * FROM scans ORDER BY id DESC', (_err,rows)=>res.json(rows));
});

// Export Excel
app.post('/api/export-to-excel', (_req,res)=>{
  const templatePath = path.join(UPLOAD_DIR,'template.xlsm');
  if(!fs.existsSync(templatePath)) return res.status(400).json({error:'template.xlsm missing'});
  const wb = XLSX.readFile(templatePath,{cellDates:true,bookVBA:true});
  const sheetName = wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const existing = XLSX.utils.sheet_to_json(ws,{defval:''});

  db.all('SELECT * FROM scans', (_err,rows)=>{
    const appended = existing.concat(rows.map(s=>({
      Serial: s.serial,
      Wagon1: s.wagon1,
      Wagon2: s.wagon2,
      Wagon3: s.wagon3,
      Grade: s.grade,
      RailType: s.railType,
      Spec: s.spec,
      LengthM: s.lengthM,
      Timestamp: s.timestamp
    })));

    const newWs = XLSX.utils.json_to_sheet(appended,{skipHeader:false});
    wb.Sheets[sheetName] = newWs;
    const outName = `Master_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR,outName);
    XLSX.writeFile(wb,outPath,{bookType:'xlsm',bookVBA:true});
    res.download(outPath,outName);
  });
});

const PORT = process.env.PORT||4000;
app.listen(PORT,()=>console.log(`Backend on :${PORT}`));
