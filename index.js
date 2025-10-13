// backend/index.js
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import multer from "multer";
import XLSX from "xlsx";
import { Server } from "socket.io";
import http from "http";
import sqlite3pkg from "sqlite3";

const __dirname = process.cwd();
const app = express();
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

app.use(cors());
app.use(express.json());

// --- SQLite setup ---
const sqlite3 = sqlite3pkg.verbose();
const DB_PATH = path.join(__dirname, "rail_scans.db");
const db = new sqlite3.Database(DB_PATH);

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
      raw TEXT,
      timestamp TEXT
    )
  `);
});

// --- Uploads directory ---
const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const upload = multer({ dest: UPLOAD_DIR });

// --- API Routes ---

// Add new scan
app.post("/api/scan", (req, res) => {
  const { serial, stage, operator, wagon1Id, wagon2Id, wagon3Id, grade, railType, spec, lengthM, raw, timestamp } = req.body;

  if (!serial) return res.status(400).json({ error: "Serial required" });

  const ts = timestamp || new Date().toISOString();
  const stmt = db.prepare(`
    INSERT INTO scans
    (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id, grade, railType, spec, lengthM, raw, timestamp)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
  `);

  stmt.run(
    serial,
    stage || "received",
    operator || "unknown",
    wagon1Id || "",
    wagon2Id || "",
    wagon3Id || "",
    grade || "",
    railType || "",
    spec || "",
    lengthM || "",
    raw || "",
    ts,
    function (err) {
      if (err) return res.status(500).json({ error: err.message });
      const inserted = {
        id: this.lastID,
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
        raw,
        timestamp: ts,
      };
      io.emit("new-scan", inserted); // Real-time broadcast
      res.json({ ok: true, id: this.lastID });
    }
  );
  stmt.finalize();
});

// Get all staged scans
app.get("/api/staged", (_req, res) => {
  db.all("SELECT * FROM scans ORDER BY id DESC", (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// Delete a scan by ID
app.delete("/api/remove-scan/:id", (req, res) => {
  const { id } = req.params;
  db.run("DELETE FROM scans WHERE id = ?", [id], function (err) {
    if (err) return res.status(500).json({ error: err.message });
    io.emit("deleted-scan", { id: Number(id) }); // Real-time broadcast
    res.json({ ok: true });
  });
});

// Clear all scans
app.post("/api/staged/clear", (_req, res) => {
  db.run("DELETE FROM scans", function (err) {
    if (err) return res.status(500).json({ error: err.message });
    io.emit("cleared-scans"); // Real-time broadcast
    res.json({ ok: true });
  });
});

// Upload Excel template
app.post("/api/upload-template", upload.single("template"), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// Export to Excel (.xlsm)
app.post("/api/export-to-excel", (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, "template.xlsm");
    if (!fs.existsSync(templatePath)) return res.status(400).json({ error: "template.xlsm not found" });

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const existing = XLSX.utils.sheet_to_json(ws, { defval: "" });

    db.all("SELECT * FROM scans ORDER BY id ASC", (err, rows) => {
      if (err) return res.status(500).json({ error: err.message });

      const appended = existing.concat(
        rows.map(s => ({
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
          Timestamp: s.timestamp,
        }))
      );

      const newWs = XLSX.utils.json_to_sheet(appended, { skipHeader: false });
      wb.Sheets[sheetName] = newWs;

      const outName = `Master_${Date.now()}.xlsm`;
      const outPath = path.join(UPLOAD_DIR, outName);
      XLSX.writeFile(wb, outPath, { bookType: "xlsm", bookVBA: true });

      res.download(outPath, outName);
    });
  } catch (err) {
    console.error("Export failed:", err);
    res.status(500).json({ error: err.message });
  }
});

// Health check
app.get("/api/health", (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// --- Start server ---
const PORT = process.env.PORT || 4000;
server.listen(PORT, () => console.log(`✅ Backend running on port ${PORT}`));
