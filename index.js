// backend/index.js
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import multer from "multer";
import XLSX from "xlsx";
import sqlite3pkg from "sqlite3";
import http from "http";
import { Server } from "socket.io";
import ExcelJS from "exceljs";
import QRCode from "qrcode"; // <-- generate PNGs for QR codes

const __dirname = process.cwd();
const app = express();
app.use(cors());
app.use(express.json());

// --- SQLite setup ---
const DB_PATH = path.join(__dirname, "rail_scans.db");
const db = new sqlite3pkg.Database(DB_PATH);

// Base schema (includes receivedAt/loadedAt + qr columns)
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
      receivedAt TEXT,
      loadedAt TEXT,
      grade TEXT,
      railType TEXT,
      spec TEXT,
      lengthM TEXT,
      qrRaw TEXT,        -- raw QR text captured
      qrPngPath TEXT,    -- path to generated PNG
      timestamp TEXT
    )
  `);
});

// Auto-migrate older DBs to add missing columns
function ensureColumns() {
  db.all(`PRAGMA table_info(scans)`, (err, cols) => {
    if (err) return console.error("PRAGMA error:", err);
    const names = new Set(cols.map((c) => c.name));
    const alters = [];
    if (!names.has("receivedAt")) alters.push(`ALTER TABLE scans ADD COLUMN receivedAt TEXT;`);
    if (!names.has("loadedAt"))  alters.push(`ALTER TABLE scans ADD COLUMN loadedAt TEXT;`);
    if (!names.has("qrRaw"))     alters.push(`ALTER TABLE scans ADD COLUMN qrRaw TEXT;`);
    if (!names.has("qrPngPath")) alters.push(`ALTER TABLE scans ADD COLUMN qrPngPath TEXT;`);

    if (alters.length) {
      db.serialize(() => alters.forEach((sql) => db.run(sql)));
      console.log("✅ Added missing columns:", alters.join(" "));
    }
  });

  // Harden SQLite against lock stalls (helps avoid 504s)
  db.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA synchronous = NORMAL;
    PRAGMA busy_timeout = 5000;
  `);
}
ensureColumns();

// --- Upload directory ---
const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const QR_DIR = path.join(UPLOAD_DIR, "qrcodes");
if (!fs.existsSync(QR_DIR)) fs.mkdirSync(QR_DIR, { recursive: true });

const upload = multer({ dest: UPLOAD_DIR });

// --- HTTP + Socket.IO ---
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

// Socket.IO events
io.on("connection", (socket) => {
  console.log("Client connected:", socket.id);
  socket.on("disconnect", () => console.log("Client disconnected:", socket.id));
});

// --- API routes ---

// Add a new scan
app.post("/api/scan", (req, res) => {
  const {
    serial,
    stage,
    operator,
    wagon1Id,
    wagon2Id,
    wagon3Id,
    receivedAt,
    loadedAt,
    grade,
    railType,
    spec,
    lengthM,
    qrRaw,       // <-- from frontend: pending.raw (full QR payload string)
    timestamp,
  } = req.body;

  if (!serial) return res.status(400).json({ error: "Serial required" });

  const ts = timestamp || new Date().toISOString();

  const stmt = db.prepare(
    `INSERT INTO scans
    (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id, receivedAt, loadedAt, grade, railType, spec, lengthM, qrRaw, qrPngPath, timestamp)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
  );

  stmt.run(
    serial,
    stage || "received",
    operator || "unknown",
    wagon1Id || "",
    wagon2Id || "",
    wagon3Id || "",
    receivedAt || "",
    loadedAt || "",
    grade || "",
    railType || "",
    spec || "",
    lengthM || "",
    qrRaw || "",         // save raw QR text
    "",                  // qrPngPath placeholder for now
    ts,
    function (err) {
      if (err) return res.status(500).json({ error: err.message });

      const newId = this.lastID;
      const pngPath = path.join(QR_DIR, `${newId}.png`);
      const pngRel = path.relative(__dirname, pngPath).replace(/\\/g, "/"); // relative path for Excel

      // Generate PNG asynchronously, then update the row's path
      const qrText = qrRaw || serial; // fall back to serial if no raw provided
      QRCode.toFile(pngPath, qrText, { type: "png", margin: 1, scale: 4 }, (qrErr) => {
        if (qrErr) {
          console.error("QR generation failed for id", newId, qrErr);
        }
        db.run(`UPDATE scans SET qrPngPath = ? WHERE id = ?`, [pngRel, newId], (upErr) => {
          if (upErr) console.error("Failed to store qrPngPath:", upErr);
          const newScan = {
            id: newId,
            serial,
            stage: stage || "received",
            operator: operator || "unknown",
            wagon1Id: wagon1Id || "",
            wagon2Id: wagon2Id || "",
            wagon3Id: wagon3Id || "",
            receivedAt: receivedAt || "",
            loadedAt: loadedAt || "",
            grade: grade || "",
            railType: railType || "",
            spec: spec || "",
            lengthM: lengthM || "",
            qrRaw: qrText,
            qrPngPath: pngRel,
            timestamp: ts,
          };
          io.emit("new-scan", newScan);
          // Respond immediately after insert/update path
          res.json({ ok: true, id: newId });
        });
      });
    }
  );
});

// Get all scans
app.get("/api/staged", (_req, res) => {
  db.all(`SELECT * FROM scans ORDER BY id DESC`, (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json(rows);
  });
});

// Delete a scan
app.delete("/api/staged/:id", (req, res) => {
  const scanId = req.params.id;
  if (!scanId) return res.status(400).json({ error: "Scan ID required" });

  db.get("SELECT * FROM scans WHERE id = ?", [scanId], (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    if (!row) return res.status(404).json({ error: "Scan not found" });

    // Remove QR PNG if present
    if (row.qrPngPath) {
      const abs = path.join(__dirname, row.qrPngPath);
      fs.existsSync(abs) && fs.unlink(abs, () => {});
    }

    db.run("DELETE FROM scans WHERE id = ?", [scanId], function (err2) {
      if (err2) return res.status(500).json({ error: err2.message });
      io.emit("deleted-scan", { id: scanId });
      res.json({ ok: true });
    });
  });
});

// Clear all scans
app.post("/api/staged/clear", (_req, res) => {
  // Clean qrcodes folder (best-effort)
  if (fs.existsSync(QR_DIR)) {
    for (const f of fs.readdirSync(QR_DIR)) {
      const p = path.join(QR_DIR, f);
      try { fs.unlinkSync(p); } catch {}
    }
  }
  db.run("DELETE FROM scans", (err) => {
    if (err) return res.status(500).json({ error: err.message });
    io.emit("cleared-scans");
    res.json({ ok: true });
  });
});

// Export scans to Excel (.xlsm). Overwrite the first sheet with fixed headers
app.post("/api/export-to-excel", (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, "template.xlsm");
    if (!fs.existsSync(templatePath))
      return res.status(400).json({ error: "template.xlsm not found" });

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0];

    // Fixed header order we guarantee every time:
    const HEADERS = [
      "Serial",
      "Stage",
      "Operator",
      "Wagon1ID",
      "Wagon2ID",
      "Wagon3ID",
      "RecievedAt",
      "LoadedAt",
      "Grade",
      "RailType",
      "Spec",
      "Length",
      "QRText",
      "QRImagePath",
      "Timestamp",
    ];

    db.all("SELECT * FROM scans ORDER BY id ASC", (err, rows) => {
      if (err) return res.status(500).json({ error: err.message });

      // Build AOA data: first row = headers, the rest = values in the exact order
      const dataRows = rows.map((s) => ([
        s.serial || "",
        s.stage || "",
        s.operator || "",
        s.wagon1Id || "",
        s.wagon2Id || "",
        s.wagon3Id || "",
        s.receivedAt || "",
        s.loadedAt || "",
        s.grade || "",
        s.railType || "",
        s.spec || "",
        s.lengthM || "",
        s.qrRaw || "",       // raw qr text we saved
        s.qrPngPath || "",   // path to PNG (if you installed qrcode)
        s.timestamp || "",
      ]));

      const aoa = [HEADERS, ...dataRows];

      // Create a fresh sheet with our headers+rows and replace the first sheet
      const newWs = XLSX.utils.aoa_to_sheet(aoa);
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

// Export to .xlsx with embedded QR images (no macros)
app.post("/api/export-xlsx-images", (_req, res) => {
  db.all("SELECT * FROM scans ORDER BY id ASC", async (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });

    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Scans");

      // Fixed column layout
      const columns = [
        { header: "Serial",      key: "serial",      width: 22 },
        { header: "Stage",       key: "stage",       width: 12 },
        { header: "Operator",    key: "operator",    width: 18 },
        { header: "Wagon1ID",    key: "wagon1Id",    width: 14 },
        { header: "Wagon2ID",    key: "wagon2Id",    width: 14 },
        { header: "Wagon3ID",    key: "wagon3Id",    width: 14 },
        { header: "RecievedAt",  key: "receivedAt",  width: 18 }, // (label kept as you typed)
        { header: "LoadedAt",    key: "loadedAt",    width: 18 },
        { header: "Grade",       key: "grade",       width: 12 },
        { header: "RailType",    key: "railType",    width: 12 },
        { header: "Spec",        key: "spec",        width: 18 },
        { header: "Length",      key: "lengthM",     width: 10 },
        { header: "QRText",      key: "qrRaw",       width: 42 },
        { header: "QR Image",    key: "qrImage",     width: 16 },  // image column
        { header: "Timestamp",   key: "timestamp",   width: 24 },
      ];
      ws.columns = columns;

      // Push the rows (text fields only; we insert images after)
      rows.forEach((s) => {
        ws.addRow({
          serial:     s.serial || "",
          stage:      s.stage || "",
          operator:   s.operator || "",
          wagon1Id:   s.wagon1Id || "",
          wagon2Id:   s.wagon2Id || "",
          wagon3Id:   s.wagon3Id || "",
          receivedAt: s.receivedAt || "",
          loadedAt:   s.loadedAt || "",
          grade:      s.grade || "",
          railType:   s.railType || "",
          spec:       s.spec || "",
          lengthM:    s.lengthM || "",
          qrRaw:      s.qrRaw || s.serial || "",
          qrImage:    "", // placeholder cell for the image
          timestamp:  s.timestamp || "",
        });
      });

      // Make header bold
      ws.getRow(1).font = { bold: true };

      // Embed QR images in the "QR Image" column
      const qrImageColIndex = columns.findIndex(c => c.key === "qrImage"); // zero-based
      const pixelSize = 90; // image size in pixels
      // set row height for visibility (row 1 is header)
      for (let i = 2; i <= rows.length + 1; i++) {
        ws.getRow(i).height = 70;
      }

      for (let i = 0; i < rows.length; i++) {
        const s = rows[i];
        const text = s.qrRaw || s.serial || "";
        if (!text) continue;

        // Generate QR PNG as a buffer
        // (scale/margin can be tuned; this size fits a 16-width column + 70 row height nicely)
        const buf = await QRCode.toBuffer(text, { type: "png", margin: 1, scale: 4 });
        const imgId = wb.addImage({ buffer: buf, extension: "png" });

        // ExcelJS uses 0-based coordinates for adding images:
        // Row index in sheet is i+2 (account for header), column index is qrImageColIndex (0-based).
        ws.addImage(imgId, {
          tl:  { col: qrImageColIndex, row: i + 1 }, // top-left (row-1 because ExcelJS rows are zero-based here)
          ext: { width: pixelSize, height: pixelSize },
        });
      }

      // Stream response
      const outName = `Master_QR_${Date.now()}.xlsx`;
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      console.error("Export (xlsx images) failed:", e);
      res.status(500).json({ error: e.message });
    }
  });
});


// Health check
app.get("/api/health", (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// --- Start server ---
const PORT = process.env.PORT || 4000;
server.listen(PORT, "0.0.0.0", () => console.log(`✅ Backend + Socket.IO on :${PORT}`));


