// index.js — Backend root (Express + SQLite + Excel export + QR images + pagination + bulk ingest)
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import multer from "multer";
import XLSX from "xlsx";
import sqlite3pkg from "sqlite3";
import http from "http";
import { Server } from "socket.io";

// ---------- Lazy loaders (so the app runs even if deps aren't installed) ----------
async function getExcelJS() {
  try {
    const m = await import("exceljs");
    return m.default || m;
  } catch {
    return null;
  }
}
async function getQRCode() {
  try {
    const m = await import("qrcode");
    return m.default || m;
  } catch {
    return null;
  }
}

// ---------- App + paths ----------
const __dirname = process.cwd();
const app = express();
app.use(cors());
app.use(express.json());

// Simple request logger
app.use((req, _res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.originalUrl}`);
  next();
});

const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const QR_DIR = path.join(UPLOAD_DIR, "qrcodes");
if (!fs.existsSync(QR_DIR)) fs.mkdirSync(QR_DIR, { recursive: true });

const upload = multer({ dest: UPLOAD_DIR });

// ---------- SQLite ----------
const DB_PATH = path.join(__dirname, "rail_scans.db");
const db = new sqlite3pkg.Database(DB_PATH);

// Base schema (includes receivedAt/loadedAt + qr fields)
db.serialize(() => {
  db.run(`
    CREATE TABLE IF NOT EXISTS scans (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      serial     TEXT,
      stage      TEXT,
      operator   TEXT,
      wagon1Id   TEXT,
      wagon2Id   TEXT,
      wagon3Id   TEXT,
      receivedAt TEXT,
      loadedAt   TEXT,
      grade      TEXT,
      railType   TEXT,
      spec       TEXT,
      lengthM    TEXT,
      qrRaw      TEXT,   -- raw QR text captured
      qrPngPath  TEXT,   -- path to generated PNG (if created)
      timestamp  TEXT
    )
  `);
});

// Auto-migrate + harden SQLite + helpful indexes (idempotent)
function bootstrapDb() {
  db.all(`PRAGMA table_info(scans)`, (err, cols) => {
    if (err) return console.error("PRAGMA error:", err);
    const names = new Set(cols.map((c) => c.name));
    const alters = [];
    for (const col of ["receivedAt", "loadedAt", "qrRaw", "qrPngPath"]) {
      if (!names.has(col)) alters.push(`ALTER TABLE scans ADD COLUMN ${col} TEXT;`);
    }
    if (alters.length) {
      db.serialize(() => alters.forEach((sql) => db.run(sql)));
      console.log("✅ Added missing columns:", alters.join(" "));
    }
  });

  db.exec(`
    PRAGMA journal_mode = WAL;
    PRAGMA synchronous = NORMAL;
    PRAGMA busy_timeout = 5000;
    PRAGMA temp_store = MEMORY;
    PRAGMA cache_size = -16000; -- ~16MB cache
  `);

  // Indexes (safe to re-run)
  db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_scans_serial ON scans(serial)`);
  db.run(`CREATE INDEX IF NOT EXISTS ix_scans_timestamp ON scans(timestamp)`);
}
bootstrapDb();

// ---------- Prepared statements (reuse under load) ----------
const insertScanStmt = db.prepare(
  `INSERT OR IGNORE INTO scans
   (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id, receivedAt, loadedAt,
    grade, railType, spec, lengthM, qrRaw, qrPngPath, timestamp)
   VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`
);
process.on("exit", () => {
  try { insertScanStmt.finalize(); } catch {}
});

// ---------- HTTP + Socket.IO ----------
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

io.on("connection", (socket) => {
  console.log("Client connected:", socket.id);
  socket.on("disconnect", () => console.log("Client disconnected:", socket.id));
});

// ---------- API Routes ----------

// Version + health
app.get("/api/version", (_req, res) => {
  res.json({ ok: true, version: "export-xlsx-images-v2" });
});
app.get("/api/health", (_req, res) => {
  res.json({ ok: true, db: fs.existsSync(DB_PATH) });
});

// Add a new scan (online save; optimized for multi-device: immediate reply, PNG in background)
app.post("/api/scan", async (req, res) => {
  const {
    serial, stage, operator,
    wagon1Id, wagon2Id, wagon3Id,
    receivedAt, loadedAt,
    grade, railType, spec, lengthM,
    qrRaw, timestamp,
  } = req.body;

  if (!serial) return res.status(400).json({ error: "Serial required" });
  const ts = timestamp || new Date().toISOString();

  insertScanStmt.run(
    String(serial),
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
    qrRaw || "",
    "", // placeholder for png path
    ts,
    async function (err) {
      if (err) return res.status(500).json({ error: err.message });

      // If it was ignored (duplicate), lastID may be undefined/null
      const newId = this.lastID || null;

      // Emit immediately so other clients update fast
      io.emit("new-scan", {
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
        qrRaw: qrRaw || "",
        timestamp: ts,
      });

      // ✅ Respond to client immediately (do not block on PNG generation)
      res.json({ ok: true, id: newId });

      // 🧵 Background: generate QR PNG (best-effort)
      try {
        const QRCode = await getQRCode();
        if (QRCode && (qrRaw || serial) && newId) {
          const pngPath = path.join(QR_DIR, `${newId}.png`);
          const pngRel  = path.relative(__dirname, pngPath).replace(/\\/g, "/");
          await QRCode.toFile(pngPath, qrRaw || serial, { type: "png", margin: 1, scale: 4 });
          db.run(`UPDATE scans SET qrPngPath = ? WHERE id = ?`, [pngRel, newId]);
        }
      } catch (e) {
        console.warn("QR PNG generation failed:", e.message);
      }
    }
  );
});

// Bulk ingest (used by offline sync)
app.post("/api/scans/bulk", (req, res) => {
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (items.length === 0) return res.json({ ok: true, inserted: 0, skipped: 0 });

  db.serialize(() => {
    db.run("BEGIN TRANSACTION");
    const stmt = db.prepare(`
      INSERT OR IGNORE INTO scans
      (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id, receivedAt, loadedAt,
       grade, railType, spec, lengthM, qrRaw, qrPngPath, timestamp)
      VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    `);

    let inserted = 0, skipped = 0;
    for (const r of items) {
      stmt.run([
        String(r.serial || ""),
        r.stage || "received",
        r.operator || "unknown",
        r.wagon1Id || "",
        r.wagon2Id || "",
        r.wagon3Id || "",
        r.receivedAt || "",
        r.loadedAt || "",
        r.grade || "",
        r.railType || "",
        r.spec || "",
        r.lengthM || "",
        r.qrRaw || "",
        "", // qrPngPath optional
        r.timestamp || new Date().toISOString(),
      ], function(err){
        if (err) skipped++;
        else if (this.changes === 1) inserted++;
        else skipped++;
      });
    }

    stmt.finalize((finalErr) => {
      if (finalErr) {
        db.run("ROLLBACK");
        return res.status(500).json({ error: finalErr.message });
      }
      db.run("COMMIT", () => res.json({ ok: true, inserted, skipped }));
    });
  });
});

// Pagination + count (scalable)
app.get("/api/staged", (req, res) => {
  const limit = Math.min(parseInt(req.query.limit || "100", 10), 500);
  const cursor = req.query.cursor ? parseInt(req.query.cursor, 10) : null;
  const dir = (req.query.dir || "desc").toLowerCase() === "asc" ? "ASC" : "DESC";

  const where =
    cursor != null
      ? (dir === "DESC" ? "WHERE id < ?" : "WHERE id > ?")
      : "";

  const params = cursor != null ? [cursor] : [];

  const sql = `
    SELECT * FROM scans
    ${where}
    ORDER BY id ${dir}
    LIMIT ${limit}
  `;

  db.all(sql, params, (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });

    const nextCursor =
      rows.length > 0
        ? rows[rows.length - 1].id
        : null;

    db.get(`SELECT COUNT(*) AS c FROM scans`, (_, countRow) => {
      res.json({
        rows,
        nextCursor,
        total: countRow?.c ?? 0,
      });
    });
  });
});

app.get("/api/staged/count", (_req, res) => {
  db.get(`SELECT COUNT(*) AS c FROM scans`, (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    res.json({ count: row.c });
  });
});

// Delete a scan (emits to all clients)
app.delete("/api/staged/:id", (req, res) => {
  const scanId = req.params.id;
  if (!scanId) return res.status(400).json({ error: "Scan ID required" });

  db.get("SELECT * FROM scans WHERE id = ?", [scanId], (err, row) => {
    if (err) return res.status(500).json({ error: err.message });
    if (!row) return res.status(404).json({ error: "Scan not found" });

    // best-effort: delete QR png
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

// Clear all scans (emits)
app.post("/api/staged/clear", (_req, res) => {
  // clean qrcodes folder
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

// Upload Excel template (.xlsm)
app.post("/api/upload-template", upload.single("template"), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// Export to .xlsm (macro-enabled; fixed headers so columns always appear)
app.post("/api/export-to-excel", (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, "template.xlsm");
    if (!fs.existsSync(templatePath)) {
      return res.status(400).json({ error: "template.xlsm not found" });
    }

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0];

    const HEADERS = [
      "Serial","Stage","Operator",
      "Wagon1ID","Wagon2ID","Wagon3ID",
      "RecievedAt","LoadedAt",
      "Grade","RailType","Spec","Length",
      "QRText","QRImagePath",
      "Timestamp",
    ];

    db.all("SELECT * FROM scans ORDER BY id ASC", (err, rows) => {
      if (err) return res.status(500).json({ error: err.message });

      const dataRows = rows.map((s) => ([
        s.serial || "", s.stage || "", s.operator || "",
        s.wagon1Id || "", s.wagon2Id || "", s.wagon3Id || "",
        s.receivedAt || "", s.loadedAt || "",
        s.grade || "", s.railType || "", s.spec || "", s.lengthM || "",
        s.qrRaw || "", s.qrPngPath || "",
        s.timestamp || "",
      ]));

      const aoa = [HEADERS, ...dataRows];
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

// Export to .xlsx with embedded QR images (no template)
// Accepts GET or POST so Netlify/Render proxies are happy
app.all("/api/export-xlsx-images", async (_req, res) => {
  const ExcelJS = await getExcelJS();
  if (!ExcelJS) {
    return res.status(400).json({ error: "exceljs not installed. Run: npm i exceljs qrcode" });
  }
  const QRCode = await getQRCode();

  db.all("SELECT * FROM scans ORDER BY id ASC", async (err, rows) => {
    if (err) return res.status(500).json({ error: err.message });

    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Scans");

      const columns = [
        { header: "Serial",      key: "serial",      width: 22 },
        { header: "Stage",       key: "stage",       width: 12 },
        { header: "Operator",    key: "operator",    width: 18 },
        { header: "Wagon1ID",    key: "wagon1Id",    width: 14 },
        { header: "Wagon2ID",    key: "wagon2Id",    width: 14 },
        { header: "Wagon3ID",    key: "wagon3Id",    width: 14 },
        { header: "RecievedAt",  key: "receivedAt",  width: 18 },
        { header: "LoadedAt",    key: "loadedAt",    width: 18 },
        { header: "Grade",       key: "grade",       width: 12 },
        { header: "RailType",    key: "railType",    width: 12 },
        { header: "Spec",        key: "spec",        width: 18 },
        { header: "Length",      key: "lengthM",     width: 10 },
        { header: "QRText",      key: "qrRaw",       width: 42 },
        { header: "QR Image",    key: "qrImage",     width: 16 },
        { header: "Timestamp",   key: "timestamp",   width: 24 },
      ];
      ws.columns = columns;

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
          qrImage:    "",
          timestamp:  s.timestamp || "",
        });
      });
      ws.getRow(1).font = { bold: true };
      for (let i = 2; i <= rows.length + 1; i++) ws.getRow(i).height = 70;

      if (QRCode) {
        const qrImageColIndex = columns.findIndex((c) => c.key === "qrImage"); // 0-based
        const pixelSize = 90;
        for (let i = 0; i < rows.length; i++) {
          const text = rows[i].qrRaw || rows[i].serial || "";
          if (!text) continue;
          const buf = await QRCode.toBuffer(text, { type: "png", margin: 1, scale: 4 });
          const imgId = wb.addImage({ buffer: buf, extension: "png" });
          ws.addImage(imgId, {
            tl: { col: qrImageColIndex, row: i + 1 }, // 0-based
            ext: { width: pixelSize, height: pixelSize },
          });
        }
      }

      const outName = `Master_QR_${Date.now()}.xlsx`;
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
      await wb.xlsx.write(res);
      res.end();
    } catch (e) {
      console.error("Export (xlsx images) failed:", e);
      res.status(500).json({ error: e.message });
    }
  });
});

// ---------- Start ----------
const PORT = process.env.PORT || 4000;
server.listen(PORT, "0.0.0.0", () =>
  console.log(`✅ Backend + Socket.IO on :${PORT}`)
);
