// index.js — Express + MySQL + Excel export + QR images + pagination + bulk ingest
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import multer from "multer";
import XLSX from "xlsx";
import http from "http";
import { Server } from "socket.io";
import mysql from "mysql2/promise";

// ---------- Lazy loaders (optional dependencies) ----------
async function getExcelJS() {
  try { const m = await import("exceljs"); return m.default || m; } catch { return null; }
}
async function getQRCode() {
  try { const m = await import("qrcode"); return m.default || m; } catch { return null; }
}

// ---------- App + paths ----------
const __dirname = process.cwd();
const app = express();
app.use(cors());
app.use(express.json({ limit: "256kb" }));

const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const QR_DIR = path.join(UPLOAD_DIR, "qrcodes");
if (!fs.existsSync(QR_DIR)) fs.mkdirSync(QR_DIR, { recursive: true });
const upload = multer({ dest: UPLOAD_DIR });

// ---------- MySQL pool ----------
function cfgFromUrl(url) {
  const u = new URL(url);
  return {
    host: u.hostname,
    port: Number(u.port || 3306),
    user: decodeURIComponent(u.username),
    password: decodeURIComponent(u.password),
    database: u.pathname.replace(/^\//, ""),
    ssl: process.env.MYSQL_SSL === "true" ? { rejectUnauthorized: false } : undefined,
  };
}

const baseConfig = process.env.MYSQL_URL
  ? cfgFromUrl(process.env.MYSQL_URL)
  : {
      host: process.env.MYSQL_HOST || "localhost",
      port: Number(process.env.MYSQL_PORT || 3306),
      user: process.env.MYSQL_USER,
      password: process.env.MYSQL_PASSWORD,
      database: process.env.MYSQL_DATABASE,
      ssl: process.env.MYSQL_SSL === "true" ? { rejectUnauthorized: false } : undefined,
    };

export const pool = mysql.createPool({
  ...baseConfig,
  connectionLimit: 10,
  waitForConnections: true,
  queueLimit: 0,
});

// Bootstrap schema (idempotent)
async function bootstrapDb() {
  const conn = await pool.getConnection();
  try {
    await conn.query(`
      CREATE TABLE IF NOT EXISTS scans (
        id          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        serial      VARCHAR(191) UNIQUE,
        stage       VARCHAR(64),
        operator    VARCHAR(128),
        wagon1Id    VARCHAR(128),
        wagon2Id    VARCHAR(128),
        wagon3Id    VARCHAR(128),
        receivedAt  VARCHAR(64),
        loadedAt    VARCHAR(64),
        grade       VARCHAR(64),
        railType    VARCHAR(64),
        spec        VARCHAR(128),
        lengthM     VARCHAR(32),
        qrRaw       TEXT,
        qrPngPath   TEXT,
        timestamp   DATETIME
      );
    `);
    await conn.query(`CREATE INDEX IF NOT EXISTS ix_scans_timestamp ON scans (timestamp);`);
  } finally {
    conn.release();
  }
}
await bootstrapDb();

// ---------- HTTP + Socket.IO ----------
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

io.on("connection", () => { /* socket connected */ });

// ---------- Helpers ----------
async function generateQrPngAndPersist(id, text) {
  try {
    const QRCode = await getQRCode();
    if (!QRCode || !id || !text) return;
    const pngPath = path.join(QR_DIR, `${id}.png`);
    const pngRel  = path.relative(__dirname, pngPath).replace(/\\/g, "/");
    await QRCode.toFile(pngPath, text, { type: "png", margin: 1, scale: 4 });
    await pool.query(`UPDATE scans SET qrPngPath = ? WHERE id = ?`, [pngRel, id]);
  } catch (e) {
    console.warn("QR PNG generation failed:", e.message);
  }
}

// ---------- API Routes ----------

// Version + health
app.get("/api/version", (_req, res) => {
  res.json({ ok: true, version: "mysql-v1" });
});
app.get("/api/health", async (_req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ ok: true, db: true });
  } catch {
    res.status(500).json({ ok: false, db: false });
  }
});

// Add a new scan — MySQL upsert variant.
// Uses ON DUPLICATE KEY UPDATE as a no-op to preserve existing row.
app.post("/api/scan", async (req, res) => {
  const {
    serial, stage, operator,
    wagon1Id, wagon2Id, wagon3Id,
    receivedAt, loadedAt,
    grade, railType, spec, lengthM,
    qrRaw, timestamp,
  } = req.body;

  if (!serial) return res.status(400).json({ error: "Serial required" });

  const ts = timestamp ? new Date(timestamp) : new Date();
  try {
    const sql = `
      INSERT INTO scans
        (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id,
         receivedAt, loadedAt, grade, railType, spec, lengthM,
         qrRaw, qrPngPath, timestamp)
      VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE id = id
    `;
    const vals = [
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
      ts,
    ];
    const [result] = await pool.execute(sql, vals);

    // If inserted, we have insertId. If duplicate, fetch existing id by serial.
    let newId = result.insertId || null;
    if (!newId) {
      const [r2] = await pool.query(`SELECT id FROM scans WHERE serial = ? LIMIT 1`, [String(serial)]);
      newId = r2[0]?.id ?? null;
    }

    io.emit("new-scan", {
      id: newId,
      serial: String(serial),
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
      timestamp: ts.toISOString(),
    });

    res.json({ ok: true, id: newId });

    if (newId) generateQrPngAndPersist(newId, qrRaw || serial);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Bulk ingest — transactional, ignore duplicates by serial.
app.post("/api/scans/bulk", async (req, res) => {
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (items.length === 0) return res.json({ ok: true, inserted: 0, skipped: 0 });

  const conn = await pool.getConnection();
  let inserted = 0, skipped = 0;
  try {
    await conn.beginTransaction();

    const text = `
      INSERT INTO scans
        (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id,
         receivedAt, loadedAt, grade, railType, spec, lengthM,
         qrRaw, qrPngPath, timestamp)
      VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE id = id
    `;

    for (const r of items) {
      const vals = [
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
        r.timestamp ? new Date(r.timestamp) : new Date(),
      ];

      try {
        const [result] = await conn.execute(text, vals);
        let id = result.insertId || null;
        if (!id && r.serial) {
          const [r2] = await conn.query(`SELECT id FROM scans WHERE serial = ? LIMIT 1`, [String(r.serial)]);
          id = r2[0]?.id ?? null;
        }
        if (id) {
          inserted++;
          io.emit("new-scan", {
            id,
            serial: String(r.serial || ""),
            stage: r.stage || "received",
            operator: r.operator || "unknown",
            wagon1Id: r.wagon1Id || "",
            wagon2Id: r.wagon2Id || "",
            wagon3Id: r.wagon3Id || "",
            receivedAt: r.receivedAt || "",
            loadedAt: r.loadedAt || "",
            grade: r.grade || "",
            railType: r.railType || "",
            spec: r.spec || "",
            lengthM: r.lengthM || "",
            qrRaw: r.qrRaw || "",
            timestamp: (r.timestamp ? new Date(r.timestamp) : new Date()).toISOString(),
          });
          generateQrPngAndPersist(id, r.qrRaw || r.serial || "").catch(()=>{});
        } else {
          skipped++;
        }
      } catch {
        skipped++;
      }
    }

    await conn.commit();
    res.json({ ok: true, inserted, skipped });
  } catch (e) {
    await conn.rollback();
    console.error(e);
    res.status(500).json({ error: e.message });
  } finally {
    conn.release();
  }
});

// Pagination + count
app.get("/api/staged", async (req, res) => {
  const limit = Math.min(parseInt(req.query.limit || "100", 10), 500);
  const cursor = req.query.cursor ? parseInt(req.query.cursor, 10) : null;
  const dir = (req.query.dir || "desc").toLowerCase() === "asc" ? "ASC" : "DESC";

  try {
    const params = [];
    let where = "";
    if (cursor != null) {
      where = dir === "DESC" ? "WHERE id < ?" : "WHERE id > ?";
      params.push(cursor);
    }
    const [rows] = await pool.query(
      `SELECT * FROM scans ${where} ORDER BY id ${dir} LIMIT ${limit}`, params
    );

    const nextCursor = rows.length ? rows[rows.length - 1].id : null;
    const [[c]] = await pool.query(`SELECT COUNT(*) AS c FROM scans`);
    const total = Number(c.c || 0);

    res.json({ rows, nextCursor, total });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/staged/count", async (_req, res) => {
  try {
    const [[c]] = await pool.query(`SELECT COUNT(*) AS c FROM scans`);
    res.json({ count: Number(c.c || 0) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Delete a scan
app.delete("/api/staged/:id", async (req, res) => {
  const scanId = Number(req.params.id);
  if (!scanId) return res.status(400).json({ error: "Scan ID required" });
  try {
    const [rows] = await pool.query(`SELECT * FROM scans WHERE id = ?`, [scanId]);
    const row = rows[0];
    if (!row) return res.status(404).json({ error: "Scan not found" });

    if (row.qrPngPath) {
      const abs = path.join(__dirname, row.qrPngPath);
      fs.existsSync(abs) && fs.unlink(abs, () => {});
    }
    await pool.query(`DELETE FROM scans WHERE id = ?`, [scanId]);
    io.emit("deleted-scan", { id: scanId });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Clear all scans
app.post("/api/staged/clear", async (_req, res) => {
  try {
    if (fs.existsSync(QR_DIR)) {
      for (const f of fs.readdirSync(QR_DIR)) {
        const p = path.join(QR_DIR, f);
        try { fs.unlinkSync(p); } catch {}
      }
    }
    await pool.query(`DELETE FROM scans`);
    io.emit("cleared-scans");
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Upload Excel template (.xlsm)
app.post("/api/upload-template", upload.single("template"), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// Export to .xlsm (macro-enabled)
app.post("/api/export-to-excel", async (_req, res) => {
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

    const [rows] = await pool.query(`SELECT * FROM scans ORDER BY id ASC`);
    const dataRows = rows.map((s) => ([
      s.serial || "", s.stage || "", s.operator || "",
      s.wagon1Id || "", s.wagon2Id || "", s.wagon3Id || "",
      s.receivedAt || "", s.loadedAt || "",
      s.grade || "", s.railType || "", s.spec || "", s.lengthM || "",
      s.qrRaw || "", s.qrPngPath || "",
      s.timestamp ? new Date(s.timestamp).toISOString() : "",
    ]));

    const aoa = [HEADERS, ...dataRows];
    const newWs = XLSX.utils.aoa_to_sheet(aoa);
    wb.Sheets[sheetName] = newWs;

    const outName = `Master_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR, outName);
    XLSX.writeFile(wb, outPath, { bookType: "xlsm", bookVBA: true });
    res.download(outPath, outName);
  } catch (err) {
    console.error("Export failed:", err);
    res.status(500).json({ error: err.message });
  }
});

// Export to .xlsx with embedded QR images (no template)
app.all("/api/export-xlsx-images", async (_req, res) => {
  const ExcelJS = await getExcelJS();
  if (!ExcelJS) return res.status(400).json({ error: "exceljs not installed. Run: npm i exceljs qrcode" });
  const QRCode = await getQRCode();

  try {
    const [rows] = await pool.query(`SELECT * FROM scans ORDER BY id ASC`);

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
        timestamp:  s.timestamp ? new Date(s.timestamp).toISOString() : "",
      });
    });
    ws.getRow(1).font = { bold: true };
    for (let i = 2; i <= rows.length + 1; i++) ws.getRow(i).height = 70;

    if (QRCode) {
      const qrImageColIndex = columns.findIndex((c) => c.key === "qrImage");
      const pixelSize = 90;
      for (let i = 0; i < rows.length; i++) {
        const text = rows[i].qrRaw || rows[i].serial || "";
        if (!text) continue;
        const buf = await QRCode.toBuffer(text, { type: "png", margin: 1, scale: 4 });
        const imgId = wb.addImage({ buffer: buf, extension: "png" });
        ws.addImage(imgId, {
          tl: { col: qrImageColIndex, row: i + 1 },
          ext: { width: pixelSize, height: pixelSize },
        });
      }
    }

    const outName = `Master_QR_${Date.now()}.xlsx`;
    res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error("Export (xlsx images) failed:", e);
    res.status(500).json({ error: e.message });
  }
});

// ---------- Start ----------
const PORT = process.env.PORT || 4000;
server.listen(PORT, "0.0.0.0", () =>
  console.log(`✅ Backend + Socket.IO + MySQL on :${PORT}`)
);
