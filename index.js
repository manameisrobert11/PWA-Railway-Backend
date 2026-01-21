// index.js — Express + MySQL + Excel export + QR images + pagination + bulk ingest + Socket.IO
// Main + ALT (separate table) — FIXED VERSION

import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import multer from "multer";
import XLSX from "xlsx";
import http from "http";
import { Server } from "socket.io";
import mysql from "mysql2/promise";

// ----- BOOT TAG (change this string each deploy to prove new code is running)
console.log("BOOT TAG:", "2025-01-21-rail-v3-fixed");

// ---------- Lazy loaders (optional dependencies) ----------
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

// ---- Allowed origins (Netlify + local dev) ----
// UPDATE THIS with your actual Netlify URL!
const ALLOWED_ORIGINS = [
  "https://pwarailway.netlify.app",
  "https://railinventoryscanner.netlify.app",  // Add your actual Netlify URL if different
  process.env.FRONTEND_URL,  // Allow setting via environment variable
  process.env.LOCAL_ORIGIN || "http://localhost:5173",
].filter(Boolean);  // Remove undefined/null values

console.log("[CORS] Allowed origins:", ALLOWED_ORIGINS);

// ---- CORS for REST ----
app.use(cors({
  origin: (origin, cb) => {
    // allow non-browser tools (no origin) and whitelisted origins
    if (!origin || ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
    console.warn(`[CORS] Blocked origin: ${origin}`);
    return cb(new Error(`CORS blocked for origin: ${origin}`));
  },
  methods: ["GET", "POST", "DELETE", "PUT", "OPTIONS"],
  credentials: true,
}));

app.use(express.json({ limit: "256kb" }));

const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });

const QR_DIR = path.join(UPLOAD_DIR, "qrcodes");
if (!fs.existsSync(QR_DIR)) fs.mkdirSync(QR_DIR, { recursive: true });

// ALT QR directory (separate from main)
const QR_DIR_ALT = path.join(UPLOAD_DIR, "qrcodes_alt");
if (!fs.existsSync(QR_DIR_ALT)) fs.mkdirSync(QR_DIR_ALT, { recursive: true });

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

// ----- DEBUG: print env + the DB config we will actually use
console.log("[ENV SNAPSHOT]", {
  MYSQL_URL: process.env.MYSQL_URL ? "***set***" : null,
  MYSQL_HOST: process.env.MYSQL_HOST || null,
  MYSQL_PORT: process.env.MYSQL_PORT || null,
  MYSQL_DATABASE: process.env.MYSQL_DATABASE || null,
  MYSQL_USER: process.env.MYSQL_USER,
  MYSQL_SSL: process.env.MYSQL_SSL || null,
  NODE_VERSION: process.env.NODE_VERSION || null,
  FRONTEND_URL: process.env.FRONTEND_URL || null,
});

console.log("[DB CONFIG about to use]", {
  host: baseConfig.host,
  port: baseConfig.port,
  user: baseConfig.user,
  db: baseConfig.database,
  ssl: !!baseConfig.ssl,
});

export const pool = mysql.createPool({
  ...baseConfig,
  connectionLimit: 30,
  waitForConnections: true,
  queueLimit: 0,
});

// ---------- Bootstrap schema (idempotent) ----------
async function bootstrapDb() {
  const conn = await pool.getConnection();
  try {
    // MAIN table
    await conn.query(`
      CREATE TABLE IF NOT EXISTS scans (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        serial VARCHAR(191) UNIQUE,
        stage VARCHAR(64),
        operator VARCHAR(128),
        wagon1Id VARCHAR(128),
        wagon2Id VARCHAR(128),
        wagon3Id VARCHAR(128),
        receivedAt VARCHAR(64),
        loadedAt VARCHAR(64),
        destination VARCHAR(128),
        grade VARCHAR(64),
        railType VARCHAR(64),
        spec VARCHAR(128),
        lengthM VARCHAR(32),
        qrRaw TEXT,
        qrPngPath TEXT,
        timestamp DATETIME,
        INDEX ix_scans_timestamp (timestamp)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    `);

    // ALT table
    await conn.query(`
      CREATE TABLE IF NOT EXISTS scans_alt (
        id BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        serial VARCHAR(191) UNIQUE,
        stage VARCHAR(64),
        operator VARCHAR(128),
        wagon1Id VARCHAR(128),
        wagon2Id VARCHAR(128),
        wagon3Id VARCHAR(128),
        receivedAt VARCHAR(64),
        loadedAt VARCHAR(64),
        destination VARCHAR(128),
        grade VARCHAR(64),
        railType VARCHAR(64),
        spec VARCHAR(128),
        lengthM VARCHAR(32),
        qrRaw TEXT,
        qrPngPath TEXT,
        timestamp DATETIME,
        INDEX ix_scans_alt_timestamp (timestamp)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    `);
  } finally {
    conn.release();
  }
}

// wait for DB reachable with retries
async function waitForDb(retries = 12) {
  for (let i = 1; i <= retries; i++) {
    try {
      await pool.query("SELECT 1");
      console.log("✅ DB connected");
      return;
    } catch (e) {
      console.log(`DB not ready (attempt ${i}/${retries})`, e.code || e.message);
      await new Promise((r) => setTimeout(r, i * 1000));
    }
  }
  throw new Error("DB not reachable after retries");
}

await waitForDb();
await bootstrapDb();

// ---------- HTTP + Socket.IO ----------
const server = http.createServer(app);

const io = new Server(server, {
  path: "/socket.io",
  transports: ["websocket", "polling"],
  cors: {
    origin: (origin, cb) => {
      if (!origin || ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
      console.warn(`[Socket CORS] Blocked origin: ${origin}`);
      return cb(new Error(`Socket CORS blocked for origin: ${origin}`));
    },
    methods: ["GET", "POST"],
    credentials: true,
  },
});

// Simple connection logging
io.on("connection", (socket) => {
  console.log("socket connected:", socket.id, "from", socket.handshake.headers.origin || "unknown");
  socket.on("disconnect", (reason) => {
    console.log("socket disconnected:", socket.id, reason);
  });
});

// ---------- Helpers ----------
async function generateQrPngAndPersist(table, id, text) {
  try {
    const QRCode = await getQRCode();
    if (!QRCode || !id || !text) return;
    const dir = table === "scans_alt" ? QR_DIR_ALT : QR_DIR;
    const pngPath = path.join(dir, `${id}.png`);
    const pngRel = path.relative(__dirname, pngPath).replace(/\\/g, "/");
    await QRCode.toFile(pngPath, text, { type: "png", margin: 1, scale: 4 });
    await pool.query(`UPDATE ${table} SET qrPngPath = ? WHERE id = ?`, [pngRel, id]);
  } catch (e) {
    console.warn("QR PNG generation failed:", e.message);
  }
}

// ---------- API Routes (common misc) ----------
app.get("/", (_req, res) => res.send("Rail backend is running."));

app.get("/api/version", (_req, res) => {
  res.json({ ok: true, version: "mysql-v3", bootTag: "2025-01-21-rail-v3-fixed" });
});

app.get("/api/health", async (_req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ ok: true, db: true });
  } catch {
    res.status(500).json({ ok: false, db: false });
  }
});

app.get("/socket-test", (_req, res) => {
  res.type("text/plain").send("OK");
});

// ---------- MAIN PIPELINE ----------

// Add a new scan — UPSERT by serial
app.post("/api/scan", async (req, res) => {
  const {
    serial, stage, operator,
    wagon1Id, wagon2Id, wagon3Id,
    receivedAt, loadedAt, destination,
    grade, railType, spec, lengthM,
    qrRaw, timestamp,
  } = req.body;

  if (!serial) return res.status(400).json({ error: "Serial required" });

  const ts = timestamp ? new Date(timestamp) : new Date();

  try {
    const sql = `
      INSERT INTO scans (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id,
        receivedAt, loadedAt, destination, grade, railType, spec, lengthM, qrRaw, qrPngPath, timestamp)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE
        stage = VALUES(stage),
        operator = VALUES(operator),
        wagon1Id = VALUES(wagon1Id),
        wagon2Id = VALUES(wagon2Id),
        wagon3Id = VALUES(wagon3Id),
        receivedAt = VALUES(receivedAt),
        loadedAt = VALUES(loadedAt),
        destination = VALUES(destination),
        grade = VALUES(grade),
        railType = VALUES(railType),
        spec = VALUES(spec),
        lengthM = VALUES(lengthM),
        qrRaw = VALUES(qrRaw),
        timestamp = VALUES(timestamp)
    `;

    const vals = [
      String(serial), stage || "received", operator || "unknown",
      wagon1Id || "", wagon2Id || "", wagon3Id || "",
      receivedAt || "", loadedAt || "", destination || "",
      grade || "", railType || "", spec || "", lengthM || "",
      qrRaw || "", ts,
    ];

    const [result] = await pool.execute(sql, vals);
    let newId = result.insertId || null;

    if (!newId) {
      const [r2] = await pool.query(`SELECT id FROM scans WHERE serial = ? LIMIT 1`, [String(serial)]);
      newId = r2[0]?.id ?? null;
    }

    const payload = {
      id: newId,
      serial: String(serial),
      stage: stage || "received",
      operator: operator || "unknown",
      wagon1Id: wagon1Id || "",
      wagon2Id: wagon2Id || "",
      wagon3Id: wagon3Id || "",
      receivedAt: receivedAt || "",
      loadedAt: loadedAt || "",
      destination: destination || "",
      grade: grade || "",
      railType: railType || "",
      spec: spec || "",
      lengthM: lengthM || "",
      qrRaw: qrRaw || "",
      timestamp: ts.toISOString(),
    };

    io.emit("new-scan", payload);
    res.json({ ok: true, id: newId });

    if (newId) generateQrPngAndPersist("scans", newId, qrRaw || serial);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Bulk ingest — MAIN
app.post("/api/scans/bulk", async (req, res) => {
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (items.length === 0) return res.json({ ok: true, inserted: 0, skipped: 0 });

  const conn = await pool.getConnection();
  let inserted = 0, skipped = 0;

  try {
    await conn.beginTransaction();

    const sql = `
      INSERT INTO scans (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id,
        receivedAt, loadedAt, destination, grade, railType, spec, lengthM, qrRaw, qrPngPath, timestamp)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE
        stage = VALUES(stage),
        operator = VALUES(operator),
        wagon1Id = VALUES(wagon1Id),
        wagon2Id = VALUES(wagon2Id),
        wagon3Id = VALUES(wagon3Id),
        receivedAt = VALUES(receivedAt),
        loadedAt = VALUES(loadedAt),
        destination = VALUES(destination),
        grade = VALUES(grade),
        railType = VALUES(railType),
        spec = VALUES(spec),
        lengthM = VALUES(lengthM),
        qrRaw = VALUES(qrRaw),
        timestamp = VALUES(timestamp)
    `;

    for (const r of items) {
      const vals = [
        String(r.serial || ""), r.stage || "received", r.operator || "unknown",
        r.wagon1Id || "", r.wagon2Id || "", r.wagon3Id || "",
        r.receivedAt || "", r.loadedAt || "", r.destination || "",
        r.grade || "", r.railType || "", r.spec || "", r.lengthM || "",
        r.qrRaw || "", r.timestamp ? new Date(r.timestamp) : new Date(),
      ];

      try {
        const [result] = await conn.execute(sql, vals);
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
            destination: r.destination || "",
            grade: r.grade || "",
            railType: r.railType || "",
            spec: r.spec || "",
            lengthM: r.lengthM || "",
            qrRaw: r.qrRaw || "",
            timestamp: (r.timestamp ? new Date(r.timestamp) : new Date()).toISOString(),
          });
          generateQrPngAndPersist("scans", id, r.qrRaw || r.serial || "").catch(() => {});
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

// Pagination + count — MAIN
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
      `SELECT * FROM scans ${where} ORDER BY id ${dir} LIMIT ${limit}`,
      params
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

// Exists (for fast duplicate checks) — MAIN
app.get("/api/exists/:serial", async (req, res) => {
  const serial = String(req.params.serial || "").trim();
  if (!serial) return res.status(400).json({ error: "Serial required" });

  try {
    const [rows] = await pool.query(`SELECT * FROM scans WHERE serial = ? LIMIT 1`, [serial]);
    if (rows.length) return res.json({ exists: true, row: rows[0] });
    return res.json({ exists: false });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Delete a scan — MAIN
app.delete("/api/staged/:id", async (req, res) => {
  const scanId = Number(req.params.id);
  if (!scanId) return res.status(400).json({ error: "Scan ID required" });

  try {
    const [rows] = await pool.query(`SELECT * FROM scans WHERE id = ?`, [scanId]);
    const row = rows[0];
    if (!row) return res.status(404).json({ error: "Scan not found" });

    if (row.qrPngPath) {
      const abs = path.join(__dirname, row.qrPngPath);
      if (fs.existsSync(abs)) fs.unlink(abs, () => {});
    }

    await pool.query(`DELETE FROM scans WHERE id = ?`, [scanId]);
    io.emit("deleted-scan", { id: scanId });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Clear all scans — MAIN
app.post("/api/staged/clear", async (req, res) => {
  try {
    console.log("[CLEAR] /api/staged/clear called - ip:", req.ip, "ua:", req.headers["user-agent"]);

    if (fs.existsSync(QR_DIR)) {
      for (const f of fs.readdirSync(QR_DIR)) {
        const p = path.join(QR_DIR, f);
        try {
          fs.unlinkSync(p);
        } catch (err) {
          console.warn("Failed to unlink QR file:", p, err && err.message);
        }
      }
    }

    await pool.query(`DELETE FROM scans`);
    io.emit("cleared-scans");
    res.json({ ok: true });
  } catch (e) {
    console.error("[CLEAR] error:", e);
    res.status(500).json({ error: e.message });
  }
});

// Upload Excel template
app.post("/api/upload-template", upload.single("template"), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// ---- Helper: programmatic XLSX export fallback ----
async function programmaticXlsmFallback({ table = "scans", outNamePrefix = "Master" }, res) {
  const ExcelJS = await getExcelJS();
  const QRCode = await getQRCode();

  const [rows] = await pool.query(`SELECT * FROM ${table} ORDER BY id ASC`);

  if (ExcelJS) {
    try {
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet("Scans");

      const columns = [
        { header: "Serial", key: "serial", width: 22 },
        { header: "Stage", key: "stage", width: 12 },
        { header: "Operator", key: "operator", width: 18 },
        { header: "Wagon1ID", key: "wagon1Id", width: 14 },
        { header: "Wagon2ID", key: "wagon2Id", width: 14 },
        { header: "Wagon3ID", key: "wagon3Id", width: 14 },
        { header: "RecievedAt", key: "receivedAt", width: 18 },
        { header: "LoadedAt", key: "loadedAt", width: 18 },
        { header: "Destination", key: "destination", width: 20 },
        { header: "Grade", key: "grade", width: 12 },
        { header: "RailType", key: "railType", width: 12 },
        { header: "Spec", key: "spec", width: 18 },
        { header: "Length", key: "lengthM", width: 10 },
        { header: "QRText", key: "qrRaw", width: 42 },
        { header: "QR Image", key: "qrImage", width: 16 },
        { header: "Timestamp", key: "timestamp", width: 24 },
      ];

      ws.columns = columns;

      rows.forEach((s) => {
        ws.addRow({
          serial: s.serial || "",
          stage: s.stage || "",
          operator: s.operator || "",
          wagon1Id: s.wagon1Id || "",
          wagon2Id: s.wagon2Id || "",
          wagon3Id: s.wagon3Id || "",
          receivedAt: s.receivedAt || "",
          loadedAt: s.loadedAt || "",
          destination: s.destination || "",
          grade: s.grade || "",
          railType: s.railType || "",
          spec: s.spec || "",
          lengthM: s.lengthM || "",
          qrRaw: s.qrRaw || s.serial || "",
          qrImage: "",
          timestamp: s.timestamp ? new Date(s.timestamp).toISOString() : "",
        });
      });

      ws.getRow(1).font = { bold: true };
      for (let i = 2; i <= rows.length + 1; i++) ws.getRow(i).height = 70;

      if (QRCode) {
        const colToLetter = (index) => {
          let num = index + 1;
          let s = "";
          while (num > 0) {
            const m = (num - 1) % 26;
            s = String.fromCharCode(65 + m) + s;
            num = Math.floor((num - 1) / 26);
          }
          return s;
        };

        const qrImageColIndex = columns.findIndex((c) => c.key === "qrImage");
        const colLetter = colToLetter(qrImageColIndex);
        let imagesAdded = 0;

        for (let i = 0; i < rows.length; i++) {
          const text = rows[i].qrRaw || rows[i].serial || "";
          if (!text) continue;

          const buf = await QRCode.toBuffer(text, { type: "png", margin: 1, scale: 4 });
          const imgId = wb.addImage({ buffer: buf, extension: "png" });
          const rowNumber = i + 2;
          const range = `${colLetter}${rowNumber}:${colLetter}${rowNumber}`;
          ws.addImage(imgId, range);
          imagesAdded++;
        }
        console.log(`[FALLBACK-${table}] images added:`, imagesAdded);
      }

      const outName = `${outNamePrefix}_fallback_${Date.now()}.xlsx`;
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
      await wb.xlsx.write(res);
      res.end();
      return;
    } catch (e) {
      console.warn(`[FALLBACK-${table}] ExcelJS fallback failed:`, e.message || e);
    }
  }

  // Last-resort fallback: use XLSX (no images)
  try {
    const HEADERS = [
      "Serial", "Stage", "Operator",
      "Wagon1ID", "Wagon2ID", "Wagon3ID",
      "RecievedAt", "LoadedAt", "Destination",
      "Grade", "RailType", "Spec", "Length",
      "QRText", "QRImagePath", "Timestamp",
    ];

    const dataRows = rows.map((s) => ([
      s.serial || "", s.stage || "", s.operator || "",
      s.wagon1Id || "", s.wagon2Id || "", s.wagon3Id || "",
      s.receivedAt || "", s.loadedAt || "", s.destination || "",
      s.grade || "", s.railType || "", s.spec || "", s.lengthM || "",
      s.qrRaw || "", s.qrPngPath || "",
      s.timestamp ? new Date(s.timestamp).toISOString() : "",
    ]));

    const aoa = [HEADERS, ...dataRows];
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, "Scans");

    const outName = `${outNamePrefix}_fallback_${Date.now()}.xlsx`;
    const buf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    res.send(Buffer.from(buf));
  } catch (e) {
    console.error(`[FALLBACK-${table}] final fallback failed:`, e);
    res.status(500).json({ error: e.message || String(e) });
  }
}

// Export to .xlsm — MAIN
app.post("/api/export-to-excel", async (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, "template.xlsm");

    if (!fs.existsSync(templatePath)) {
      console.warn("[EXPORT-XLSM] template.xlsm not found — falling back to programmatic export");
      return await programmaticXlsmFallback({ table: "scans", outNamePrefix: "Master" }, res);
    }

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0] || "Sheet1";

    const HEADERS = [
      "Serial", "Stage", "Operator",
      "Wagon1ID", "Wagon2ID", "Wagon3ID",
      "RecievedAt", "LoadedAt", "Destination",
      "Grade", "RailType", "Spec", "Length",
      "QRText", "QRImagePath", "Timestamp",
    ];

    const [rows] = await pool.query(`SELECT * FROM scans ORDER BY id ASC`);

    const dataRows = rows.map((s) => ([
      s.serial || "", s.stage || "", s.operator || "",
      s.wagon1Id || "", s.wagon2Id || "", s.wagon3Id || "",
      s.receivedAt || "", s.loadedAt || "", s.destination || "",
      s.grade || "", s.railType || "", s.spec || "", s.lengthM || "",
      s.qrRaw || "", s.qrPngPath || "",
      s.timestamp ? new Date(s.timestamp).toISOString() : "",
    ]));

    const aoa = [HEADERS, ...dataRows];
    const newWs = XLSX.utils.aoa_to_sheet(aoa);
    wb.Sheets[sheetName] = newWs;

    if (!wb.SheetNames.includes(sheetName)) wb.SheetNames = [sheetName, ...wb.SheetNames];

    const outBuffer = XLSX.write(wb, { bookType: "xlsm", bookVBA: true, type: "buffer" });
    const outName = `Master_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR, outName);
    fs.writeFileSync(outPath, outBuffer);

    res.download(outPath, outName, (err) => {
      if (err) console.error("[EXPORT-XLSM] download error:", err);
    });
  } catch (err) {
    console.error("Export failed:", err);
    res.status(500).json({ error: err.message });
  }
});

// Export to .xlsx with embedded QR images — MAIN
app.all("/api/export-xlsx-images", async (_req, res) => {
  const ExcelJS = await getExcelJS();
  if (!ExcelJS) return res.status(400).json({ error: "exceljs not installed" });

  const QRCode = await getQRCode();

  try {
    const [rows] = await pool.query(`SELECT * FROM scans ORDER BY id ASC`);
    console.log("[EXPORT-XLSX-IMG] rows:", rows.length);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Scans");

    const columns = [
      { header: "Serial", key: "serial", width: 22 },
      { header: "Stage", key: "stage", width: 12 },
      { header: "Operator", key: "operator", width: 18 },
      { header: "Wagon1ID", key: "wagon1Id", width: 14 },
      { header: "Wagon2ID", key: "wagon2Id", width: 14 },
      { header: "Wagon3ID", key: "wagon3Id", width: 14 },
      { header: "RecievedAt", key: "receivedAt", width: 18 },
      { header: "LoadedAt", key: "loadedAt", width: 18 },
      { header: "Destination", key: "destination", width: 20 },
      { header: "Grade", key: "grade", width: 12 },
      { header: "RailType", key: "railType", width: 12 },
      { header: "Spec", key: "spec", width: 18 },
      { header: "Length", key: "lengthM", width: 10 },
      { header: "QRText", key: "qrRaw", width: 42 },
      { header: "QR Image", key: "qrImage", width: 16 },
      { header: "Timestamp", key: "timestamp", width: 24 },
    ];

    ws.columns = columns;

    rows.forEach((s) => {
      ws.addRow({
        serial: s.serial || "",
        stage: s.stage || "",
        operator: s.operator || "",
        wagon1Id: s.wagon1Id || "",
        wagon2Id: s.wagon2Id || "",
        wagon3Id: s.wagon3Id || "",
        receivedAt: s.receivedAt || "",
        loadedAt: s.loadedAt || "",
        destination: s.destination || "",
        grade: s.grade || "",
        railType: s.railType || "",
        spec: s.spec || "",
        lengthM: s.lengthM || "",
        qrRaw: s.qrRaw || s.serial || "",
        qrImage: "",
        timestamp: s.timestamp ? new Date(s.timestamp).toISOString() : "",
      });
    });

    ws.getRow(1).font = { bold: true };
    for (let i = 2; i <= rows.length + 1; i++) ws.getRow(i).height = 70;

    if (QRCode) {
      const colToLetter = (index) => {
        let num = index + 1;
        let s = "";
        while (num > 0) {
          const m = (num - 1) % 26;
          s = String.fromCharCode(65 + m) + s;
          num = Math.floor((num - 1) / 26);
        }
        return s;
      };

      const qrImageColIndex = columns.findIndex((c) => c.key === "qrImage");
      const colLetter = colToLetter(qrImageColIndex);
      let imagesAdded = 0;

      for (let i = 0; i < rows.length; i++) {
        const text = rows[i].qrRaw || rows[i].serial || "";
        if (!text) continue;

        const buf = await QRCode.toBuffer(text, { type: "png", margin: 1, scale: 4 });
        const imgId = wb.addImage({ buffer: buf, extension: "png" });
        const rowNumber = i + 2;
        const range = `${colLetter}${rowNumber}:${colLetter}${rowNumber}`;
        ws.addImage(imgId, range);
        imagesAdded++;
      }
      console.log("[EXPORT-XLSX-IMG] images added:", imagesAdded);
    }

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

// ---------- ALT PIPELINE ----------

// Add a new scan — ALT
app.post("/api/scan-alt", async (req, res) => {
  const {
    serial, stage, operator,
    wagon1Id, wagon2Id, wagon3Id,
    receivedAt, loadedAt, destination,
    grade, railType, spec, lengthM,
    qrRaw, timestamp,
  } = req.body;

  if (!serial) return res.status(400).json({ error: "Serial required" });

  const ts = timestamp ? new Date(timestamp) : new Date();

  try {
    const sql = `
      INSERT INTO scans_alt (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id,
        receivedAt, loadedAt, destination, grade, railType, spec, lengthM, qrRaw, qrPngPath, timestamp)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE
        stage = VALUES(stage),
        operator = VALUES(operator),
        wagon1Id = VALUES(wagon1Id),
        wagon2Id = VALUES(wagon2Id),
        wagon3Id = VALUES(wagon3Id),
        receivedAt = VALUES(receivedAt),
        loadedAt = VALUES(loadedAt),
        destination = VALUES(destination),
        grade = VALUES(grade),
        railType = VALUES(railType),
        spec = VALUES(spec),
        lengthM = VALUES(lengthM),
        qrRaw = VALUES(qrRaw),
        timestamp = VALUES(timestamp)
    `;

    const vals = [
      String(serial), stage || "received", operator || "unknown",
      wagon1Id || "", wagon2Id || "", wagon3Id || "",
      receivedAt || "", loadedAt || "", destination || "",
      grade || "", railType || "", spec || "", lengthM || "",
      qrRaw || "", ts,
    ];

    const [result] = await pool.execute(sql, vals);
    let newId = result.insertId || null;

    if (!newId) {
      const [r2] = await pool.query(`SELECT id FROM scans_alt WHERE serial = ? LIMIT 1`, [String(serial)]);
      newId = r2[0]?.id ?? null;
    }

    const payload = {
      id: newId,
      serial: String(serial),
      stage: stage || "received",
      operator: operator || "unknown",
      wagon1Id: wagon1Id || "",
      wagon2Id: wagon2Id || "",
      wagon3Id: wagon3Id || "",
      receivedAt: receivedAt || "",
      loadedAt: loadedAt || "",
      destination: destination || "",
      grade: grade || "",
      railType: railType || "",
      spec: spec || "",
      lengthM: lengthM || "",
      qrRaw: qrRaw || "",
      timestamp: ts.toISOString(),
    };

    io.emit("new-scan-alt", payload);
    res.json({ ok: true, id: newId });

    if (newId) generateQrPngAndPersist("scans_alt", newId, qrRaw || serial);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Bulk ingest — ALT
app.post("/api/scans-alt/bulk", async (req, res) => {
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (items.length === 0) return res.json({ ok: true, inserted: 0, skipped: 0 });

  const conn = await pool.getConnection();
  let inserted = 0, skipped = 0;

  try {
    await conn.beginTransaction();

    const sql = `
      INSERT INTO scans_alt (serial, stage, operator, wagon1Id, wagon2Id, wagon3Id,
        receivedAt, loadedAt, destination, grade, railType, spec, lengthM, qrRaw, qrPngPath, timestamp)
      VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE
        stage = VALUES(stage),
        operator = VALUES(operator),
        wagon1Id = VALUES(wagon1Id),
        wagon2Id = VALUES(wagon2Id),
        wagon3Id = VALUES(wagon3Id),
        receivedAt = VALUES(receivedAt),
        loadedAt = VALUES(loadedAt),
        destination = VALUES(destination),
        grade = VALUES(grade),
        railType = VALUES(railType),
        spec = VALUES(spec),
        lengthM = VALUES(lengthM),
        qrRaw = VALUES(qrRaw),
        timestamp = VALUES(timestamp)
    `;

    for (const r of items) {
      const vals = [
        String(r.serial || ""), r.stage || "received", r.operator || "unknown",
        r.wagon1Id || "", r.wagon2Id || "", r.wagon3Id || "",
        r.receivedAt || "", r.loadedAt || "", r.destination || "",
        r.grade || "", r.railType || "", r.spec || "", r.lengthM || "",
        r.qrRaw || "", r.timestamp ? new Date(r.timestamp) : new Date(),
      ];

      try {
        const [result] = await conn.execute(sql, vals);
        let id = result.insertId || null;

        if (!id && r.serial) {
          const [r2] = await conn.query(`SELECT id FROM scans_alt WHERE serial = ? LIMIT 1`, [String(r.serial)]);
          id = r2[0]?.id ?? null;
        }

        if (id) {
          inserted++;
          io.emit("new-scan-alt", {
            id,
            serial: String(r.serial || ""),
            stage: r.stage || "received",
            operator: r.operator || "unknown",
            wagon1Id: r.wagon1Id || "",
            wagon2Id: r.wagon2Id || "",
            wagon3Id: r.wagon3Id || "",
            receivedAt: r.receivedAt || "",
            loadedAt: r.loadedAt || "",
            destination: r.destination || "",
            grade: r.grade || "",
            railType: r.railType || "",
            spec: r.spec || "",
            lengthM: r.lengthM || "",
            qrRaw: r.qrRaw || "",
            timestamp: (r.timestamp ? new Date(r.timestamp) : new Date()).toISOString(),
          });
          generateQrPngAndPersist("scans_alt", id, r.qrRaw || r.serial || "").catch(() => {});
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

// Pagination + count — ALT
app.get("/api/staged-alt", async (req, res) => {
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
      `SELECT * FROM scans_alt ${where} ORDER BY id ${dir} LIMIT ${limit}`,
      params
    );

    const nextCursor = rows.length ? rows[rows.length - 1].id : null;
    const [[c]] = await pool.query(`SELECT COUNT(*) AS c FROM scans_alt`);
    const total = Number(c.c || 0);

    res.json({ rows, nextCursor, total });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/staged-alt/count", async (_req, res) => {
  try {
    const [[c]] = await pool.query(`SELECT COUNT(*) AS c FROM scans_alt`);
    res.json({ count: Number(c.c || 0) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Exists — ALT
app.get("/api/exists-alt/:serial", async (req, res) => {
  const serial = String(req.params.serial || "").trim();
  if (!serial) return res.status(400).json({ error: "Serial required" });

  try {
    const [rows] = await pool.query(`SELECT * FROM scans_alt WHERE serial = ? LIMIT 1`, [serial]);
    if (rows.length) return res.json({ exists: true, row: rows[0] });
    return res.json({ exists: false });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Delete a scan — ALT
app.delete("/api/staged-alt/:id", async (req, res) => {
  const scanId = Number(req.params.id);
  if (!scanId) return res.status(400).json({ error: "Scan ID required" });

  try {
    const [rows] = await pool.query(`SELECT * FROM scans_alt WHERE id = ?`, [scanId]);
    const row = rows[0];
    if (!row) return res.status(404).json({ error: "Scan not found" });

    if (row.qrPngPath) {
      const abs = path.join(__dirname, row.qrPngPath);
      if (fs.existsSync(abs)) fs.unlink(abs, () => {});
    }

    await pool.query(`DELETE FROM scans_alt WHERE id = ?`, [scanId]);
    io.emit("deleted-scan-alt", { id: scanId });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Clear all scans — ALT
app.post("/api/staged-alt/clear", async (req, res) => {
  try {
    console.log("[CLEAR-ALT] /api/staged-alt/clear called - ip:", req.ip, "ua:", req.headers["user-agent"]);

    if (fs.existsSync(QR_DIR_ALT)) {
      for (const f of fs.readdirSync(QR_DIR_ALT)) {
        const p = path.join(QR_DIR_ALT, f);
        try {
          fs.unlinkSync(p);
        } catch (err) {
          console.warn("Failed to unlink ALT QR file:", p, err && err.message);
        }
      }
    }

    await pool.query(`DELETE FROM scans_alt`);
    io.emit("cleared-scans-alt");
    res.json({ ok: true });
  } catch (e) {
    console.error("[CLEAR-ALT] error:", e);
    res.status(500).json({ error: e.message });
  }
});

// Export to .xlsm — ALT
app.post("/api/export-alt-to-excel", async (_req, res) => {
  try {
    const templatePath = path.join(UPLOAD_DIR, "template_alt.xlsm");

    if (!fs.existsSync(templatePath)) {
      console.warn("[EXPORT-ALT-XLSM] template_alt.xlsm not found — falling back");
      return await programmaticXlsmFallback({ table: "scans_alt", outNamePrefix: "Alt" }, res);
    }

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0] || "Sheet1";

    const HEADERS = [
      "Serial", "Stage", "Operator",
      "Wagon1ID", "Wagon2ID", "Wagon3ID",
      "RecievedAt", "LoadedAt", "Destination",
      "Grade", "RailType", "Spec", "Length",
      "QRText", "QRImagePath", "Timestamp",
    ];

    const [rows] = await pool.query(`SELECT * FROM scans_alt ORDER BY id ASC`);

    const dataRows = rows.map((s) => ([
      s.serial || "", s.stage || "", s.operator || "",
      s.wagon1Id || "", s.wagon2Id || "", s.wagon3Id || "",
      s.receivedAt || "", s.loadedAt || "", s.destination || "",
      s.grade || "", s.railType || "", s.spec || "", s.lengthM || "",
      s.qrRaw || "", s.qrPngPath || "",
      s.timestamp ? new Date(s.timestamp).toISOString() : "",
    ]));

    const aoa = [HEADERS, ...dataRows];
    const newWs = XLSX.utils.aoa_to_sheet(aoa);
    wb.Sheets[sheetName] = newWs;

    if (!wb.SheetNames.includes(sheetName)) wb.SheetNames = [sheetName, ...wb.SheetNames];

    const outName = `Alt_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR, outName);
    const outBuffer = XLSX.write(wb, { bookType: "xlsm", bookVBA: true, type: "buffer" });
    fs.writeFileSync(outPath, outBuffer);

    res.download(outPath, outName);
  } catch (err) {
    console.error("Export ALT failed:", err);
    res.status(500).json({ error: err.message });
  }
});

// Export to .xlsx with embedded QR images — ALT
app.all("/api/export-alt-xlsx-images", async (_req, res) => {
  const ExcelJS = await getExcelJS();
  if (!ExcelJS) return res.status(400).json({ error: "exceljs not installed" });

  const QRCode = await getQRCode();

  try {
    const [rows] = await pool.query(`SELECT * FROM scans_alt ORDER BY id ASC`);

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Scans ALT");

    const columns = [
      { header: "Serial", key: "serial", width: 22 },
      { header: "Stage", key: "stage", width: 12 },
      { header: "Operator", key: "operator", width: 18 },
      { header: "Wagon1ID", key: "wagon1Id", width: 14 },
      { header: "Wagon2ID", key: "wagon2Id", width: 14 },
      { header: "Wagon3ID", key: "wagon3Id", width: 14 },
      { header: "RecievedAt", key: "receivedAt", width: 18 },
      { header: "LoadedAt", key: "loadedAt", width: 18 },
      { header: "Destination", key: "destination", width: 20 },
      { header: "Grade", key: "grade", width: 12 },
      { header: "RailType", key: "railType", width: 12 },
      { header: "Spec", key: "spec", width: 18 },
      { header: "Length", key: "lengthM", width: 10 },
      { header: "QRText", key: "qrRaw", width: 42 },
      { header: "QR Image", key: "qrImage", width: 16 },
      { header: "Timestamp", key: "timestamp", width: 24 },
    ];

    ws.columns = columns;

    rows.forEach((s) => {
      ws.addRow({
        serial: s.serial || "",
        stage: s.stage || "",
        operator: s.operator || "",
        wagon1Id: s.wagon1Id || "",
        wagon2Id: s.wagon2Id || "",
        wagon3Id: s.wagon3Id || "",
        receivedAt: s.receivedAt || "",
        loadedAt: s.loadedAt || "",
        destination: s.destination || "",
        grade: s.grade || "",
        railType: s.railType || "",
        spec: s.spec || "",
        lengthM: s.lengthM || "",
        qrRaw: s.qrRaw || s.serial || "",
        qrImage: "",
        timestamp: s.timestamp ? new Date(s.timestamp).toISOString() : "",
      });
    });

    ws.getRow(1).font = { bold: true };
    for (let i = 2; i <= rows.length + 1; i++) ws.getRow(i).height = 70;

    if (QRCode) {
      const colToLetter = (index) => {
        let num = index + 1;
        let s = "";
        while (num > 0) {
          const m = (num - 1) % 26;
          s = String.fromCharCode(65 + m) + s;
          num = Math.floor((num - 1) / 26);
        }
        return s;
      };

      const qrImageColIndex = columns.findIndex((c) => c.key === "qrImage");
      const colLetter = colToLetter(qrImageColIndex);
      let imagesAdded = 0;

      for (let i = 0; i < rows.length; i++) {
        const text = rows[i].qrRaw || rows[i].serial || "";
        if (!text) continue;

        const buf = await QRCode.toBuffer(text, { type: "png", margin: 1, scale: 4 });
        const imgId = wb.addImage({ buffer: buf, extension: "png" });
        const rowNumber = i + 2;
        const range = `${colLetter}${rowNumber}:${colLetter}${rowNumber}`;
        ws.addImage(imgId, range);
        imagesAdded++;
      }
      console.log("[EXPORT-ALT-XLSX-IMG] images added:", imagesAdded);
    }

    const outName = `Alt_QR_${Date.now()}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${outName}"`);
    await wb.xlsx.write(res);
    res.end();
  } catch (e) {
    console.error("Export (alt xlsx images) failed:", e);
    res.status(500).json({ error: e.message });
  }
});

// ---------- Start ----------
const PORT = process.env.PORT || 4000;
server.listen(PORT, "0.0.0.0", () => console.log(`✅ Backend + Socket.IO listening on :${PORT}`));
