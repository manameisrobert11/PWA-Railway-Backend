// index.js — Express + MySQL + Excel export + QR images + pagination + bulk ingest + Socket.IO + SHEETS

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
console.log("BOOT TAG:", "2025-11-07-rail-v2-sheets");

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

// ---- Allowed origins (Netlify + local dev) ----
const ALLOWED_ORIGINS = [
  "https://pwarailway.netlify.app",
  process.env.LOCAL_ORIGIN || "http://localhost:5173",
];

// ---- CORS for REST ----
app.use(cors({
  origin: (origin, cb) => {
    // allow non-browser tools (no origin) and whitelisted origins
    if (!origin || ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
    return cb(new Error(`CORS blocked for origin: ${origin}`));
  },
  methods: ["GET","POST","DELETE","PUT","OPTIONS"],
  credentials: true,
}));
app.use(express.json({ limit: "256kb" }));

const UPLOAD_DIR = path.join(__dirname, "uploads");
if (!fs.existsSync(UPLOAD_DIR)) fs.mkdirSync(UPLOAD_DIR, { recursive: true });
const QR_DIR = path.join(UPLOAD_DIR, "qrcodes");
if (!fs.existsSync(QR_DIR)) fs.mkdirSync(QR_DIR, { recursive: true });
const upload = multer({ dest: UPLOAD_DIR });

// ---------- TEMP OVERRIDE (remove after things are stable)
if (!process.env.MYSQL_URL && !process.env.MYSQL_HOST) {
  process.env.MYSQL_URL = "mysql://railuser:Test1234!@mysql-1ec8:3306/rail";
  console.log("[TEMP OVERRIDE] Set MYSQL_URL to mysql-1ec8 for this boot.");
}

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

console.log("[ENV SNAPSHOT]", {
  MYSQL_URL: process.env.MYSQL_URL || null,
  MYSQL_HOST: process.env.MYSQL_HOST || null,
  MYSQL_PORT: process.env.MYSQL_PORT || null,
  MYSQL_DATABASE: process.env.MYSQL_DATABASE || null,
  MYSQL_USER: process.env.MYSQL_USER || null,
  MYSQL_SSL: process.env.MYSQL_SSL || null,
  NODE_VERSION: process.env.NODE_VERSION || null,
});

console.log("[DB CONFIG about to use]", {
  host: baseConfig.host,
  port: baseConfig.port,
  user: baseConfig.user,
  db: baseConfig.database,
  ssl: !!baseConfig.ssl,
});

// Hard-fail if we’d connect to localhost in prod hosts
if (!baseConfig.host || baseConfig.host === "localhost" || baseConfig.host === "127.0.0.1") {
  throw new Error("DB host resolved to localhost/empty. Set MYSQL_URL or MYSQL_HOST to your Render MySQL hostname (e.g., mysql-1ec8).");
}

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
    await conn.query(`
      CREATE TABLE IF NOT EXISTS \`scans\` (
        \`id\`          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        \`sheet\`       VARCHAR(64) NOT NULL DEFAULT 'main',
        \`serial\`      VARCHAR(191),
        \`stage\`       VARCHAR(64),
        \`operator\`    VARCHAR(128),
        \`wagon1Id\`    VARCHAR(128),
        \`wagon2Id\`    VARCHAR(128),
        \`wagon3Id\`    VARCHAR(128),
        \`receivedAt\`  VARCHAR(64),
        \`loadedAt\`    VARCHAR(64),
        \`grade\`       VARCHAR(64),
        \`railType\`    VARCHAR(64),
        \`spec\`        VARCHAR(128),
        \`lengthM\`     VARCHAR(32),
        \`qrRaw\`       TEXT,
        \`qrPngPath\`   TEXT,
        \`timestamp\`   DATETIME,
        INDEX \`ix_scans_timestamp\` (\`timestamp\`),
        INDEX \`ix_scans_sheet\` (\`sheet\`),
        UNIQUE KEY \`ux_sheet_serial\` (\`sheet\`, \`serial\`)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    `);
  } finally {
    conn.release();
  }
}

// wait for DB reachable with a few retries
async function waitForDb(retries = 12) {
  for (let i = 1; i <= retries; i++) {
    try {
      await pool.query("SELECT 1");
      console.log("DB reachable");
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
      return cb(new Error(`Socket CORS blocked for origin: ${origin}`));
    },
    methods: ["GET","POST"],
    credentials: true,
  },
});

io.on("connection", (socket) => {
  console.log("socket connected:", socket.id, "from", socket.handshake.headers.origin || "unknown");
  socket.on("disconnect", (reason) => {
    console.log("socket disconnected:", socket.id, reason);
  });
});

// ---------- Helpers ----------
async function generateQrPngAndPersist(id, text) {
  try {
    const QRCode = await getQRCode();
    if (!QRCode || !id || !text) return;
    const pngPath = path.join(QR_DIR, `${id}.png`);
    const pngRel  = path.relative(__dirname, pngPath).replace(/\\/g, "/");
    await QRCode.toFile(pngPath, text, { type: "png", margin: 1, scale: 4 });
    await pool.query(`UPDATE \`scans\` SET \`qrPngPath\` = ? WHERE \`id\` = ?`, [pngRel, id]);
  } catch (e) {
    console.warn("QR PNG generation failed:", e.message);
  }
}

function sanitizeSheet(s) {
  const v = (s || 'main').toString().trim().toLowerCase();
  return v === 'alt' ? 'alt' : 'main';
}

// ---------- API Routes ----------

// Version + health + socket test
app.get("/", (_req, res) => res.send("Rail backend is running."));
app.get("/api/version", (_req, res) => {
  res.json({ ok: true, version: "mysql-v2-sheets", bootTag: "2025-11-07-rail-v2-sheets" });
});
app.get("/api/health", async (_req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ ok: true, db: true });
  } catch {
    res.status(500).json({ ok: false, db: false });
  }
});
app.get("/socket-test", (_req, res) => res.type("text/plain").send("OK"));

// Existence check (sheet-aware)
app.get("/api/exists/:serial", async (req, res) => {
  const sheet = sanitizeSheet(req.query.sheet);
  const serial = (req.params.serial || '').toString();
  if (!serial) return res.json({ exists: false });
  try {
    const [rows] = await pool.query(
      `SELECT * FROM \`scans\` WHERE \`sheet\` = ? AND \`serial\` = ? LIMIT 1`,
      [sheet, serial]
    );
    if (rows.length) return res.json({ exists: true, row: rows[0] });
    return res.json({ exists: false });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Add a new scan — UPSERT by (sheet, serial)
app.post("/api/scan", async (req, res) => {
  const {
    sheet: rawSheet,
    serial, stage, operator,
    wagon1Id, wagon2Id, wagon3Id,
    receivedAt, loadedAt,
    grade, railType, spec, lengthM,
    qrRaw, timestamp,
    destination, // optional front-end field already present in your code
  } = req.body;

  if (!serial) return res.status(400).json({ error: "Serial required" });

  const sheet = sanitizeSheet(rawSheet);
  const ts = timestamp ? new Date(timestamp) : new Date();
  try {
    const sql = `
      INSERT INTO \`scans\`
        (\`sheet\`, \`serial\`, \`stage\`, \`operator\`, \`wagon1Id\`, \`wagon2Id\`, \`wagon3Id\`,
         \`receivedAt\`, \`loadedAt\`, \`grade\`, \`railType\`, \`spec\`, \`lengthM\`,
         \`qrRaw\`, \`qrPngPath\`, \`timestamp\`)
      VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE
        \`stage\`      = VALUES(\`stage\`),
        \`operator\`   = VALUES(\`operator\`),
        \`wagon1Id\`   = VALUES(\`wagon1Id\`),
        \`wagon2Id\`   = VALUES(\`wagon2Id\`),
        \`wagon3Id\`   = VALUES(\`wagon3Id\`),
        \`receivedAt\` = VALUES(\`receivedAt\`),
        \`loadedAt\`   = VALUES(\`loadedAt\`),
        \`grade\`      = VALUES(\`grade\`),
        \`railType\`   = VALUES(\`railType\`),
        \`spec\`       = VALUES(\`spec\`),
        \`lengthM\`    = VALUES(\`lengthM\`),
        \`qrRaw\`      = VALUES(\`qrRaw\`),
        \`timestamp\`  = VALUES(\`timestamp\`)
    `;
    const vals = [
      sheet,
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

    let newId = result.insertId || null;
    if (!newId) {
      const [r2] = await pool.query(
        `SELECT \`id\` FROM \`scans\` WHERE \`sheet\`=? AND \`serial\`=? LIMIT 1`,
        [sheet, String(serial)]
      );
      newId = r2[0]?.id ?? null;
    }

    const payload = {
      id: newId,
      sheet,
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
      destination: destination || "",
    };

    io.emit("new-scan", payload);

    res.json({ ok: true, id: newId });
    if (newId) generateQrPngAndPersist(newId, qrRaw || serial);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Bulk ingest — UPSERT by (sheet, serial)
app.post("/api/scans/bulk", async (req, res) => {
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (items.length === 0) return res.json({ ok: true, inserted: 0, skipped: 0 });

  const conn = await pool.getConnection();
  let inserted = 0, skipped = 0;
  try {
    await conn.beginTransaction();

    const text = `
      INSERT INTO \`scans\`
        (\`sheet\`, \`serial\`, \`stage\`, \`operator\`, \`wagon1Id\`, \`wagon2Id\`, \`wagon3Id\`,
         \`receivedAt\`, \`loadedAt\`, \`grade\`, \`railType\`, \`spec\`, \`lengthM\`,
         \`qrRaw\`, \`qrPngPath\`, \`timestamp\`)
      VALUES
        (?,?,?,?,?,?,?,?,?,?,?,?,?,'',?)
      ON DUPLICATE KEY UPDATE
        \`stage\`      = VALUES(\`stage\`),
        \`operator\`   = VALUES(\`operator\`),
        \`wagon1Id\`   = VALUES(\`wagon1Id\`),
        \`wagon2Id\`   = VALUES(\`wagon2Id\`),
        \`wagon3Id\`   = VALUES(\`wagon3Id\`),
        \`receivedAt\` = VALUES(\`receivedAt\`),
        \`loadedAt\`   = VALUES(\`loadedAt\`),
        \`grade\`      = VALUES(\`grade\`),
        \`railType\`   = VALUES(\`railType\`),
        \`spec\`       = VALUES(\`spec\`),
        \`lengthM\`    = VALUES(\`lengthM\`),
        \`qrRaw\`      = VALUES(\`qrRaw\`),
        \`timestamp\`  = VALUES(\`timestamp\`)
    `;

    for (const r of items) {
      const sheet = sanitizeSheet(r.sheet);
      const vals = [
        sheet,
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
          const [r2] = await conn.query(
            `SELECT \`id\` FROM \`scans\` WHERE \`sheet\`=? AND \`serial\`=? LIMIT 1`,
            [sheet, String(r.serial)]
          );
          id = r2[0]?.id ?? null;
        }
        if (id) {
          inserted++;
          io.emit("new-scan", {
            id,
            sheet,
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
            destination: r.destination || "",
          });
          generateQrPngAndPersist(id, r.qrRaw || r.serial || "").catch(() => {});
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

// Pagination + count (sheet-aware)
app.get("/api/staged", async (req, res) => {
  const limit = Math.min(parseInt(req.query.limit || "100", 10), 500);
  const cursor = req.query.cursor ? parseInt(req.query.cursor, 10) : null;
  const dir = (req.query.dir || "desc").toLowerCase() === "asc" ? "ASC" : "DESC";
  const sheet = sanitizeSheet(req.query.sheet);

  try {
    const params = [sheet];
    let where = "WHERE `sheet` = ?";
    if (cursor != null) {
      where += dir === "DESC" ? " AND `id` < ?" : " AND `id` > ?";
      params.push(cursor);
    }
    const [rows] = await pool.query(
      `SELECT * FROM \`scans\` ${where} ORDER BY \`id\` ${dir} LIMIT ${limit}`, params
    );

    const nextCursor = rows.length ? rows[rows.length - 1].id : null;
    const [[c]] = await pool.query(`SELECT COUNT(*) AS c FROM \`scans\` WHERE \`sheet\` = ?`, [sheet]);
    const total = Number(c.c || 0);

    res.json({ rows, nextCursor, total });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/staged/count", async (req, res) => {
  const sheet = sanitizeSheet(req.query.sheet);
  try {
    const [[c]] = await pool.query(`SELECT COUNT(*) AS c FROM \`scans\` WHERE \`sheet\` = ?`, [sheet]);
    res.json({ count: Number(c.c || 0) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Delete a scan (also sheet-aware broadcast)
app.delete("/api/staged/:id", async (req, res) => {
  const scanId = Number(req.params.id);
  if (!scanId) return res.status(400).json({ error: "Scan ID required" });
  try {
    const [rows] = await pool.query(`SELECT * FROM \`scans\` WHERE \`id\` = ?`, [scanId]);
    const row = rows[0];
    if (!row) return res.status(404).json({ error: "Scan not found" });

    if (row.qrPngPath) {
      const abs = path.join(__dirname, row.qrPngPath);
      fs.existsSync(abs) && fs.unlink(abs, () => {});
    }
    await pool.query(`DELETE FROM \`scans\` WHERE \`id\` = ?`, [scanId]);
    io.emit("deleted-scan", { id: scanId, sheet: row.sheet });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Clear all scans on a given sheet (SAFE; images for that sheet only)
app.post("/api/staged/clear", async (req, res) => {
  const sheet = sanitizeSheet(req.query.sheet || req.body?.sheet);
  try {
    const [rows] = await pool.query(
      `SELECT id, qrPngPath FROM \`scans\` WHERE \`sheet\` = ?`,
      [sheet]
    );

    for (const r of rows) {
      if (!r.qrPngPath) continue;
      const abs = path.join(__dirname, r.qrPngPath);
      if (fs.existsSync(abs)) {
        try { fs.unlinkSync(abs); } catch {}
      }
    }

    await pool.query(`DELETE FROM \`scans\` WHERE \`sheet\` = ?`, [sheet]);
    io.emit("cleared-scans", { sheet });
    res.json({ ok: true, cleared: rows.length, sheet });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Upload Excel template (.xlsm)
app.post("/api/upload-template", upload.single("template"), (req, res) => {
  res.json({ ok: true, path: req.file?.path });
});

// Export to .xlsm (macro-enabled) — sheet-aware
app.post("/api/export-to-excel", async (req, res) => {
  const sheet = sanitizeSheet(req.query.sheet || req.body?.sheet);
  try {
    const templatePath = path.join(UPLOAD_DIR, "template.xlsm");
    if (!fs.existsSync(templatePath)) {
      return res.status(400).json({ error: "template.xlsm not found" });
    }

    const wb = XLSX.readFile(templatePath, { cellDates: true, bookVBA: true });
    const sheetName = wb.SheetNames[0];

    const HEADERS = [
      "Sheet","Serial","Stage","Operator",
      "Wagon1ID","Wagon2ID","Wagon3ID",
      "RecievedAt","LoadedAt",
      "Grade","RailType","Spec","Length",
      "QRText","QRImagePath",
      "Timestamp",
    ];

    const [rows] = await pool.query(
      `SELECT * FROM \`scans\` WHERE \`sheet\` = ? ORDER BY \`id\` ASC`,
      [sheet]
    );
    const dataRows = rows.map((s) => ([
      s.sheet || "main",
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

    const outName = `Master_${sheet}_${Date.now()}.xlsm`;
    const outPath = path.join(UPLOAD_DIR, outName);
    XLSX.writeFile(wb, outPath, { bookType: "xlsm", bookVBA: true });
    res.download(outPath, outName);
  } catch (err) {
    console.error("Export failed:", err);
    res.status(500).json({ error: err.message });
  }
});

// Export to .xlsx with embedded QR images (sheet-aware)
app.all("/api/export-xlsx-images", async (req, res) => {
  const ExcelJS = await getExcelJS();
  if (!ExcelJS) return res.status(400).json({ error: "exceljs not installed. Run: npm i exceljs qrcode" });
  const QRCode = await getQRCode();
  const sheet = sanitizeSheet(req.query.sheet || req.body?.sheet);

  try {
    const [rows] = await pool.query(
      `SELECT * FROM \`scans\` WHERE \`sheet\` = ? ORDER BY \`id\` ASC`,
      [sheet]
    );

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("Scans");

    const columns = [
      { header: "Sheet",       key: "sheet",       width: 10 },
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
        sheet:      s.sheet || "main",
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

    const outName = `Master_QR_${sheet}_${Date.now()}.xlsx`;
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
  console.log(`✅ Backend + Socket.IO + MySQL (sheets) on :${PORT}`)
);
