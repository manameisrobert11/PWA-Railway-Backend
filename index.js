// index.js — Express + MySQL + Excel export + QR images + pagination + bulk ingest + Socket.IO
import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import multer from "multer";
import XLSX from "xlsx";
import http from "http";
import { Server } from "socket.io";
import mysql from "mysql2/promise";

// ----- BOOT TAG
console.log("BOOT TAG:", "2025-11-07-rail-v2-main-alt");

// ---------- Lazy loaders (optional deps) ----------
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

app.use(cors({
  origin: (origin, cb) => {
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
// ALT QR folder
const QR_DIR_ALT = path.join(UPLOAD_DIR, "qrcodes_alt");
if (!fs.existsSync(QR_DIR_ALT)) fs.mkdirSync(QR_DIR_ALT, { recursive: true });

const upload = multer({ dest: UPLOAD_DIR });

// ---------- TEMP OVERRIDE (only for devs missing env) ----------
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

if (!baseConfig.host || baseConfig.host === "localhost" || baseConfig.host === "127.0.0.1") {
  throw new Error("DB host resolved to localhost/empty. Set MYSQL_URL or MYSQL_HOST to your Render MySQL hostname.");
}

export const pool = mysql.createPool({
  ...baseConfig,
  connectionLimit: 30,
  waitForConnections: true,
  queueLimit: 0,
});

// ---------- Bootstrap schema ----------
async function bootstrapDb() {
  const conn = await pool.getConnection();
  try {
    // MAIN
    await conn.query(`
      CREATE TABLE IF NOT EXISTS \`scans\` (
        \`id\`          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        \`serial\`      VARCHAR(191) UNIQUE,
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
        INDEX \`ix_scans_timestamp\` (\`timestamp\`)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    `);

    // ALT (mirrors main)
    await conn.query(`
      CREATE TABLE IF NOT EXISTS \`scans_alt\` (
        \`id\`          BIGINT UNSIGNED NOT NULL AUTO_INCREMENT PRIMARY KEY,
        \`serial\`      VARCHAR(191) UNIQUE,
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
        INDEX \`ix_scans_alt_timestamp\` (\`timestamp\`)
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4;
    `);
  } finally {
    conn.release();
  }
}

// wait for DB reachable
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
async function generateQrPngAndPersist(table, id, text) {
  try {
    const QRCode = await getQRCode();
    if (!QRCode || !id || !text) return;

    const dir = table === 'scans_alt' ? QR_DIR_ALT : QR_DIR;
    const pngPath = path.join(dir, `${id}.png`);
    const pngRel  = path.relative(__dirname, pngPath).replace(/\\/g, "/");

    await QRCode.toFile(pngPath, text, { type: "png", margin: 1, scale: 4 });
    await pool.query(`UPDATE \`${table}\` SET \`qrPngPath\` = ? WHERE \`id\` = ?`, [pngRel, id]);
  } catch (e) {
    console.warn("QR PNG generation failed:", e.message);
  }
}

// ---------- Health ----------
app.get("/", (_req, res) => res.send("Rail backend is running."));
app.get("/api/version", (_req, res) => {
  res.json({ ok: true, version: "mysql-v2", bootTag: "2025-11-07-rail-v2-main-alt" });
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

// ---------- MAIN ROUTES ----------
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
      INSERT INTO \`scans\`
        (\`serial\`, \`stage\`, \`operator\`, \`wagon1Id\`, \`wagon2Id\`, \`wagon3Id\`,
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
      const [r2] = await pool.query(`SELECT \`id\` FROM \`scans\` WHERE \`serial\` = ? LIMIT 1`, [String(serial)]);
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
      grade: grade || "",
      railType: railType || "",
      spec: spec || "",
      lengthM: lengthM || "",
      qrRaw: qrRaw || "",
      timestamp: ts.toISOString(),
    };

    io.emit("new-scan", payload);
    res.json({ ok: true, id: newId });
    if (newId) generateQrPngAndPersist('scans', newId, qrRaw || serial);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/scans/bulk", async (req, res) => {
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (items.length === 0) return res.json({ ok: true, inserted: 0, skipped: 0 });

  const conn = await pool.getConnection();
  let inserted = 0, skipped = 0;
  try {
    await conn.beginTransaction();

    const text = `
      INSERT INTO \`scans\`
        (\`serial\`, \`stage\`, \`operator\`, \`wagon1Id\`, \`wagon2Id\`, \`wagon3Id\`,
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
        const [result] = await conn.execute(text, vals
