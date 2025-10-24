// index.js — Express + Postgres + Excel export + QR images + pagination + bulk ingest
// Concurrency-friendly: uses PostgreSQL with pooled connections & ON CONFLICT upserts.

import express from "express";
import cors from "cors";
import fs from "fs";
import path from "path";
import multer from "multer";
import XLSX from "xlsx";
import http from "http";
import { Server } from "socket.io";
import pkg from "pg";
const { Pool } = pkg;

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

// ---------- Postgres ----------
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  // ssl: { rejectUnauthorized: false }, // enable if your provider requires SSL
});

// Bootstrap schema (idempotent)
async function bootstrapDb() {
  const client = await pool.connect();
  try {
    await client.query(`
      CREATE TABLE IF NOT EXISTS scans (
        id          BIGSERIAL PRIMARY KEY,
        serial      TEXT UNIQUE,
        stage       TEXT,
        operator    TEXT,
        wagon1Id    TEXT,
        wagon2Id    TEXT,
        wagon3Id    TEXT,
        receivedAt  TEXT,
        loadedAt    TEXT,
        grade       TEXT,
        railType    TEXT,
        spec        TEXT,
        lengthM     TEXT,
        qrRaw       TEXT,
        qrPngPath   TEXT,
        timestamp   TIMESTAMPTZ
      );
      CREATE INDEX IF NOT EXISTS ix_scans_timestamp ON scans(timestamp);
    `);
  } finally {
    client.release();
  }
}
await bootstrapDb();

// ---------- HTTP + Socket.IO ----------
const server = http.createServer(app);
const io = new Server(server, { cors: { origin: "*" } });

io.on("connection", (socket) => {
  // console.log("socket connected", socket.id);
});

// ---------- Helpers ----------
async function generateQrPngAndPersist(id, text) {
  try {
    const QRCode = await getQRCode();
    if (!QRCode || !id || !text) return;
    const pngPath = path.join(QR_DIR, `${id}.png`);
    const pngRel  = path.relative(__dirname, pngPath).replace(/\\/g, "/");
    await QRCode.toFile(pngPath, text, { type: "png", margin: 1, scale: 4 });
    await pool.query(`UPDATE scans SET "qrPngPath"=$1 WHERE id=$2`, [pngRel, id]);
  } catch (e) {
    console.warn("QR PNG generation failed:", e.message);
  }
}

// ---------- API Routes ----------

// Version + health
app.get("/api/version", (_req, res) => {
  res.json({ ok: true, version: "pg-v1" });
});
app.get("/api/health", async (_req, res) => {
  try {
    await pool.query("SELECT 1");
    res.json({ ok: true, db: true });
  } catch {
    res.status(500).json({ ok: false, db: false });
  }
});

// Add a new scan — concurrent, fast. Uses ON CONFLICT to ignore duplicates by serial.
// Responds immediately; QR PNG generated in background.
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
    const q = `
      INSERT INTO scans
        (serial, stage, operator, "wagon1Id", "wagon2Id", "wagon3Id",
         "receivedAt", "loadedAt", grade, "railType", spec, "lengthM",
         "qrRaw", "qrPngPath", timestamp)
      VALUES
        ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,'',$14)
      ON CONFLICT (serial) DO NOTHING
      RETURNING id
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
    const { rows } = await pool.query(q, vals);
    const newId = rows[0]?.id ?? null;

    // Emit immediately (even if duplicate, other clients may already have it)
    io.emit("new-scan", {
      id: newId,
      serial, stage: stage || "received", operator: operator || "unknown",
      wagon1Id: wagon1Id || "", wagon2Id: wagon2Id || "", wagon3Id: wagon3Id || "",
      receivedAt: receivedAt || "", loadedAt: loadedAt || "",
      grade: grade || "", railType: railType || "", spec: spec || "", lengthM: lengthM || "",
      qrRaw: qrRaw || "", timestamp: ts.toISOString(),
    });

    res.json({ ok: true, id: newId });

    // Background QR png (only if actually inserted)
    if (newId) generateQrPngAndPersist(newId, qrRaw || serial);
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// Bulk ingest — ON CONFLICT DO NOTHING for duplicates; runs inside a transaction.
app.post("/api/scans/bulk", async (req, res) => {
  const items = Array.isArray(req.body?.items) ? req.body.items : [];
  if (items.length === 0) return res.json({ ok: true, inserted: 0, skipped: 0 });

  const client = await pool.connect();
  let inserted = 0, skipped = 0;
  try {
    await client.query("BEGIN");
    const text = `
      INSERT INTO scans
        (serial, stage, operator, "wagon1Id", "wagon2Id", "wagon3Id",
         "receivedAt", "loadedAt", grade, "railType", spec, "lengthM",
         "qrRaw", "qrPngPath", timestamp)
      VALUES
        ($1,$2,$3,$4,$5,$6,$7,$8,$9,$10,$11,$12,$13,'',$14)
      ON CONFLICT (serial) DO NOTHING
      RETURNING id
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
        const { rows } = await client.query(text, vals);
        const id = rows[0]?.id;
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
          // Fire-and-forget png
          generateQrPngAndPersist(id, r.qrRaw || r.serial || "").catch(()=>{});
        } else {
          skipped++;
        }
      } catch {
        skipped++;
      }
    }
    await client.query("COMMIT");
    res.json({ ok: true, inserted, skipped });
  } catch (e) {
    await client.query("ROLLBACK");
    console.error(e);
    res.status(500).json({ error: e.message });
  } finally {
    client.release();
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
      where = dir === "DESC" ? "WHERE id < $1" : "WHERE id > $1";
      params.push(cursor);
    }
    const rows = (await pool.query(
      `SELECT * FROM scans ${where} ORDER BY id ${dir} LIMIT ${limit}`,
      params
    )).rows;

    const nextCursor = rows.length ? rows[rows.length - 1].id : null;
    const total = Number((await pool.query(`SELECT COUNT(*) AS c FROM scans`)).rows[0].c || 0);

    res.json({ rows, nextCursor, total });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.get("/api/staged/count", async (_req, res) => {
  try {
    const { rows } = await pool.query(`SELECT COUNT(*) AS c FROM scans`);
    res.json({ count: Number(rows[0].c || 0) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// Delete a scan
app.delete("/api/staged/:id", async (req, res) => {
  const scanId = Number(req.params.id);
  if (!scanId) return res.status(400).json({ error: "Scan ID required" });
  try {
    const { rows } = await pool.query(`SELECT * FROM scans WHERE id=$1`, [scanId]);
    const row = rows[0];
    if (!row) return res.status(404).json({ error: "Scan not found" });

    if (row.qrPngPath) {
      const abs = path.join(__dirname, row.qrPngPath);
      fs.existsSync(abs) && fs.unlink(abs, () => {});
    }
    await pool.query(`DELETE FROM scans WHERE id=$1`, [scanId]);
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

    const { rows } = await pool.query(`SELECT * FROM scans ORDER BY id ASC`);
    const dataRows = rows.map((s) => ([
      s.serial || "", s.stage || "", s.operator || "",
      s.wagon1id || "", s.wagon2id || "", s.wagon3id || "",
      s.receivedat || "", s.loadedat || "",
      s.grade || "", s.railtype || "", s.spec || "", s.lengthm || "",
      s.qrraw || "", s.qrpngpath || "",
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
    const { rows } = await pool.query(`SELECT * FROM scans ORDER BY id ASC`);

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
        wagon1Id:   s.wagon1id || "",
        wagon2Id:   s.wagon2id || "",
        wagon3Id:   s.wagon3id || "",
        receivedAt: s.receivedat || "",
        loadedAt:   s.loadedat || "",
        grade:      s.grade || "",
        railType:   s.railtype || "",
        spec:       s.spec || "",
        lengthM:    s.lengthm || "",
        qrRaw:      s.qrraw || s.serial || "",
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
        const text = rows[i].qrraw || rows[i].serial || "";
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
  console.log(`✅ Backend + Socket.IO + Postgres on :${PORT}`)
);
