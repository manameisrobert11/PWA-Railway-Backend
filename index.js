import express from "express";
import http from "http";
import { Server } from "socket.io";
import mysql from "mysql2/promise";
import cors from "cors";

/* =========================
   App & Server
========================= */

const app = express();
const server = http.createServer(app);

const io = new Server(server, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"]
  }
});

app.use(cors());
app.use(express.json());

/* =========================
   Config
========================= */

const PORT = process.env.PORT || 10000;

const DB_CONFIG = {
  host: process.env.MYSQL_HOST,
  port: Number(process.env.MYSQL_PORT || 3306),
  user: process.env.MYSQL_USER,
  password: process.env.MYSQL_PASSWORD,
  database: process.env.MYSQL_DATABASE,
  ssl: process.env.MYSQL_SSL === "true"
};

/* =========================
   State
========================= */

let pool = null;
let dbReady = false;

/**
 * Offline queue:
 * stores DB operations while DB is unavailable
 */
const offlineQueue = [];

/* =========================
   DB Connection (Background)
========================= */

async function connectDbWithRetry(maxAttempts = 12, delayMs = 3000) {
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      pool = mysql.createPool({
        ...DB_CONFIG,
        waitForConnections: true,
        connectionLimit: 10,
        queueLimit: 0
      });

      await pool.query("SELECT 1");
      dbReady = true;

      console.log("✅ DB connected");
      await flushOfflineQueue();
      return;
    } catch (err) {
      dbReady = false;
      console.warn(`DB not ready (attempt ${attempt}/${maxAttempts})`, err.code);
      await new Promise(r => setTimeout(r, delayMs));
    }
  }

  console.error("❌ DB failed to connect after retries");
}

/* =========================
   Offline Queue Handling
========================= */

async function flushOfflineQueue() {
  if (!dbReady || !pool || offlineQueue.length === 0) return;

  console.log(`🔄 Flushing ${offlineQueue.length} queued operations`);

  while (offlineQueue.length > 0) {
    const job = offlineQueue.shift();
    try {
      await pool.query(job.sql, job.values);
    } catch (err) {
      console.error("❌ Failed to flush queued job, re-queueing");
      offlineQueue.unshift(job);
      break;
    }
  }
}

/* =========================
   Helpers
========================= */

function requireDb(req, res, next) {
  if (!dbReady) {
    return res.status(503).json({ error: "Database warming up" });
  }
  next();
}

async function safeQuery(sql, values) {
  if (!dbReady || !pool) {
    offlineQueue.push({ sql, values });
    return { queued: true };
  }

  return pool.query(sql, values);
}

/* =========================
   Routes
========================= */

app.get("/api/health", (_req, res) => {
  res.status(dbReady ? 200 : 503).json({
    ok: dbReady,
    db: dbReady,
    queuedWrites: offlineQueue.length
  });
});

app.get("/api/staged", requireDb, async (req, res) => {
  const [rows] = await pool.query(
    "SELECT * FROM staged ORDER BY id DESC LIMIT 200"
  );
  res.json(rows);
});

app.get("/api/staged/count", requireDb, async (_req, res) => {
  const [[row]] = await pool.query(
    "SELECT COUNT(*) as count FROM staged"
  );
  res.json({ count: row.count });
});

app.post("/api/staged", async (req, res) => {
  const { name, value } = req.body;

  const result = await safeQuery(
    "INSERT INTO staged (name, value) VALUES (?, ?)",
    [name, value]
  );

  io.emit("staged:new", { name, value });

  res.json({
    ok: true,
    queued: result?.queued === true
  });
});

/* =========================
   Socket.IO
========================= */

io.on("connection", socket => {
  console.log("🔌 Socket connected:", socket.id);

  socket.on("disconnect", () => {
    console.log("🔌 Socket disconnected:", socket.id);
  });
});

/* =========================
   START SERVER FIRST (IMPORTANT)
========================= */

server.listen(PORT, "0.0.0.0", () => {
  console.log(`✅ Backend + Socket.IO listening on :${PORT}`);
});

/* =========================
   Boot DB in Background
========================= */

connectDbWithRetry();
