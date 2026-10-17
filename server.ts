import express from "express";
import path from "path";
import { createServer as createViteServer } from "vite";
import dotenv from "dotenv";
import fs from "fs/promises";
import { existsSync } from "fs";
import { createRemoteJWKSet, jwtVerify } from "jose";

dotenv.config();

const app = express();
app.use(express.json({ limit: "50mb" }));

const PORT = 3000;

// ---------------- Entra ID token verification ----------------
// Same tenant/app used by src/lib/entraAuth.ts on the frontend.
const TENANT_ID = process.env.ENTRA_TENANT_ID || "b7be7686-6f97-4db7-9081-a23cf09a96b5";
const CLIENT_ID = process.env.ENTRA_CLIENT_ID || "3c402a5a-d1f3-4387-86c5-1418152cc654";
const JWKS = createRemoteJWKSet(
  new URL(`https://login.microsoftonline.com/${TENANT_ID}/discovery/v2.0/keys`)
);

interface CallerIdentity {
  oid: string;
  email: string;
  name: string;
}

async function requireAuth(req: express.Request, res: express.Response, next: express.NextFunction) {
  try {
    const authHeader = req.headers.authorization || "";
    const token = authHeader.startsWith("Bearer ") ? authHeader.slice(7) : null;
    if (!token) {
      return res.status(401).json({ error: "Missing bearer token" });
    }
    const { payload } = await jwtVerify(token, JWKS, {
      issuer: [`https://login.microsoftonline.com/${TENANT_ID}/v2.0`],
      audience: CLIENT_ID,
    });
    const caller: CallerIdentity = {
      oid: String(payload.oid || payload.sub),
      email: String((payload as any).preferred_username || (payload as any).email || ""),
      name: String((payload as any).name || ""),
    };
    (req as any).callerIdentity = caller;
    next();
  } catch (err: any) {
    console.error("Token verification failed:", err.message);
    return res.status(401).json({ error: "Invalid or expired token" });
  }
}

// Central file-based datastore for logs, abbreviations and users
const DATA_DIR = path.join(process.cwd(), "data");
const LOGS_FILE = path.join(DATA_DIR, "logs.json");
const ABBREVIATIONS_FILE = path.join(DATA_DIR, "abbreviations.json");
const USERS_FILE = path.join(DATA_DIR, "users.json");

async function ensureDataDir() {
  try {
    await fs.mkdir(DATA_DIR, { recursive: true });
  } catch (err) {
    console.error("Error creating data directory:", err);
  }
}

async function readUsersFile(): Promise<any[]> {
  await ensureDataDir();
  if (!existsSync(USERS_FILE)) return [];
  try {
    const content = await fs.readFile(USERS_FILE, "utf-8");
    if (!content || !content.trim()) return [];
    const parsed = JSON.parse(content);
    return Array.isArray(parsed) ? parsed : [];
  } catch (e) {
    console.error("Error reading users.json:", e);
    return [];
  }
}

async function writeUsersFile(users: any[]): Promise<void> {
  await ensureDataDir();
  await fs.writeFile(USERS_FILE, JSON.stringify(users, null, 2), "utf-8");
  notifyDataChanged("users", { count: users.length });
}

async function getUserInternal(id: string): Promise<any | null> {
  const users = await readUsersFile();
  return users.find((u: any) => u.id === id) || null;
}

async function upsertUserInternal(user: any): Promise<void> {
  const users = await readUsersFile();
  const idx = users.findIndex((u: any) => u.id === user.id);
  if (idx !== -1) users[idx] = user;
  else users.push(user);
  await writeUsersFile(users);
}

// In-memory store for live-pushed Excel files from local file-watcher script
let latestExcelBuffer: Buffer | null = null;
let latestExcelUpdatedAt: Date | null = null;

// Server-Sent Events (SSE) Real-time Multi-User Broadcast System
const sseClients = new Set<express.Response>();

app.get("/api/events", (req, res) => {
  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.flushHeaders();

  sseClients.add(res);
  res.write(`data: ${JSON.stringify({ type: "connected", timestamp: Date.now() })}\n\n`);

  req.on("close", () => {
    sseClients.delete(res);
  });
});

function notifyDataChanged(topic: string, payload?: any) {
  const message = `data: ${JSON.stringify({ topic, payload, timestamp: Date.now() })}\n\n`;
  for (const client of sseClients) {
    try {
      client.write(message);
    } catch (e) {
      sseClients.delete(client);
    }
  }
}

// Presence / Active Users Tracking System
interface ActiveUserSession {
  sessionId: string;
  userId?: string;
  username?: string;
  name?: string;
  role?: string;
  ip?: string;
  activeTab?: string;
  activeTabName?: string;
  activeUrl?: string;
  lastSeen: number;
}

const activeSessions = new Map<string, ActiveUserSession>();

function getActiveUsersList() {
  const now = Date.now();
  const active: ActiveUserSession[] = [];
  for (const [sessionId, session] of activeSessions.entries()) {
    if (now - session.lastSeen > 12000) {
      activeSessions.delete(sessionId);
    } else {
      active.push(session);
    }
  }
  return active;
}

function notifyPresenceUpdate() {
  const active = getActiveUsersList();
  notifyDataChanged("presence", {
    activeCount: active.length,
    activeUsers: active
  });
}

setInterval(() => {
  const prevCount = activeSessions.size;
  const active = getActiveUsersList();
  if (activeSessions.size !== prevCount) {
    notifyPresenceUpdate();
  }
}, 3000);

app.post("/api/presence/heartbeat", (req, res) => {
  const { sessionId, userId, username, name, role, activeTab, activeTabName, activeUrl } = req.body || {};
  if (!sessionId) {
    return res.status(400).json({ error: "sessionId required" });
  }

  const rawIp = (req.headers["x-forwarded-for"] as string)?.split(",")[0]?.trim() || req.ip || req.socket.remoteAddress || "Client Node";
  let clientIp = rawIp.startsWith("::ffff:") ? rawIp.replace("::ffff:", "") : rawIp;

  const prevCount = activeSessions.size;
  activeSessions.set(sessionId, {
    sessionId,
    userId,
    username: username || "user",
    name: name || "Active User",
    role: role || "viewer",
    ip: clientIp,
    activeTab: activeTab || "logs",
    activeTabName: activeTabName || "MWD/LWD Logs",
    activeUrl: activeUrl || "/#logs",
    lastSeen: Date.now()
  });

  const active = getActiveUsersList();
  if (activeSessions.size !== prevCount) {
    notifyPresenceUpdate();
  }

  res.json({ success: true, activeCount: active.length, activeUsers: active });
});

app.get("/api/presence", (req, res) => {
  const active = getActiveUsersList();
  res.json({ activeCount: active.length, activeUsers: active });
});

// AI Query Route disabled
app.post("/api/query", (req, res) => {
  return res.status(404).json({ error: "AI Assistant has been disabled." });
});

// ---------------- Logs ----------------
app.get("/api/logs", requireAuth, async (req, res) => {
  try {
    await ensureDataDir();
    if (existsSync(LOGS_FILE)) {
      const content = await fs.readFile(LOGS_FILE, "utf-8");
      if (!content || !content.trim()) return res.json({ logs: [], exists: false });
      try {
        const logs = JSON.parse(content);
        return res.json({ logs: Array.isArray(logs) ? logs : [], exists: true });
      } catch (e) {
        return res.json({ logs: [], exists: false });
      }
    }
    return res.json({ logs: [], exists: false });
  } catch (error: any) {
    console.error("Error reading logs:", error);
    res.json({ logs: [], exists: false });
  }
});

app.post("/api/logs", requireAuth, async (req, res) => {
  try {
    await ensureDataDir();
    const { logs } = req.body;
    if (!Array.isArray(logs)) {
      return res.status(400).json({ error: "Logs must be an array" });
    }
    await fs.writeFile(LOGS_FILE, JSON.stringify(logs, null, 2), "utf-8");
    notifyDataChanged("logs", { count: logs.length });
    res.json({ success: true, count: logs.length });
  } catch (error: any) {
    console.error("Error saving logs:", error);
    res.status(500).json({ error: "Failed to save logs to server-side store" });
  }
});

app.post("/api/logs/row", requireAuth, async (req, res) => {
  try {
    await ensureDataDir();
    const { row } = req.body;
    if (!row || !row.id) {
      return res.status(400).json({ error: "Row with a valid ID is required in the body" });
    }
    let logs: any[] = [];
    if (existsSync(LOGS_FILE)) {
      const content = await fs.readFile(LOGS_FILE, "utf-8");
      if (content && content.trim()) {
        try {
          const parsed = JSON.parse(content);
          if (Array.isArray(parsed)) logs = parsed;
        } catch (e) {}
      }
    }
    const idx = logs.findIndex((r: any) => r.id === row.id);
    if (idx !== -1) logs[idx] = row;
    else logs.push(row);
    await fs.writeFile(LOGS_FILE, JSON.stringify(logs, null, 2), "utf-8");
    notifyDataChanged("logs", { rowId: row.id });
    res.json({ success: true, row });
  } catch (error: any) {
    console.error("Error saving row:", error);
    res.status(500).json({ error: "Failed to save row to server-side store" });
  }
});

app.delete("/api/logs/row/:id", requireAuth, async (req, res) => {
  try {
    await ensureDataDir();
    const { id } = req.params;
    if (existsSync(LOGS_FILE)) {
      const content = await fs.readFile(LOGS_FILE, "utf-8");
      let logs: any[] = [];
      if (content && content.trim()) {
        try {
          const parsed = JSON.parse(content);
          if (Array.isArray(parsed)) logs = parsed;
        } catch (e) {}
      }
      const filtered = logs.filter((r: any) => r.id !== id);
      await fs.writeFile(LOGS_FILE, JSON.stringify(filtered, null, 2), "utf-8");
      notifyDataChanged("logs", { deletedId: id });
      return res.json({ success: true, deletedId: id });
    }
    res.json({ success: true, message: "No logs found to delete from" });
  } catch (error: any) {
    console.error("Error deleting row:", error);
    res.status(500).json({ error: "Failed to delete row from server-side store" });
  }
});

// ---------------- Abbreviations ----------------
app.get("/api/abbreviations", requireAuth, async (req, res) => {
  try {
    await ensureDataDir();
    if (existsSync(ABBREVIATIONS_FILE)) {
      const content = await fs.readFile(ABBREVIATIONS_FILE, "utf-8");
      if (content && content.trim()) {
        try {
          const abbreviations = JSON.parse(content);
          return res.json({ abbreviations, exists: true });
        } catch (e) {
          return res.json({ abbreviations: [], exists: false });
        }
      }
    }
    return res.json({ abbreviations: [], exists: false });
  } catch (error: any) {
    console.error("Error reading abbreviations:", error);
    res.status(500).json({ error: "Failed to read abbreviations from server-side store" });
  }
});

app.post("/api/abbreviations", requireAuth, async (req, res) => {
  try {
    await ensureDataDir();
    const { abbreviations } = req.body;
    if (!Array.isArray(abbreviations)) {
      return res.status(400).json({ error: "Abbreviations must be an array" });
    }
    await fs.writeFile(ABBREVIATIONS_FILE, JSON.stringify(abbreviations, null, 2), "utf-8");
    notifyDataChanged("abbreviations", { count: abbreviations.length });
    res.json({ success: true, count: abbreviations.length });
  } catch (error: any) {
    console.error("Error saving abbreviations:", error);
    res.status(500).json({ error: "Failed to save abbreviations to server-side store" });
  }
});

// ---------------- Users (profiles + roles + approval status) ----------------
// Authorization rules, all enforced server-side — the client is never trusted
// to say who it is or what role/status it should get:
//  - Writing your OWN profile for the first time: if the users store is
//    completely empty, you become admin + approved (bootstrap — there's no
//    admin yet to grant it). Every subsequent new sign-in starts as
//    viewer + pending, no matter what the request body claims.
//  - Writing your OWN profile when one already exists: role/status/permissions
//    are preserved from the existing record — you cannot self-promote or
//    self-approve by calling this again. Only cosmetic fields refresh.
//  - Writing someone ELSE's profile (role changes, approve/reject/revoke):
//    only allowed if your own stored role is currently "admin".

app.get("/api/users", requireAuth, async (req, res) => {
  try {
    const caller = (req as any).callerIdentity as CallerIdentity;
    const callerRecord = await getUserInternal(caller.oid);
    if (!callerRecord || callerRecord.role !== "admin") {
      return res.status(403).json({ error: "Admin access required" });
    }
    const users = await readUsersFile();
    res.json({ users });
  } catch (error: any) {
    console.error("Error reading users:", error);
    res.status(500).json({ error: "Failed to read users from server-side store" });
  }
});

app.get("/api/users/:id", requireAuth, async (req, res) => {
  try {
    const caller = (req as any).callerIdentity as CallerIdentity;
    const { id } = req.params;
    if (id !== caller.oid) {
      const callerRecord = await getUserInternal(caller.oid);
      if (!callerRecord || callerRecord.role !== "admin") {
        return res.status(403).json({ error: "Admin access required" });
      }
    }
    const user = await getUserInternal(id);
    if (!user) return res.status(404).json({ error: "User not found" });
    res.json(user);
  } catch (error: any) {
    res.status(500).json({ error: "Failed to fetch user" });
  }
});

app.post("/api/users/row", requireAuth, async (req, res) => {
  try {
    const caller = (req as any).callerIdentity as CallerIdentity;
    const { user } = req.body;
    if (!user || !user.id) {
      return res.status(400).json({ error: "User with a valid ID is required in the body" });
    }

    let finalUser: any;

    if (user.id === caller.oid) {
      const existing = await getUserInternal(caller.oid);
      if (!existing) {
        const allUsers = await readUsersFile();
        const isFirstUserEver = allUsers.length === 0;
        finalUser = isFirstUserEver
          ? {
              id: caller.oid,
              email: caller.email,
              name: user.name || caller.name,
              role: "admin",
              permissions: "both",
              status: "approved",
              title: "Administrator",
              createdAt: new Date().toISOString(),
            }
          : {
              id: caller.oid,
              email: caller.email,
              name: user.name || caller.name,
              role: "viewer",
              permissions: "read",
              status: "pending",
              title: "Pending Access Approval",
              createdAt: new Date().toISOString(),
            };
      } else {
        // Returning user: role/status/permissions stay exactly as stored.
        finalUser = { ...existing, name: user.name || existing.name, email: caller.email };
      }
    } else {
      const callerRecord = await getUserInternal(caller.oid);
      if (!callerRecord || callerRecord.role !== "admin") {
        return res.status(403).json({ error: "Admin access required to modify another user's profile" });
      }
      const existingTarget = await getUserInternal(user.id);
      finalUser = { ...existingTarget, ...user };
    }

    await upsertUserInternal(finalUser);
    notifyDataChanged("users", { userId: finalUser.id });
    res.json({ success: true, user: finalUser });
  } catch (error: any) {
    console.error("Error saving user:", error);
    res.status(500).json({ error: "Failed to save user to server-side store" });
  }
});

app.delete("/api/users/row/:id", requireAuth, async (req, res) => {
  try {
    const caller = (req as any).callerIdentity as CallerIdentity;
    const { id } = req.params;

    if (id === caller.oid) {
      return res.status(400).json({ error: "Cannot delete your own active account" });
    }
    const callerRecord = await getUserInternal(caller.oid);
    if (!callerRecord || callerRecord.role !== "admin") {
      return res.status(403).json({ error: "Admin access required" });
    }

    const users = await readUsersFile();
    await writeUsersFile(users.filter((u: any) => u.id !== id));
    notifyDataChanged("users", { deletedId: id });
    res.json({ success: true, deletedId: id });
  } catch (error: any) {
    console.error("Error deleting user:", error);
    res.status(500).json({ error: "Failed to delete user from server-side store" });
  }
});

// Serve health status
app.get("/api/health", (req, res) => {
  res.json({ status: "healthy", time: new Date().toISOString() });
});

// ================= PUSH EXCEL & OTHER API ROUTES =================

app.post(
  "/api/push-excel",
  express.raw({ type: "application/octet-stream", limit: "20mb" }),
  async (req, res) => {
    try {
      const buffer = req.body;
      if (!Buffer.isBuffer(buffer) || buffer.length === 0) {
        return res.status(400).json({ error: "Empty or invalid body payload." });
      }
      if (
        buffer.length < 4 ||
        buffer[0] !== 0x50 ||
        buffer[1] !== 0x4B ||
        buffer[2] !== 0x03 ||
        buffer[3] !== 0x04
      ) {
        return res.status(400).json({ error: "Invalid Excel file format. Must be a valid xlsx (ZIP container)." });
      }

      latestExcelBuffer = buffer;
      latestExcelUpdatedAt = new Date();

      return res.json({ success: true, bytes: buffer.length, updatedAt: latestExcelUpdatedAt });
    } catch (error: any) {
      console.error("Error pushing excel:", error);
      res.status(500).json({ error: "Failed to process pushed Excel file" });
    }
  }
);

app.get("/api/latest-excel", (req, res) => {
  try {
    if (!latestExcelBuffer) {
      return res.status(404).json({ error: "No Excel file has been pushed yet." });
    }
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", 'attachment; filename="latest.xlsx"');
    res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate");
    return res.send(latestExcelBuffer);
  } catch (error: any) {
    console.error("Error serving latest excel:", error);
    res.status(500).json({ error: "Failed to download latest Excel file" });
  }
});

// Configure Vite middleware or serve static dist
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://0.0.0.0:${PORT}`);
  });
}

startServer();
