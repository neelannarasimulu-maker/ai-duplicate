const http = require("node:http");
const fs = require("node:fs/promises");
const path = require("node:path");

const port = Number(process.env.API_PORT || 8787);
const dataDir = path.join(process.cwd(), "data");
const dbFile = path.join(dataDir, "workbench-db.json");

const defaultDb = {
  tasks: [],
  outputs: [],
  meta: {},
};

async function readDb() {
  await fs.mkdir(dataDir, { recursive: true });
  try {
    const content = await fs.readFile(dbFile, "utf8");
    return { ...defaultDb, ...JSON.parse(content) };
  } catch (error) {
    if (error.code !== "ENOENT") throw error;
    await writeDb(defaultDb);
    return defaultDb;
  }
}

async function writeDb(db) {
  await fs.mkdir(dataDir, { recursive: true });
  await fs.writeFile(dbFile, JSON.stringify(db, null, 2));
}

async function readBody(request) {
  const chunks = [];
  for await (const chunk of request) chunks.push(chunk);
  const raw = Buffer.concat(chunks).toString("utf8");
  return raw ? JSON.parse(raw) : {};
}

function send(response, status, payload) {
  response.writeHead(status, {
    "Access-Control-Allow-Origin": "*",
    "Access-Control-Allow-Methods": "GET,POST,PUT,DELETE,OPTIONS",
    "Access-Control-Allow-Headers": "Content-Type",
    "Content-Type": "application/json",
  });
  response.end(JSON.stringify(payload));
}

const server = http.createServer(async (request, response) => {
  try {
    if (request.method === "OPTIONS") return send(response, 204, {});

    const url = new URL(request.url || "/", `http://${request.headers.host}`);
    const db = await readDb();

    if (url.pathname === "/api/health") return send(response, 200, { ok: true });
    if (url.pathname === "/api/tasks" && request.method === "GET") return send(response, 200, db.tasks);
    if (url.pathname === "/api/outputs" && request.method === "GET") return send(response, 200, db.outputs);

    if (url.pathname === "/api/tasks" && request.method === "PUT") {
      const task = await readBody(request);
      db.tasks = [task, ...db.tasks.filter((item) => item.id !== task.id)];
      await writeDb(db);
      return send(response, 200, task);
    }

    if (url.pathname.startsWith("/api/tasks/") && request.method === "DELETE") {
      const id = decodeURIComponent(url.pathname.replace("/api/tasks/", ""));
      db.tasks = db.tasks.filter((item) => item.id !== id);
      await writeDb(db);
      return send(response, 200, { ok: true });
    }

    if (url.pathname === "/api/outputs" && request.method === "PUT") {
      const output = await readBody(request);
      db.outputs = [output, ...db.outputs.filter((item) => item.id !== output.id)];
      await writeDb(db);
      return send(response, 200, output);
    }

    if (url.pathname.startsWith("/api/meta/")) {
      const key = decodeURIComponent(url.pathname.replace("/api/meta/", ""));
      if (request.method === "GET") return send(response, 200, { key, value: db.meta[key] });
      if (request.method === "PUT") {
        const body = await readBody(request);
        db.meta[key] = body.value;
        await writeDb(db);
        return send(response, 200, { key, value: db.meta[key] });
      }
    }

    return send(response, 404, { error: "Not found" });
  } catch (error) {
    return send(response, 500, { error: error.message || "Server error" });
  }
});

server.listen(port, "0.0.0.0", () => {
  console.log(`AI Workbench API listening on http://0.0.0.0:${port}`);
});
