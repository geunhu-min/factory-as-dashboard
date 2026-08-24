const http = require("http");
const fs = require("fs");
const path = require("path");

const root = __dirname;
const port = Number(process.env.PORT || 5175);

const dataDir = path.join(root, "data");
const configFile = path.join(dataDir, "config.json");
const defaultConfig = { url: "", monthlyUrl: "", recoveryCsvUrl: "", claimDashboardUrl: "", claimClosingArchiveUrl: "", monthlyClosingUrl: "", dailyPackagingUrl: "", dailyReceiptUrl: "", dailyReceiptStatusUrl: "", dailyRecoveryUrl: "", dailyRecoveryAccumulateUrl: "", dailyPackagingLastDate: "" };

function send(res, status, body, headers = {}) {
  res.writeHead(status, { "Cache-Control": "no-store", ...headers });
  res.end(body);
}

function sendJson(res, status, payload) {
  send(res, status, JSON.stringify(payload), {
    "Content-Type": "application/json; charset=utf-8"
  });
}

async function ensureConfigFile() {
  await fs.promises.mkdir(dataDir, { recursive: true });
  try {
    await fs.promises.access(configFile);
  } catch {
    await fs.promises.writeFile(configFile, JSON.stringify(defaultConfig, null, 2), "utf8");
  }
}

async function readConfig() {
  await ensureConfigFile();
  const raw = await fs.promises.readFile(configFile, "utf8");
  const parsed = JSON.parse(raw || "null") || {};
  return { ...defaultConfig, ...parsed };
}

async function writeConfig(partial) {
  const current = await readConfig();
  const merged = { ...current, ...partial };
  await ensureConfigFile();
  await fs.promises.writeFile(configFile, JSON.stringify(merged, null, 2), "utf8");
  return merged;
}

async function readRequestJson(req) {
  let body = "";
  for await (const chunk of req) {
    body += chunk;
    if (body.length > 1024 * 1024) throw new Error("Request body too large");
  }
  return JSON.parse(body || "{}");
}

// 브라우저(localStorage)가 아니라 이 서버에 URL 3개(정리파일/월현황/회수누적)를
// 저장해서, 같은 서버에 접속하면 어떤 브라우저에서든 같은 값이 보이게 합니다.
async function handleConfigApi(req, res) {
  try {
    if (req.method === "GET") {
      sendJson(res, 200, await readConfig());
      return true;
    }

    if (req.method === "POST") {
      const body = await readRequestJson(req);
      sendJson(res, 200, await writeConfig(body));
      return true;
    }

    return false;
  } catch (err) {
    sendJson(res, 500, { error: err.message });
    return true;
  }
}

async function handleGooglePublishedCsv(req, res) {
  const parsed = new URL(req.url, `http://${req.headers.host}`);
  const sheetUrl = parsed.searchParams.get("url") || "";
  try {
    if (!/^https:\/\/docs\.google\.com\/spreadsheets\//.test(sheetUrl)) {
      throw new Error("올바른 웹에 게시 주소가 아닙니다.");
    }
    const response = await fetch(sheetUrl, {
      headers: { "User-Agent": "Mozilla/5.0" }
    });
    if (!response.ok) throw new Error(`Google 응답 오류 ${response.status}`);
    const text = await response.text();
    send(res, 200, text, { "Content-Type": "text/csv; charset=utf-8" });
  } catch (err) {
    send(res, 500, JSON.stringify({ error: err.message }), {
      "Content-Type": "application/json; charset=utf-8"
    });
  }
}

function contentType(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  return {
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".csv": "text/csv; charset=utf-8"
  }[ext] || "application/octet-stream";
}

function handleStatic(req, res) {
  const parsed = new URL(req.url, `http://${req.headers.host}`);
  const safePath = decodeURIComponent(parsed.pathname === "/" ? "/index.html" : parsed.pathname);
  const filePath = path.resolve(root, `.${safePath}`);

  if (!filePath.startsWith(root)) {
    send(res, 403, "Forbidden");
    return;
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      send(res, 404, "Not found");
      return;
    }
    send(res, 200, data, { "Content-Type": contentType(filePath) });
  });
}

const server = http.createServer((req, res) => {
  if (req.url.startsWith("/api/google-published-csv")) {
    handleGooglePublishedCsv(req, res);
    return;
  }
  if (req.url.startsWith("/api/config")) {
    handleConfigApi(req, res).then((handled) => {
      if (!handled) send(res, 405, "Method not allowed");
    });
    return;
  }
  handleStatic(req, res);
});

server.on("error", (err) => {
  console.error(err.message);
});

server.listen(port, () => {
  console.log(`공장 AS 추가건 수정 샘플: http://localhost:${port}`);
});
