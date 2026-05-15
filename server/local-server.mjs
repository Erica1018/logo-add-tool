import { spawn } from "node:child_process";
import { createServer } from "node:https";
import {
  createReadStream,
  existsSync,
  mkdtempSync,
  readFileSync,
  rmSync,
  statSync,
  writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, extname, join, normalize, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = dirname(dirname(fileURLToPath(import.meta.url)));
const distDir = join(rootDir, "dist");
const scriptsDir = join(rootDir, "scripts");
const certDir = join(homedir(), ".office-addin-dev-certs");
const port = Number(process.env.LOGO_ADD_TOOL_PORT ?? 3001);

const certPath = join(certDir, "localhost.crt");
const keyPath = join(certDir, "localhost.key");

if (!existsSync(certPath) || !existsSync(keyPath)) {
  console.error("Missing trusted localhost certificate.");
  console.error("Run install-windows.ps1 or install-mac.command first.");
  process.exit(1);
}

if (!existsSync(distDir)) {
  console.error("Missing dist folder. Use the packaged release, or run npm run build first.");
  process.exit(1);
}

const server = createServer(
  {
    cert: readFileSync(certPath),
    key: readFileSync(keyPath),
  },
  async (req, res) => {
    try {
      if (req.url === "/api/powerpoint/apply-all") {
        await handlePowerPointComRequest(req, res, "apply-logo-com.ps1");
        return;
      }

      if (req.url === "/api/powerpoint/read-selected-placement") {
        await handlePowerPointComRequest(req, res, "read-selected-placement-com.ps1");
        return;
      }

      serveStatic(req, res);
    } catch (error) {
      sendJson(res, 500, { error: error instanceof Error ? error.message : String(error) });
    }
  },
);

server.listen(port, "localhost", () => {
  console.log(`Logo Add Tool local server: https://localhost:${port}/`);
  console.log("Keep this window open while using the PowerPoint add-in.");
});

async function handlePowerPointComRequest(req, res, scriptName) {
  if (req.method !== "POST") {
    sendJson(res, 405, { error: "Method not allowed." });
    return;
  }

  if (process.platform !== "win32") {
    sendJson(res, 400, { error: "PowerPoint COM helper is only available on Windows." });
    return;
  }

  const body = await readRequestBody(req);
  const tempDir = mkdtempSync(join(tmpdir(), "logo-add-tool-"));
  const payloadPath = join(tempDir, "payload.json");
  writeFileSync(payloadPath, body, "utf8");

  const result = await runPowerShellScript(join(scriptsDir, scriptName), payloadPath).finally(() => {
    rmSync(tempDir, { recursive: true, force: true });
  });

  if (result.code !== 0) {
    sendJson(res, 500, { error: result.stderr.trim() || result.stdout.trim() || "PowerPoint COM failed." });
    return;
  }

  res.writeHead(200, { "Content-Type": "application/json; charset=utf-8" });
  res.end(result.stdout.trim() || JSON.stringify({ ok: true }));
}

function runPowerShellScript(scriptPath, payloadPath) {
  return new Promise((resolveResult) => {
    const child = spawn(
      "powershell.exe",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptPath, "-PayloadPath", payloadPath],
      { cwd: rootDir, windowsHide: true },
    );

    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => {
      stdout += String(chunk);
    });
    child.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });
    child.on("close", (code) => resolveResult({ code, stdout, stderr }));
  });
}

function readRequestBody(req) {
  return new Promise((resolveBody, rejectBody) => {
    let body = "";
    req.setEncoding("utf8");
    req.on("data", (chunk) => {
      body += chunk;
    });
    req.on("end", () => resolveBody(body));
    req.on("error", rejectBody);
  });
}

function serveStatic(req, res) {
  const url = new URL(req.url ?? "/", `https://localhost:${port}`);
  const requestedPath = url.pathname === "/" ? "/index.html" : decodeURIComponent(url.pathname);
  const normalizedPath = normalize(requestedPath).replace(/^(\.\.[/\\])+/, "");
  const absolutePath = resolve(distDir, `.${normalizedPath}`);

  if (!absolutePath.startsWith(resolve(distDir))) {
    sendText(res, 403, "Forbidden");
    return;
  }

  if (!existsSync(absolutePath) || !statSync(absolutePath).isFile()) {
    sendText(res, 404, "Not found");
    return;
  }

  res.writeHead(200, {
    "Content-Type": contentTypeFor(absolutePath),
    "Cache-Control": "no-cache",
  });
  createReadStream(absolutePath).pipe(res);
}

function contentTypeFor(filePath) {
  const extension = extname(filePath).toLowerCase();
  const types = {
    ".css": "text/css; charset=utf-8",
    ".html": "text/html; charset=utf-8",
    ".js": "text/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".svg": "image/svg+xml",
  };
  return types[extension] ?? "application/octet-stream";
}

function sendJson(res, statusCode, value) {
  res.writeHead(statusCode, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(value));
}

function sendText(res, statusCode, value) {
  res.writeHead(statusCode, { "Content-Type": "text/plain; charset=utf-8" });
  res.end(value);
}
