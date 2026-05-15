import { defineConfig, type Plugin, type ViteDevServer } from "vite";
import react from "@vitejs/plugin-react";
import { spawn } from "node:child_process";
import type { IncomingMessage, ServerResponse } from "node:http";
import { readFileSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const officeCertDir = join(homedir(), ".office-addin-dev-certs");
const projectDir = dirname(fileURLToPath(import.meta.url));
const https = {
  cert: readFileSync(join(officeCertDir, "localhost.crt")),
  key: readFileSync(join(officeCertDir, "localhost.key")),
};

export default defineConfig({
  plugins: [react(), powerPointComHelper()],
  server: {
    host: "localhost",
    port: 3001,
    strictPort: true,
    https,
  },
  preview: {
    host: "localhost",
    port: 3001,
    strictPort: true,
    https,
  },
});

function powerPointComHelper(): Plugin {
  return {
    name: "powerpoint-com-helper",
    configureServer(server: ViteDevServer) {
      server.middlewares.use("/api/powerpoint/apply-all", (req: IncomingMessage, res: ServerResponse, next) => {
        handlePowerPointComRequest(req, res, next, "apply-logo-com.ps1");
      });

      server.middlewares.use("/api/powerpoint/read-selected-placement", (req: IncomingMessage, res: ServerResponse, next) => {
        handlePowerPointComRequest(req, res, next, "read-selected-placement-com.ps1");
      });
    },
  };
}

function handlePowerPointComRequest(
  req: IncomingMessage,
  res: ServerResponse,
  next: () => void,
  scriptName: string,
): void {
        if (req.method !== "POST") {
          next();
          return;
        }

        let body = "";
        req.setEncoding("utf8");
        req.on("data", (chunk: string) => {
          body += chunk;
        });
        req.on("end", () => {
          const tempDir = mkdtempSync(join(tmpdir(), "brand-logo-stamp-"));
          const payloadPath = join(tempDir, "payload.json");
          writeFileSync(payloadPath, body, "utf8");

          const child = spawn(
            "powershell.exe",
            [
              "-NoProfile",
              "-ExecutionPolicy",
              "Bypass",
              "-File",
              join(projectDir, "scripts", scriptName),
              "-PayloadPath",
              payloadPath,
            ],
            { cwd: projectDir, windowsHide: true },
          );

          let stdout = "";
          let stderr = "";
          child.stdout.on("data", (chunk) => {
            stdout += String(chunk);
          });
          child.stderr.on("data", (chunk) => {
            stderr += String(chunk);
          });
          child.on("close", (code) => {
            rmSync(tempDir, { recursive: true, force: true });
            res.setHeader("Content-Type", "application/json; charset=utf-8");

            if (code !== 0) {
              res.statusCode = 500;
              res.end(JSON.stringify({ error: stderr.trim() || stdout.trim() || "PowerPoint COM 执行失败。" }));
              return;
            }

            res.end(stdout.trim() || JSON.stringify({ ok: true }));
          });
        });
}
