import { readFileSync, writeFileSync } from "node:fs";
import { dirname, isAbsolute, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = dirname(dirname(fileURLToPath(import.meta.url)));
const rawUrl = readUrl();
const taskpaneOrigin = normalizeOrigin(rawUrl);
const outputPath = readOutputPath();

const templatePath = join(rootDir, "manifest.template.xml");
const template = readFileSync(templatePath, "utf8");

if (template.includes("__TASKPANE_ORIGIN__") === false) {
  throw new Error("manifest.template.xml is missing __TASKPANE_ORIGIN__.");
}

writeFileSync(outputPath, template.replaceAll("__TASKPANE_ORIGIN__", taskpaneOrigin), "utf8");
console.log(`Wrote ${outputPath} for ${taskpaneOrigin}`);

function readUrl() {
  const cliUrl = process.argv.find((arg) => arg.startsWith("--url="))?.slice("--url=".length);
  const envUrl = process.env.LOGO_ADD_TOOL_URL;
  const url = cliUrl || envUrl;

  if (!url) {
    throw new Error("Set LOGO_ADD_TOOL_URL or pass --url=https://your-domain/logo-add-tool.");
  }

  return url;
}

function readOutputPath() {
  const cliPath = process.argv.find((arg) => arg.startsWith("--out="))?.slice("--out=".length);
  if (!cliPath) {
    return join(rootDir, "manifest.xml");
  }

  return isAbsolute(cliPath) ? cliPath : resolve(process.cwd(), cliPath);
}

function normalizeOrigin(value) {
  const normalized = value.trim().replace(/\/+$/, "");
  const url = new URL(normalized);

  if (url.protocol !== "https:") {
    throw new Error("PowerPoint Office add-ins require an HTTPS task pane URL.");
  }

  const placeholderHost =
    url.hostname.endsWith("example.com") ||
    url.hostname.includes("your-domain") ||
    url.hostname.includes("yourdomain") ||
    url.hostname.endsWith(".invalid") ||
    url.hostname.endsWith(".test");

  if (placeholderHost && process.env.ALLOW_PLACEHOLDER !== "1") {
    throw new Error("Use a real HTTPS host, not a placeholder domain.");
  }

  return normalized;
}
