import { spawnSync } from "node:child_process";
import { cpSync, existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const rootDir = dirname(dirname(fileURLToPath(import.meta.url)));
const version = readFileSync(join(rootDir, "VERSION"), "utf8").trim();
const releaseRoot = join(rootDir, "release");
const taskpaneUrl = readTaskpaneUrl();
const placeholderPackage = isPlaceholderUrl(taskpaneUrl);
const releaseName = `logo-add-tool-mac-portable-${version}${placeholderPackage ? "-developer-template-do-not-send" : ""}`;
const stagingDir = join(releaseRoot, releaseName);
const zipPath = join(releaseRoot, `${releaseName}.zip`);

run("npm", ["run", "build"]);

if (!existsSync(join(rootDir, "dist", "index.html"))) {
  throw new Error("dist/index.html was not generated.");
}

rmSync(stagingDir, { recursive: true, force: true });
rmSync(zipPath, { force: true });
mkdirSync(stagingDir, { recursive: true });

copy("dist", "dist");
copy("docs", "docs");
copy("CHANGELOG.md", "CHANGELOG.md");
copy("VERSION", "VERSION");
copy("install-mac.command", "install-mac.command");
copy("uninstall-mac.command", "uninstall-mac.command");
if (existsSync(join(rootDir, "LICENSE"))) {
  copy("LICENSE", "LICENSE");
}

run("node", ["tools/render-manifest.mjs", `--url=${taskpaneUrl}`, `--out=${join(stagingDir, "manifest.xml")}`], {
  env: placeholderPackage ? { ALLOW_PLACEHOLDER: "1" } : {},
});

writeDeliveryNote();
if (placeholderPackage) {
  writeDeveloperTemplateReadme();
  writeBlockedInstaller();
} else {
  writeEndUserReadme();
}

zipRelease();
console.log(`Packaged ${zipPath}`);

function readTaskpaneUrl() {
  const value = process.env.LOGO_ADD_TOOL_URL;
  if (!value) {
    throw new Error("Set LOGO_ADD_TOOL_URL before packaging the Mac no-install release.");
  }
  return value.trim().replace(/\/+$/, "");
}

function copy(from, to) {
  cpSync(join(rootDir, from), join(stagingDir, to), { recursive: true });
}

function writeDeliveryNote() {
  writeFileSync(
    join(stagingDir, "DELIVERY.txt"),
    [
      "Logo 添加工具 Mac 免安装版",
      "",
      placeholderPackage
        ? "注意：这是开发模板包，manifest 里仍然是示例地址，不能直接发给普通用户使用。"
        : "这是最终用户包，可以直接发给 Mac 用户。",
      "",
      "普通用户只需要：",
      "1. 解压这个文件夹。",
      "2. 双击 install-mac.command。",
      "3. 重启 PowerPoint，在“开始 > 加载项”里打开“Logo 添加工具”。",
      "",
      `侧边栏托管地址：${taskpaneUrl}`,
    ].join("\n"),
    "utf8",
  );
}

function writeEndUserReadme() {
  writeFileSync(
    join(stagingDir, "README.md"),
    [
      "# Logo 添加工具 Mac 使用指南",
      "",
      "这个版本用于 PowerPoint for Mac。它不需要安装 Node.js，不需要启动本地服务，也不需要管理员密码。",
      "",
      "## 安装",
      "",
      "1. 解压压缩包。",
      "2. 双击 `install-mac.command`。",
      "3. 关闭并重新打开 PowerPoint。",
      "4. 打开一个 PPT。",
      "5. 进入 `开始 > 加载项`，选择 `Logo 添加工具`。",
      "",
      "如果 macOS 提示无法打开脚本：右键点 `install-mac.command`，选择“打开”。",
      "",
      "## 使用",
      "",
      "1. 点击 `导入 Logo`。",
      "2. 选择一个或多个 PNG、JPEG 或 WebP 文件。",
      "3. 选中 Logo，点击 `插入当前页`。",
      "4. 在 PowerPoint 页面里拖拽、缩放 Logo。",
      "5. 点击 `使用当前选中 Logo 的位置`。",
      "6. 点击 `应用全部`。",
      "",
      "## 卸载",
      "",
      "双击 `uninstall-mac.command`，然后重启 PowerPoint。",
      "",
      "## 说明",
      "",
      "这个安装动作只是把插件 manifest 放到当前用户自己的 PowerPoint 加载项目录，不会安装系统软件。",
    ].join("\n"),
    "utf8",
  );
}

function writeDeveloperTemplateReadme() {
  writeFileSync(
    join(stagingDir, "README.md"),
    [
      "# Logo 添加工具 Mac 开发模板包",
      "",
      "这个包还不能给普通 Mac 用户使用。",
      "",
      "原因：`manifest.xml` 里仍然是示例地址，PowerPoint 无法从这个地址加载侧边栏。",
      "",
      "要生成最终用户包，需要先把 `dist/` 上传到一个真实 HTTPS 地址，然后重新打包：",
      "",
      "```bash",
      "LOGO_ADD_TOOL_URL=https://你的正式域名/logo-add-tool npm run package:mac",
      "```",
      "",
      "Windows PowerShell：",
      "",
      "```powershell",
      "$env:LOGO_ADD_TOOL_URL=\"https://你的正式域名/logo-add-tool\"",
      "npm run package:mac",
      "```",
      "",
      "生成出来的不带 `developer-template-do-not-send` 字样的 zip，才是可以发给 Mac 用户的包。",
    ].join("\n"),
    "utf8",
  );
}

function writeBlockedInstaller() {
  writeFileSync(
    join(stagingDir, "install-mac.command"),
    [
      "#!/bin/zsh",
      "echo \"这个是开发模板包，不能直接安装。\"",
      "echo \"\"",
      "echo \"原因：manifest.xml 里还是示例 HTTPS 地址，PowerPoint 无法加载侧边栏。\"",
      "echo \"\"",
      "echo \"请先把 dist/ 上传到正式 HTTPS 地址，然后由交付方重新生成最终用户包。\"",
      "echo \"最终用户包的文件名不应该包含 developer-template-do-not-send。\"",
      "echo \"\"",
      "read -k 1 \"?按任意键关闭...\"",
    ].join("\n"),
    "utf8",
  );
}

function run(command, args, options = {}) {
  const useWindowsNpm = process.platform === "win32" && command === "npm";
  const executable = useWindowsNpm ? "cmd.exe" : command;
  const finalArgs = useWindowsNpm ? ["/d", "/s", "/c", "npm", ...args] : args;
  const result = spawnSync(executable, finalArgs, {
    cwd: options.cwd ?? rootDir,
    stdio: "inherit",
    shell: options.shell ?? false,
    env: { ...process.env, ...(options.env ?? {}) },
  });

  if (result.status !== 0) {
    const reason = result.error ? `: ${result.error.message}` : "";
    throw new Error(`${command} ${args.join(" ")} failed with code ${result.status}${reason}.`);
  }
}

function zipRelease() {
  if (process.platform === "win32") {
    const command = [
      "$ErrorActionPreference='Stop'",
      `$source = '${escapePowerShellPath(join(stagingDir, "*"))}'`,
      `$dest = '${escapePowerShellPath(zipPath)}'`,
      "Compress-Archive -Path $source -DestinationPath $dest -Force",
    ].join("; ");
    run("powershell.exe", ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command], { shell: false });
    return;
  }

  if (process.platform === "darwin" && existsSync("/usr/bin/ditto")) {
    run("/usr/bin/ditto", ["-c", "-k", "--sequesterRsrc", "--keepParent", stagingDir, zipPath]);
    return;
  }

  run("zip", ["-r", zipPath, releaseName], { cwd: releaseRoot });
}

function isPlaceholderUrl(value) {
  const url = new URL(value);
  return (
    url.hostname.endsWith("example.com") ||
    url.hostname.includes("your-domain") ||
    url.hostname.includes("yourdomain") ||
    url.hostname.endsWith(".invalid") ||
    url.hostname.endsWith(".test")
  );
}

function escapePowerShellPath(value) {
  return value.replaceAll("'", "''");
}
