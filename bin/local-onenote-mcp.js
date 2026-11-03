#!/usr/bin/env node
"use strict";

const crypto = require("crypto");
const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawn, spawnSync } = require("child_process");

const packageRoot = fs.realpathSync(path.resolve(__dirname, ".."));
const packageJson = require(path.join(packageRoot, "package.json"));
const markerName = `.installed-${packageJson.version}`;

function sourceFingerprint() {
  const hash = crypto.createHash("sha256");
  const sourceDir = path.join(packageRoot, "src", "local_onenote_mcp");
  const files = [path.join(packageRoot, "pyproject.toml")];
  for (const name of fs.readdirSync(sourceDir).sort()) {
    if (name.endsWith(".py")) {
      files.push(path.join(sourceDir, name));
    }
  }
  for (const file of files) {
    hash.update(path.relative(packageRoot, file));
    hash.update(fs.readFileSync(file));
  }
  return hash.digest("hex");
}

const markerPayload = JSON.stringify(
  {
    packageRoot,
    version: packageJson.version,
    source: sourceFingerprint(),
  },
  null,
  2
);

function log(message) {
  process.stderr.write(`[local-onenote-mcp] ${message}\n`);
}

function run(command, args, options = {}) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    windowsHide: true,
    ...options,
  });
  if (result.stdout) {
    process.stderr.write(result.stdout);
  }
  if (result.stderr) {
    process.stderr.write(result.stderr);
  }
  return result;
}

function pythonCandidates() {
  const candidates = [];
  if (process.env.LOCAL_ONENOTE_MCP_PYTHON) {
    candidates.push({ command: process.env.LOCAL_ONENOTE_MCP_PYTHON, args: [] });
  }
  if (process.platform === "win32") {
    candidates.push({ command: "py", args: ["-3.11"] });
    candidates.push({ command: "py", args: ["-3"] });
    const found = spawnSync("where.exe", ["python"], { encoding: "utf8", windowsHide: true });
    for (const line of (found.stdout || "").split(/\r?\n/)) {
      const exe = line.trim();
      if (exe && !exe.toLowerCase().includes("\\windowsapps\\")) {
        candidates.push({ command: exe, args: [] });
      }
    }
  }
  candidates.push({ command: "python", args: [] });
  candidates.push({ command: "python3", args: [] });
  return candidates;
}

const PYTHON_PROBE = "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}'); print(sys.executable)";

function probePython(candidate, timeout) {
  const result = spawnSync(candidate.command, [...(candidate.args || []), "-c", PYTHON_PROBE], {
    encoding: "utf8",
    windowsHide: true,
    timeout,
  });
  if (result.status !== 0) {
    const detail = (result.stderr || result.stdout || result.error?.message || "").trim();
    return { error: `${candidate.command} ${(candidate.args || []).join(" ")} ${detail}`.trim() };
  }
  const [versionLine = "", executable = ""] = (result.stdout || "").trim().split(/\r?\n/);
  const version = versionLine.trim().split(".").map((part) => Number(part));
  if (!(version[0] > 3 || (version[0] === 3 && version[1] >= 11))) {
    return { error: `${candidate.command} found Python ${versionLine.trim()}, need >=3.11` };
  }
  // Cache the absolute interpreter path so a later PATH change (or the
  // Microsoft Store "python" stub shadowing it) does not break startup.
  return { python: executable.trim() ? { command: executable.trim(), args: [] } : candidate };
}

function findPython() {
  const root = cacheRoot();
  const cacheFile = path.join(root, "detected-python.json");

  if (fs.existsSync(cacheFile)) {
    try {
      const cached = JSON.parse(fs.readFileSync(cacheFile, "utf8"));
      if (cached && cached.command) {
        const probed = probePython(cached, 5000);
        if (probed.python) {
          return probed.python;
        }
      }
    } catch (e) {
      // Ignore cache read errors
    }
  }

  const errors = [];
  for (const candidate of pythonCandidates()) {
    const probed = probePython(candidate);
    if (probed.error) {
      errors.push(probed.error);
      continue;
    }
    try {
      fs.mkdirSync(root, { recursive: true });
      fs.writeFileSync(cacheFile, JSON.stringify(probed.python, null, 2), "utf8");
    } catch (e) {
      // Ignore cache write errors
    }
    return probed.python;
  }
  throw new Error(
    "Python 3.11+ was not found. Install Python, or set LOCAL_ONENOTE_MCP_PYTHON to a Python executable.\n" +
      errors.map((line) => `  - ${line}`).join("\n")
  );
}

function cacheRoot() {
  if (process.env.LOCAL_ONENOTE_MCP_NPM_CACHE) {
    return path.resolve(process.env.LOCAL_ONENOTE_MCP_NPM_CACHE);
  }
  if (process.platform === "win32" && process.env.LOCALAPPDATA) {
    return path.join(process.env.LOCALAPPDATA, "local-onenote-mcp", "npm-runner");
  }
  const base = process.env.XDG_CACHE_HOME || path.join(os.homedir(), ".cache");
  return path.join(base, "local-onenote-mcp", "npm-runner");
}

function venvPython(venvDir) {
  return process.platform === "win32"
    ? path.join(venvDir, "Scripts", "python.exe")
    : path.join(venvDir, "bin", "python");
}

function ensureVenv() {
  const root = cacheRoot();
  const venvDir = path.join(root, `venv-${packageJson.version}`);
  const marker = path.join(venvDir, markerName);
  const py = venvPython(venvDir);

  const installedMarker = fs.existsSync(marker) ? fs.readFileSync(marker, "utf8") : "";
  if (fs.existsSync(py) && installedMarker === markerPayload) {
    return py;
  }

  fs.mkdirSync(root, { recursive: true });

  if (!fs.existsSync(py)) {
    const python = findPython();
    log(`creating Python environment in ${venvDir}`);
    const created = run(python.command, [...python.args, "-m", "venv", venvDir]);
    if (created.status !== 0) {
      throw new Error(`Failed to create Python venv at ${venvDir}.`);
    }
  }

  if (installedMarker !== markerPayload) {
    log("installing Python package dependencies");
    const installed = run(py, ["-m", "pip", "install", "--upgrade", packageRoot]);
    if (installed.status !== 0) {
      throw new Error("Failed to install local-onenote-mcp into the cached Python environment.");
    }
    fs.writeFileSync(marker, markerPayload, "utf8");
  }

  return py;
}

function main() {
  try {
    const py = ensureVenv();
    const child = spawn(py, ["-m", "local_onenote_mcp.server"], {
      stdio: "inherit",
      windowsHide: true,
      env: process.env,
    });

    for (const signal of ["SIGINT", "SIGTERM"]) {
      process.on(signal, () => {
        if (!child.killed) {
          child.kill(signal);
        }
      });
    }

    child.on("exit", (code, signal) => {
      if (signal) {
        process.kill(process.pid, signal);
      } else {
        process.exit(code ?? 0);
      }
    });
  } catch (error) {
    process.stderr.write(`[local-onenote-mcp] ${error.message || error}\n`);
    process.exit(1);
  }
}

main();
