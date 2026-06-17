#!/usr/bin/env node
"use strict";

const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawn, spawnSync } = require("child_process");

const packageRoot = path.resolve(__dirname, "..");
const packageJson = require(path.join(packageRoot, "package.json"));
const markerName = `.installed-${packageJson.version}`;
const markerPayload = JSON.stringify(
  {
    packageRoot,
    version: packageJson.version,
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
  }
  candidates.push({ command: "python", args: [] });
  candidates.push({ command: "python3", args: [] });
  return candidates;
}

function findPython() {
  const probe = "import sys; print(f'{sys.version_info[0]}.{sys.version_info[1]}')";
  const errors = [];
  for (const candidate of pythonCandidates()) {
    const result = spawnSync(candidate.command, [...candidate.args, "-c", probe], {
      encoding: "utf8",
      windowsHide: true,
    });
    if (result.status !== 0) {
      const detail = (result.stderr || result.stdout || result.error?.message || "").trim();
      errors.push(`${candidate.command} ${candidate.args.join(" ")} ${detail}`.trim());
      continue;
    }
    const version = (result.stdout || "").trim().split(".").map((part) => Number(part));
    if (version[0] > 3 || (version[0] === 3 && version[1] >= 11)) {
      return candidate;
    }
    errors.push(`${candidate.command} ${candidate.args.join(" ")} found Python ${version.join(".")}, need >=3.11`.trim());
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

function ensureVenv(python) {
  const root = cacheRoot();
  const venvDir = path.join(root, `venv-${packageJson.version}`);
  const marker = path.join(venvDir, markerName);
  const py = venvPython(venvDir);

  fs.mkdirSync(root, { recursive: true });

  if (!fs.existsSync(py)) {
    log(`creating Python environment in ${venvDir}`);
    const created = run(python.command, [...python.args, "-m", "venv", venvDir]);
    if (created.status !== 0) {
      throw new Error(`Failed to create Python venv at ${venvDir}.`);
    }
  }

  const installedMarker = fs.existsSync(marker) ? fs.readFileSync(marker, "utf8") : "";
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
    const python = findPython();
    const py = ensureVenv(python);
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
