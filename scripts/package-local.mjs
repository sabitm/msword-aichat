#!/usr/bin/env node
import { cpSync, existsSync, mkdirSync, readdirSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

var root = join(dirname(fileURLToPath(import.meta.url)), "..");
var scriptsDir = join(root, "scripts");
var distDir = join(root, "dist");
var releaseDir = join(root, "release", "msword-aichat-local");
var appDir = join(releaseDir, "app");
var buildExe = process.argv.includes("--exe");

var launcherScripts = [
  "local-server.mjs",
  "install-certificate.mjs",
  "sideload-add-in.mjs",
  "runtime-root.mjs",
];

function run(command, args, cwd) {
  var result = spawnSync(command, args, {
    cwd: cwd || root,
    stdio: "inherit",
    shell: true,
  });
  if (result.status !== 0) {
    process.exit(result.status || 1);
  }
}

var PKG_TARGET = "node22-win-x64";

function buildPkgExe(scriptName, outputName) {
  var scriptPath = join(releaseDir, scriptName);
  var outputPath = join(releaseDir, outputName);

  console.log("Building " + outputName + " (" + PKG_TARGET + ") ...");
  run(
    "npx",
    [
      "--yes",
      "@yao-pkg/pkg",
      scriptPath,
      "--targets",
      PKG_TARGET,
      "--output",
      outputPath,
      "--compress",
      "GZip",
      "--public-packages",
      "*",
    ],
    releaseDir,
  );

  if (!existsSync(outputPath)) {
    console.error("pkg did not produce " + outputName);
    process.exit(1);
  }
}

if (!existsSync(distDir)) {
  console.log("Building production bundle...");
  run("npm", ["run", "build"]);
}

if (existsSync(releaseDir)) {
  rmSync(releaseDir, { recursive: true, force: true });
}
mkdirSync(appDir, { recursive: true });

for (var entry of readdirSync(distDir)) {
  cpSync(join(distDir, entry), join(appDir, entry), { recursive: true });
}

cpSync(join(root, "manifest.xml"), join(releaseDir, "manifest.xml"));

for (var i = 0; i < launcherScripts.length; i += 1) {
  cpSync(join(scriptsDir, launcherScripts[i]), join(releaseDir, launcherScripts[i]));
}

writeFileSync(
  join(releaseDir, "package.json"),
  JSON.stringify(
    {
      name: "msword-aichat-local",
      private: true,
      version: "0.1.0",
      description: "Launcher dependencies for Word AI Chat local distribution",
      dependencies: {
        "office-addin-dev-certs": "^2.0.9",
        "office-addin-dev-settings": "^2.1.5",
      },
      pkg: {
        assets: ["node_modules/office-addin-dev-certs/scripts/**/*"],
      },
    },
    null,
    2,
  ) + "\n",
);

function launcherBatLines(exeName, scriptName) {
  return [
    "if exist " + exeName + " (",
    "  " + exeName,
    "  goto :done",
    ")",
    'where node >nul 2>&1',
    "if errorlevel 1 (",
    "  echo Missing " + exeName + " and Node.js.",
    "  echo Rebuild with: npm run package:local:exe",
    "  pause",
    "  exit /b 1",
    ")",
    "if not exist node_modules (",
    "  echo Installing launcher dependencies...",
    "  call npm install --omit=dev --no-audit --no-fund",
    ")",
    "node " + scriptName,
    ":done",
  ];
}

writeFileSync(
  join(releaseDir, "Install-certificate.bat"),
  [
    "@echo off",
    "setlocal",
    'cd /d "%~dp0"',
    "echo.",
    "echo Word AI Chat - install localhost HTTPS certificate",
    "echo Right-click this file and choose Run as administrator.",
    "echo.",
  ]
    .concat(launcherBatLines("msword-aichat-install-certs.exe", "install-certificate.mjs"))
    .concat(["pause", ""])
    .join("\r\n"),
);

writeFileSync(
  join(releaseDir, "Launch-server.bat"),
  [
    "@echo off",
    "setlocal",
    'cd /d "%~dp0"',
    "if exist msword-aichat-server.exe (",
    "  start \"Word AI Chat Server\" msword-aichat-server.exe",
    "  goto :done",
    ")",
    'where node >nul 2>&1',
    "if errorlevel 1 (",
    "  echo Missing msword-aichat-server.exe and Node.js.",
    "  pause",
    "  exit /b 1",
    ")",
    "if not exist node_modules\\office-addin-dev-certs (",
    "  call npm install --omit=dev --no-audit --no-fund",
    ")",
    "start \"Word AI Chat Server\" node local-server.mjs",
    ":done",
    "echo.",
    "echo Server at https://localhost:3000 - keep the server window open.",
    "echo.",
    "pause",
    "",
  ].join("\r\n"),
);

writeFileSync(
  join(releaseDir, "Sideload-add-in.bat"),
  ["@echo off", "setlocal", 'cd /d "%~dp0"', "echo Registering Word AI Chat with Word...", ""]
    .concat(launcherBatLines("msword-aichat-sideload.exe", "sideload-add-in.mjs"))
    .concat(["pause", ""])
    .join("\r\n"),
);

writeFileSync(
  join(releaseDir, "README-LOCAL.txt"),
  [
    "Word AI Chat — local distribution",
    "",
    "Runs the add-in from https://localhost:3000 on each user's PC (no remote server).",
    "",
    "Recommended package: built with npm run package:local:exe",
    "  - msword-aichat-install-certs.exe  (Node runtime bundled)",
    "  - msword-aichat-sideload.exe         (Node runtime bundled)",
    "  - msword-aichat-server.exe           (Node runtime bundled)",
    "",
    "Requirements:",
    "  - Windows + Microsoft Word 2016 or newer",
    "  - No Node.js install needed when using the .exe launchers",
    "",
    "First-time setup (once per machine):",
    "  1. Right-click Install-certificate.bat -> Run as administrator",
    "  2. Double-click Sideload-add-in.bat",
    "",
    "Every session:",
    "  1. Double-click Launch-server.bat (keep window open)",
    "  2. Word -> Home -> AI Chat",
    "  3. Settings -> configure LLM endpoint and API key",
    "",
    "Build from source:",
    "  npm run package:local       (scripts + Node fallback)",
    "  npm run package:local:exe   (Windows .exe launchers, no Node for users)",
    "",
    "Share the whole msword-aichat-local folder as a ZIP.",
    "",
  ].join("\r\n"),
);

if (buildExe) {
  console.log("Installing launcher dependencies in release folder...");
  run("npm", ["install", "--omit=dev", "--no-audit", "--no-fund", "--prefix", releaseDir]);

  var exeNames = [
    "msword-aichat-install-certs.exe",
    "msword-aichat-sideload.exe",
    "msword-aichat-server.exe",
  ];

  buildPkgExe("install-certificate.mjs", exeNames[0]);
  buildPkgExe("sideload-add-in.mjs", exeNames[1]);
  buildPkgExe("local-server.mjs", exeNames[2]);

  for (var j = 0; j < exeNames.length; j += 1) {
    if (!existsSync(join(releaseDir, exeNames[j]))) {
      console.error("Missing executable after pkg build: " + exeNames[j]);
      process.exit(1);
    }
  }

  console.log("Removing launcher scripts and node_modules (exes are self-contained)...");
  for (var k = 0; k < launcherScripts.length; k += 1) {
    rmSync(join(releaseDir, launcherScripts[k]), { force: true });
  }
  rmSync(join(releaseDir, "package.json"), { force: true });
  rmSync(join(releaseDir, "package-lock.json"), { force: true });
  rmSync(join(releaseDir, "node_modules"), { recursive: true, force: true });
}

console.log("");
console.log("Local distribution written to:");
console.log("  " + releaseDir);
console.log("");
if (buildExe) {
  console.log("Included executables (Node.js runtime bundled inside each):");
  console.log("  - msword-aichat-install-certs.exe");
  console.log("  - msword-aichat-sideload.exe");
  console.log("  - msword-aichat-server.exe");
} else {
  console.log("Build Windows executables (no Node required for end users):");
  console.log("  npm run package:local:exe");
}
console.log("");