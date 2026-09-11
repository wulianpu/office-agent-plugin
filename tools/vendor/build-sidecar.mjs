/**
 * Build the vendored XLSX Rust sidecar (§30): cargo build --release over
 * vendor/genoffice/apps/sheets/native/xlsx-engine. Locates cargo on PATH or
 * in the common relocated rustup install roots.
 */

import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const root = join(dirname(fileURLToPath(import.meta.url)), "..", "..");
const crateDir = join(root, "vendor", "genoffice", "apps", "sheets", "native", "xlsx-engine");

const candidates = [
  "cargo",
  "C:\\Library\\Rust\\cargo\\bin\\cargo.exe",
  join(process.env.USERPROFILE ?? "", ".cargo", "bin", "cargo.exe"),
];

const cargo = candidates.find((candidate) => {
  if (candidate.includes("\\") && !existsSync(candidate)) return false;
  const probe = spawnSync(candidate, ["--version"], { windowsHide: true, encoding: "utf8" });
  return probe.status === 0;
});

if (!cargo) {
  console.error("cargo not found on PATH or in known install roots — install Rust to build the sidecar");
  process.exit(2);
}

console.log(`building xlsx-sidecar with ${cargo} (release)…`);
const build = spawnSync(cargo, ["build", "--release"], {
  cwd: crateDir,
  stdio: "inherit",
  windowsHide: true,
  env: {
    ...process.env,
    RUSTUP_HOME: process.env.RUSTUP_HOME ?? "C:\\Library\\Rust\\rustup",
    CARGO_HOME: process.env.CARGO_HOME ?? "C:\\Library\\Rust\\cargo"
  }
});
process.exit(build.status ?? 1);
