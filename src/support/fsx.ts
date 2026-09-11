import { createHash } from "node:crypto";
import {
  copyFile,
  mkdir,
  open,
  readFile,
  rename,
  rm,
  stat,
  writeFile
} from "node:fs/promises";
import type { FileFingerprint } from "../contracts/artifact.js";
import { OfficeRuntimeError } from "../contracts/document.js";

export async function pathExists(path: string): Promise<boolean> {
  try {
    await stat(path);
    return true;
  } catch {
    return false;
  }
}

export async function ensureDir(path: string): Promise<void> {
  await mkdir(path, { recursive: true });
}

/** Fast-path identity: stat fingerprint, no hashing (§17). */
export async function fileFingerprint(path: string): Promise<FileFingerprint> {
  const s = await stat(path).catch((error) => {
    throw new OfficeRuntimeError(
      "artifact-missing",
      `cannot stat artifact source: ${path} (${String(error)})`
    );
  });
  const fileId =
    s.ino !== undefined && s.ino !== 0 ? `${s.dev.toString(36)}-${s.ino.toString(36)}` : undefined;
  return {
    size: BigInt(s.size),
    mtimeNs: BigInt(Math.round(s.mtimeMs * 1_000_000)),
    fileId
  };
}

export function fingerprintKey(fp: FileFingerprint): string {
  return `${fp.size.toString()}:${fp.mtimeNs.toString()}:${fp.fileId ?? "-"}`;
}

export async function sha256File(path: string): Promise<string> {
  const hash = createHash("sha256");
  const handle = await open(path, "r");
  try {
    const buffer = Buffer.alloc(1024 * 1024);
    let position = 0n;
    for (;;) {
      const { bytesRead } = await handle.read(buffer, 0, buffer.length, position);
      if (bytesRead === 0) break;
      hash.update(buffer.subarray(0, bytesRead));
      position += BigInt(bytesRead);
    }
  } finally {
    await handle.close();
  }
  return hash.digest("hex");
}

export function sha256Buffer(data: Uint8Array): string {
  return createHash("sha256").update(data).digest("hex");
}

/**
 * Candidate clone (§66): prefer copy semantics. Hard links are forbidden as an
 * isolation mechanism; fs.copyFile gives independent bytes on all platforms.
 */
export async function cloneFile(source: string, target: string): Promise<void> {
  await copyFile(source, target);
}

/** fsync a file path by briefly opening it (used by the atomic committer). */
export async function fsyncFile(path: string): Promise<void> {
  const handle = await open(path, "r+");
  try {
    await handle.sync();
  } finally {
    await handle.close();
  }
}

export async function fsyncParentDir(path: string): Promise<void> {
  // Directory fsync is not expressible on Windows; POSIX best-effort.
  if (process.platform === "win32") return;
  const parent = path.slice(0, path.lastIndexOf(path.includes("\\") ? "\\" : "/")) || "/";
  const handle = await open(parent, "r").catch(() => undefined);
  if (!handle) return;
  try {
    await handle.sync();
  } catch {
    // Some filesystems refuse directory fsync; not fatal.
  } finally {
    await handle.close();
  }
}

export async function writeFileAtomicBuffered(
  target: string,
  data: Uint8Array
): Promise<void> {
  const temp = `${target}.tmp-${Date.now().toString(36)}`;
  await writeFile(temp, data);
  await fsyncFile(temp);
  await rename(temp, target);
}

export async function removeQuiet(path: string): Promise<void> {
  // Windows: AV/indexer/just-closed residents can hold handles briefly.
  const backoffs = [0, 100, 250, 500, 1000, 2000];
  let lastError: unknown;
  for (const wait of backoffs) {
    if (wait > 0) await new Promise((resolve) => setTimeout(resolve, wait));
    try {
      await rm(path, { force: true, recursive: false });
      return;
    } catch (error) {
      lastError = error;
      const code = (error as NodeJS.ErrnoException).code;
      if (code !== "EBUSY" && code !== "EPERM") throw error;
    }
  }
  throw lastError;
}

export async function readFileBuffer(path: string): Promise<Buffer> {
  return readFile(path);
}
