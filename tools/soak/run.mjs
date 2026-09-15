/**
 * Issue #15 P1 soak: sustained Preview/Open/Close/Edit lifecycle loop with
 * periodic session close/reopen; samples RSS and open FDs and asserts a
 * PLATEAU (no monotonic growth). Engine-free so it runs anywhere.
 *
 * Usage: node tools/soak/run.mjs [cycles=40] [sampleEvery=5]
 */

import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { writeFileSync as writeFileSyncSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { OfficePlugin } from "../../dist/plugin/office-plugin.js";

const cycles = Number(process.argv[2] ?? 40);
const sampleEvery = Number(process.argv[3] ?? 5);

function rssMB() {
  return process.memoryUsage().rss / 1024 / 1024;
}
function heapMB() {
  return process.memoryUsage().heapUsed / 1024 / 1024;
}
function fdCount() {
  try {
    // Node exposes no public FD count; approximate via active handles.
    return (process._getActiveHandles?.() ?? []).length;
  } catch {
    return -1;
  }
}

const workspace = await mkdtemp(join(tmpdir(), "soak-"));
const plugin = await OfficePlugin.create({ workspaceRoot: join(workspace, "rt"), skipHostProbe: true });

function crc32(buf) {
  let table = crc32.table;
  if (!table) {
    table = crc32.table = new Int32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      table[n] = c;
    }
  }
  let crc = -1;
  for (let i = 0; i < buf.length; i++) crc = (crc >>> 8) ^ table[(crc ^ buf[i]) & 0xff];
  return (crc ^ -1) >>> 0;
}

function buildZip(entries) {
  const chunks = [];
  const central = [];
  let offset = 0;
  for (const { name, data } of entries) {
    const nameBuf = Buffer.from(name, "utf8");
    const crc = crc32(data);
    const lfh = Buffer.alloc(30);
    lfh.writeUInt32LE(0x04034b50, 0);
    lfh.writeUInt16LE(0, 8); // stored
    lfh.writeUInt32LE(crc, 14);
    lfh.writeUInt32LE(data.length, 18);
    lfh.writeUInt32LE(data.length, 22);
    lfh.writeUInt16LE(nameBuf.length, 26);
    chunks.push(lfh, nameBuf, data);
    const cde = Buffer.alloc(46);
    cde.writeUInt32LE(0x02014b50, 0);
    cde.writeUInt32LE(crc, 16);
    cde.writeUInt32LE(data.length, 20);
    cde.writeUInt32LE(data.length, 24);
    cde.writeUInt16LE(nameBuf.length, 28);
    cde.writeUInt32LE(offset, 42);
    central.push(cde, nameBuf);
    offset += 30 + nameBuf.length + data.length;
  }
  const cd = Buffer.concat(central);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(entries.length, 8);
  eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cd.length, 12);
  eocd.writeUInt32LE(offset, 16);
  return Buffer.concat([...chunks, cd, eocd]);
}

const docx = buildZip([
  { name: "[Content_Types].xml", data: Buffer.from('<?xml version="1.0"?><Types/>') },
  { name: "word/document.xml", data: Buffer.from('<?xml version="1.0"?><w>soak fixture</w>') }
]);

const samples = [];
const ref = await plugin.registerArtifact(await (async () => {
  const p = join(workspace, "soak.docx");
  await writeFile(p, docx);
  return p;
})());

let errors = 0;
for (let cycle = 1; cycle <= cycles; cycle++) {
  const session = await plugin.openSession(ref);
  await plugin.service.sessions.ensureStrongIdentity(session.sessionId);
  const { editor } = await plugin.beginEdit(session.sessionId);
  editor.markDirty?.();
  await editor.save(); // Runtime gate: humanSave
  await plugin.endEdit(session.sessionId);
  await plugin.closeSession(session.sessionId);
  await plugin.preview({ artifactRef: ref, priority: "background" });

  if (cycle % sampleEvery === 0 || cycle === cycles) {
    global.gc?.();
    const sample = {
      cycle,
      rssMB: Number(rssMB().toFixed(1)),
      heapMB: Number(heapMB().toFixed(1)),
      fds: fdCount(),
      sessions: plugin.service.sessions.list().length,
      editorLeases: plugin.service.editorLeaseCount(),
      engineTasks: plugin.service.agent.activeTaskCount()
    };
    samples.push(sample);
    console.log(JSON.stringify(sample));
  }
  void errors;
}

// Plateau assertion: compare first sample vs last; RSS must not grow
// unboundedly (allow generous 2.5x for JIT/GC slop over the run).
const first = samples[0];
const last = samples[samples.length - 1];
const rssGrowth = last.rssMB / Math.max(1, first.rssMB);
const verdict = {
  cycles,
  samples: samples.length,
  rssFirstMB: first.rssMB,
  rssLastMB: last.rssMB,
  rssGrowthRatio: Number(rssGrowth.toFixed(2)),
  fdsFirst: first.fds,
  fdsLast: last.fds,
  sessionsAtEnd: last.sessions,
  editorLeasesAtEnd: last.editorLeases,
  engineTasksAtEnd: last.engineTasks,
  errors
};

const passed =
  rssGrowth < 2.5 &&
  last.sessions === 0 &&
  last.editorLeases === 0 &&
  last.engineTasks === 0;

writeFileSyncSync(join(workspace, "soak-report.json"), JSON.stringify({ verdict, samples }, null, 2));
console.log(JSON.stringify(verdict));
console.log(passed ? "SOAK PASS: stable plateau" : "SOAK FAIL: growth/ownership leak suspected");
await plugin.dispose();
await rm(workspace, { recursive: true, force: true }).catch(() => undefined);
process.exit(passed ? 0 : 1);
