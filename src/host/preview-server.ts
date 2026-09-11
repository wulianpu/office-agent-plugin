/**
 * Localhost web host (§P7 renderer container): serves the plugin's preview
 * surfaces over HTTP so ANY browser becomes the rendering process — no
 * Electron required. An Electron host is then a thin BrowserWindow wrapper
 * pointed at the same origin (tools/host/electron.mjs).
 *
 * Surfaces: gallery (registered artifacts), slide viewer (SVG/PNG), xlsx
 * grid window, JSON APIs. Physical paths never appear in responses (INV-12).
 */

import { createServer, type IncomingMessage, type Server, type ServerResponse } from "node:http";
import type { OfficePlugin } from "../plugin/office-plugin.js";
import { RasterCache } from "../preview/raster.js";

export interface PreviewHostOptions {
  port?: number;
  host?: string;
}

export class PreviewHost {
  readonly raster: RasterCache;
  private server?: Server;
  private boundPort?: number;

  constructor(
    private readonly plugin: OfficePlugin,
    raster?: RasterCache
  ) {
    this.raster = raster ?? plugin.service.raster;
  }

  get port(): number | undefined {
    return this.boundPort;
  }

  async start(options: PreviewHostOptions = {}): Promise<number> {
    if (this.server) return this.boundPort!;
    const server = createServer((req, res) => {
      void this.handle(req, res).catch((error) => {
        res.writeHead(500, { "content-type": "application/json" });
        res.end(JSON.stringify({ error: String((error as Error)?.message ?? error) }));
      });
    });
    await new Promise<void>((resolve) => server.listen(options.port ?? 0, options.host ?? "127.0.0.1", resolve));
    const address = server.address();
    this.server = server;
    this.boundPort = typeof address === "object" && address ? address.port : options.port ?? 0;
    return this.boundPort;
  }

  async stop(): Promise<void> {
    const server = this.server;
    if (!server) return;
    await new Promise<void>((resolve) => server.close(() => resolve()));
    this.server = undefined;
    this.boundPort = undefined;
  }

  url(path = "/"): string {
    return `http://127.0.0.1:${this.boundPort}${path}`;
  }

  private async handle(req: IncomingMessage, res: ServerResponse): Promise<void> {
    const url = new URL(req.url ?? "/", "http://localhost");
    const parts = url.pathname.split("/").filter(Boolean);

    if (url.pathname === "/" || url.pathname === "/app") {
      return html(res, galleryPage());
    }
    if (url.pathname === "/api/artifacts") {
      const artifacts = this.plugin.service.repos
        .loadArtifacts()
        .filter((a) => a.kind === "source")
        .map((a) => ({ artifactRef: a.ref, format: a.format }));
      return json(res, 200, { artifacts });
    }
    if (parts[0] === "api" && parts[1] === "preview" && parts[2]) {
      const model = await this.plugin.preview({ artifactRef: parts[2], priority: "visible" });
      // INV-12: strip anything that could carry a physical path.
      return json(res, 200, JSON.parse(JSON.stringify({ format: model.model.format, outline: model.model.outline, svgCount: model.model.svgSlides?.length ?? 0 })));
    }
    if (parts[0] === "svg" && parts[1] && parts[2]) {
      const model = await this.plugin.preview({ artifactRef: parts[1], priority: "visible", visual: true });
      const svg = model.model.svgSlides?.[Number(parts[2])];
      if (!svg) return json(res, 404, { error: "slide svg not found" });
      res.writeHead(200, { "content-type": "image/svg+xml", "cache-control": "no-store" });
      res.end(svg);
      return;
    }
    if (parts[0] === "png" && parts[1] && parts[2]) {
      const model = await this.plugin.preview({ artifactRef: parts[1], priority: "visible", visual: true });
      const svg = model.model.svgSlides?.[Number(parts[2])];
      if (!svg) return json(res, 404, { error: "slide svg not found" });
      const raster = await this.raster.rasterize(svg, { width: Number(url.searchParams.get("w") ?? 960) });
      res.writeHead(200, { "content-type": "image/png", "cache-control": "no-store" });
      res.end(raster.png);
      return;
    }
    json(res, 404, { error: "not found" });
  }
}

function json(res: ServerResponse, status: number, payload: unknown): void {
  res.writeHead(status, { "content-type": "application/json", "cache-control": "no-store" });
  res.end(JSON.stringify(payload));
}

function html(res: ServerResponse, page: string): void {
  res.writeHead(200, { "content-type": "text/html; charset=utf-8", "cache-control": "no-store" });
  res.end(page);
}

/** Single-page host UI: gallery → slide viewer (SVG) / sheet grid window. */
function galleryPage(): string {
  return `<!doctype html>
<html lang="zh">
<head>
<meta charset="utf-8">
<title>Office Plugin Host</title>
<style>
  body { font-family: system-ui, sans-serif; margin: 0; background: #1e1e2e; color: #cdd6f4; }
  header { padding: 12px 20px; background: #181825; border-bottom: 1px solid #313244; }
  header h1 { font-size: 16px; margin: 0; }
  main { display: flex; height: calc(100vh - 46px); }
  #list { width: 260px; overflow: auto; border-right: 1px solid #313244; padding: 8px; }
  #list button { display: block; width: 100%; text-align: left; margin-bottom: 6px; padding: 8px;
    background: #313244; color: #cdd6f4; border: 0; border-radius: 6px; cursor: pointer; }
  #list button.active { background: #89b4fa; color: #1e1e2e; }
  #stage { flex: 1; overflow: auto; padding: 16px; position: relative; }
  #stage img, #stage svg { max-width: 100%; box-shadow: 0 4px 24px rgba(0,0,0,.5); border-radius: 4px; background: #fff; }
  #nav { margin-top: 10px; display: flex; gap: 8px; align-items: center; }
  #nav button { background: #45475a; color: #cdd6f4; border: 0; border-radius: 6px; padding: 6px 14px; cursor: pointer; }
  table { border-collapse: collapse; background: #fff; color: #111; }
  td, th { border: 1px solid #ccc; padding: 3px 10px; font-size: 13px; min-width: 48px; }
  .pill { font-size: 11px; background: #45475a; border-radius: 8px; padding: 2px 8px; margin-left: 8px; }
</style>
</head>
<body>
<header><h1>Office Plugin — localhost 宿主<span class="pill">SVG/PNG 视觉 · Rust sidecar</span></h1></header>
<main>
  <nav id="list"></nav>
  <section id="stage"></section>
</main>
<script type="module">
const list = document.getElementById('list');
const stage = document.getElementById('stage');
let current = null, state = { slide: 0, svgCount: 0, format: null };

async function refresh() {
  const { artifacts } = await (await fetch('/api/artifacts')).json();
  list.innerHTML = '';
  for (const a of artifacts) {
    const b = document.createElement('button');
    b.textContent = a.format.toUpperCase() + ' · ' + a.artifactRef.slice(-8);
    b.onclick = () => { document.querySelectorAll('#list button').forEach(x=>x.classList.remove('active')); b.classList.add('active'); open(a.artifactRef); };
    list.appendChild(b);
  }
}
async function open(ref) {
  current = ref; state.slide = 0;
  const model = await (await fetch('/api/preview/' + ref)).json();
  state.svgCount = model.svgCount; state.format = model.format;
  render(model);
}
function render(model) {
  if (state.format === 'pptx' && state.svgCount > 0) {
    stage.innerHTML = '';
    const img = document.createElement('img');
    img.src = '/svg/' + current + '/' + state.slide;
    img.alt = 'slide ' + (state.slide + 1);
    stage.appendChild(img);
    const nav = document.createElement('div'); nav.id = 'nav';
    const prev = document.createElement('button'); prev.textContent = '← 上一页';
    const label = document.createElement('span'); label.textContent = (state.slide + 1) + ' / ' + state.svgCount;
    const next = document.createElement('button'); next.textContent = '下一页 →';
    prev.onclick = () => { if (state.slide > 0) { state.slide--; render(model); } };
    next.onclick = () => { if (state.slide < state.svgCount - 1) { state.slide++; render(model); } };
    nav.append(prev, label, next); stage.appendChild(nav);
  } else if (model.outline.kind === 'xlsx') {
    const table = document.createElement('table');
    for (const sheet of model.outline.sheets) {
      const cap = document.createElement('caption');
      cap.textContent = sheet.name + ' (' + sheet.rowCount + ' rows)';
      cap.style.cssText = 'caption-side:top;text-align:left;padding:6px;color:#89b4fa';
      table.appendChild(cap);
      for (const row of sheet.window) {
        const tr = document.createElement('tr');
        for (const cell of row) { const td = document.createElement('td'); td.textContent = cell; tr.appendChild(td); }
        table.appendChild(tr);
      }
    }
    stage.innerHTML = ''; stage.appendChild(table);
  } else {
    const pre = document.createElement('pre');
    pre.style.cssText = 'white-space:pre-wrap;background:#11111b;padding:16px;border-radius:8px';
    for (const block of (model.outline.blocks ?? [])) {
      pre.textContent += (block.style ? '[' + block.style + '] ' : '') + block.text + '\\n';
    }
    stage.innerHTML = ''; stage.appendChild(pre);
  }
}
refresh();
</script>
</body>
</html>`;
}
