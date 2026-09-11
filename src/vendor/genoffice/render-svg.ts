/**
 * RenderTree → SVG adapter (headless visual rendering, §P7 data plane).
 *
 * The vendored pptx-render engine (unmodified) converts slides into a pure
 * data draw list (RenderSlide: pixel geometry + resolved fills + laid-out
 * glyph boxes — "unit-testable without a canvas" per its design). This module
 * is our thin adapter that serializes that list into standalone SVG strings:
 * zero DOM, zero canvas, usable from Node, the utility process and any host.
 *
 * v1 coverage: backgrounds, shapes (path/polygon/roundRect/line/ellipse),
 * laid-out text runs, pictures (dataUrl), groups, tables, gradient fills.
 * Charts/placeholder chips render as labeled boxes (honest approximation).
 */

import type { PptxDeck } from "./genoffice-types.js";

// Structural mirrors of the engine's render types (kept local so engine type
// churn inside the vendored tree cannot break this adapter's compilation).
interface PlacedBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

type RenderFill =
  | { kind: "none" }
  | { kind: "solid"; color: string }
  | { kind: "gradient"; stops: Array<{ pos: number; color: string }>; angleDeg: number; radial?: boolean }
  | { kind: "image"; dataUrl?: string; mode?: "stretch" | "tile" }
  | { kind: "pattern"; fgColor?: string; bgColor?: string }
  | { [key: string]: unknown };

interface RenderStroke {
  color: string;
  widthPx: number;
  dashStyle?: string;
  [key: string]: unknown;
}

interface GlyphRun {
  text: string;
  x: number;
  baselineY: number;
  fontFamily: string;
  fontSizePx: number;
  color: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike?: boolean;
  highlight?: string;
}

interface RenderTextLayout {
  lines: Array<{ runs: GlyphRun[]; top: number; height: number }>;
  insets: { l: number; t: number; r: number; b: number };
  anchor: "top" | "middle" | "bottom";
  contentHeight: number;
}

interface RenderNodeBase {
  type: string;
  box: PlacedBox;
  [key: string]: unknown;
}

interface ShapeNode extends RenderNodeBase {
  type: "shape" | "text";
  presetGeometry?: string;
  cornerRadiusPx?: number;
  polygonPoints?: number[];
  pathData?: string;
  line?: { points: number[] };
  fill: RenderFill;
  stroke?: RenderStroke;
  text?: RenderTextLayout;
}

interface PictureNode extends RenderNodeBase {
  type: "picture";
  dataUrl?: string;
  bgColor?: string;
  opacity?: number;
}

interface GroupNode extends RenderNodeBase {
  type: "group";
  children: RenderNode[];
}

interface TableCellNode {
  x: number;
  y: number;
  w: number;
  h: number;
  fill: RenderFill;
  borders?: { l?: RenderStroke; r?: RenderStroke; t?: RenderStroke; b?: RenderStroke };
  text?: RenderTextLayout;
}

interface TableNode extends RenderNodeBase {
  type: "table";
  cells: TableCellNode[];
  bgFill?: RenderFill;
}

interface ChipNode extends RenderNodeBase {
  type: "placeholder-chip";
  label?: string;
}

interface ChartNode extends RenderNodeBase {
  type: "chart";
}

type RenderNode = ShapeNode | PictureNode | GroupNode | TableNode | ChartNode | ChipNode;

interface RenderSlide {
  widthPx: number;
  heightPx: number;
  background: RenderFill;
  nodes: RenderNode[];
}

interface SvgContext {
  defs: string[];
  defSeq: number;
}

const esc = (text: string): string =>
  text.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const n = (value: number): string => (Number.isInteger(value) ? String(value) : value.toFixed(2));

/** Render one slide's draw list to a standalone SVG string. */
export function renderSlideToSvg(slide: RenderSlide): string {
  const ctx: SvgContext = { defs: [], defSeq: 0 };
  const background = fillToAttr(ctx, slide.background, { x: 0, y: 0, w: slide.widthPx, h: slide.heightPx }, true);
  const body = slide.nodes.map((node) => renderNode(ctx, node, 0, 0)).join("\n");
  const defs = ctx.defs.length > 0 ? `<defs>${ctx.defs.join("")}</defs>` : "";
  return (
    `<svg xmlns="http://www.w3.org/2000/svg" width="${n(slide.widthPx)}" height="${n(slide.heightPx)}" ` +
    `viewBox="0 0 ${n(slide.widthPx)} ${n(slide.heightPx)}">` +
    defs +
    `<rect x="0" y="0" width="${n(slide.widthPx)}" height="${n(slide.heightPx)}" ${background}/>` +
    body +
    `</svg>`
  );
}

function renderNode(ctx: SvgContext, node: RenderNode, ox: number, oy: number): string {
  const x = ox + node.box.x;
  const y = oy + node.box.y;
  switch (node.type) {
    case "group":
      // Children carry group-local coordinates relative to the group top-left.
      return `<g transform="translate(${n(x)} ${n(y)})">${(node as GroupNode).children
        .map((child) => renderNode(ctx, child, 0, 0))
        .join("")}</g>`;
    case "picture":
      return renderPicture(node as PictureNode, x, y);
    case "table":
      return renderTable(ctx, node as TableNode, x, y);
    case "placeholder-chip":
      return renderChip(node as ChipNode, x, y);
    case "chart":
      return renderChart(node as ChartNode, x, y);
    default:
      return renderShape(ctx, node as ShapeNode, x, y);
  }
}

function renderShape(ctx: SvgContext, node: ShapeNode, x: number, y: number): string {
  const { w, h } = node.box;
  const fill = fillToAttr(ctx, node.fill, { x, y, w, h }, false);
  const stroke = strokeToAttr(node.stroke);
  let geometry = "";

  if (node.pathData) {
    geometry = `<path d="${esc(node.pathData)}" transform="translate(${n(x)} ${n(y)})" ${fill} ${stroke}/>`;
  } else if (node.polygonPoints && node.polygonPoints.length >= 6) {
    const pts = node.polygonPoints.map((v, i) => (i % 2 === 0 ? n(x + v) : n(y + v))).join(" ");
    geometry = `<polygon points="${pts}" ${fill} ${stroke}/>`;
  } else if (node.line?.points && node.line.points.length >= 4) {
    const pts = node.line.points.map((v, i) => (i % 2 === 0 ? n(x + v) : n(y + v))).join(" ");
    geometry = `<polyline points="${pts}" fill="none" ${stroke}/>`;
  } else if (node.cornerRadiusPx && node.cornerRadiusPx > 0.5) {
    geometry = `<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" rx="${n(Math.min(node.cornerRadiusPx, Math.min(w, h) / 2))}" ${fill} ${stroke}/>`;
  } else {
    geometry = `<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" ${fill} ${stroke}/>`;
  }

  const text = node.text ? renderText(node.text, x, y, node.box) : "";
  return geometry + text;
}

function renderText(text: RenderTextLayout, x: number, y: number, box: PlacedBox): string {
  const innerW = Math.max(1, box.w - text.insets.l - text.insets.r);
  let offsetY = text.insets.t;
  if (text.anchor === "middle") {
    offsetY += Math.max(0, (box.h - text.insets.t - text.insets.b - text.contentHeight) / 2);
  } else if (text.anchor === "bottom") {
    offsetY += Math.max(0, box.h - text.insets.t - text.insets.b - text.contentHeight);
  }
  const parts: string[] = [];
  for (const line of text.lines) {
    for (const run of line.runs) {
      if (!run.text) continue;
      const style = `font-family:${esc(run.fontFamily)};font-size:${n(run.fontSizePx)}px;` +
        `${run.bold ? "font-weight:bold;" : ""}${run.italic ? "font-style:italic;" : ""}fill:${run.color}`;
      let extra = "";
      if (run.underline) extra += `<rect x="${n(x + text.insets.l + run.x)}" y="${n(y + offsetY + run.baselineY + 2)}" width="${n(run.text.length * run.fontSizePx * 0.5)}" height="1" fill="${run.color}"/>`;
      const highlight = run.highlight
        ? `<rect x="${n(x + text.insets.l + run.x - 1)}" y="${n(y + offsetY + run.baselineY - run.fontSizePx)}" width="${n(run.text.length * run.fontSizePx * 0.55 + 2)}" height="${n(run.fontSizePx * 1.2)}" fill="${run.highlight}"/>`
        : "";
      parts.push(
        `${highlight}<text x="${n(x + text.insets.l + run.x)}" y="${n(y + offsetY + run.baselineY)}" ` +
          `xml:space="preserve" style="${style}">${esc(run.text)}</text>${extra}`
      );
    }
  }
  // Clip overflow when the layout says the text wraps within the box.
  if (parts.length === 0) return "";
  const clipId = `clip-${x.toFixed(0)}-${y.toFixed(0)}-${Math.random().toString(36).slice(2, 7)}`;
  return `<clipPath id="${clipId}"><rect x="${n(x)}" y="${n(y)}" width="${n(box.w)}" height="${n(box.h)}"/></clipPath>` +
    `<g clip-path="url(#${clipId})">${parts.join("")}</g>`;
}

function renderPicture(node: PictureNode, x: number, y: number): string {
  const { w, h } = node.box;
  if (!node.dataUrl) {
    const bg = node.bgColor ?? "#f0f0f0";
    return `<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" fill="${bg}" stroke="#ccc"/>`;
  }
  const opacity = node.opacity !== undefined && node.opacity < 1 ? ` opacity="${n(node.opacity)}"` : "";
  const backdrop = node.bgColor ? `<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" fill="${node.bgColor}"/>` : "";
  return `${backdrop}<image x="${n(x)}" y="${n(y)}" width="${n(w)}" height="${n(h)}" href="${esc(node.dataUrl)}"${opacity}/>`;
}

function renderTable(ctx: SvgContext, node: TableNode, x: number, y: number): string {
  const parts: string[] = [];
  if (node.bgFill && node.bgFill.kind !== "none") {
    parts.push(`<rect x="${n(x)}" y="${n(y)}" width="${n(node.box.w)}" height="${n(node.box.h)}" ${fillToAttr(ctx, node.bgFill, { x, y, w: node.box.w, h: node.box.h }, true)}/>`);
  }
  for (const cell of node.cells) {
    const cx = x + cell.x;
    const cy = y + cell.y;
    if (cell.fill && cell.fill.kind !== "none") {
      parts.push(`<rect x="${n(cx)}" y="${n(cy)}" width="${n(cell.w)}" height="${n(cell.h)}" ${fillToAttr(ctx, cell.fill, { x: cx, y: cy, w: cell.w, h: cell.h }, true)}/>`);
    }
    const b = cell.borders ?? {};
    const border = (side: "l" | "r" | "t" | "b", x1: number, y1: number, x2: number, y2: number) => {
      const s = b[side];
      if (!s) return "";
      return `<line x1="${n(x1)}" y1="${n(y1)}" x2="${n(x2)}" y2="${n(y2)}" stroke="${s.color}" stroke-width="${n(s.widthPx)}"/>`;
    };
    parts.push(
      border("t", cx, cy, cx + cell.w, cy) +
        border("b", cx, cy + cell.h, cx + cell.w, cy + cell.h) +
        border("l", cx, cy, cx, cy + cell.h) +
        border("r", cx + cell.w, cy, cx + cell.w, cy + cell.h)
    );
    if (cell.text) {
      parts.push(renderText(cell.text, cx, cy, { x: 0, y: 0, w: cell.w, h: cell.h }));
    }
  }
  return `<g>${parts.join("")}</g>`;
}

function renderChip(node: ChipNode, x: number, y: number): string {
  const label = esc(node.label ?? String((node as { label?: string }).label ?? "Element"));
  return (
    `<rect x="${n(x)}" y="${n(y)}" width="${n(node.box.w)}" height="${n(node.box.h)}" fill="#f5f5f5" stroke="#bbb" rx="4"/>` +
    `<text x="${n(x + 6)}" y="${n(y + node.box.h / 2)}" font-size="12" fill="#666">${label}</text>`
  );
}

function renderChart(node: ChartNode, x: number, y: number): string {
  return (
    `<rect x="${n(x)}" y="${n(y)}" width="${n(node.box.w)}" height="${n(node.box.h)}" fill="#fafafa" stroke="#ccc" rx="4"/>` +
    `<text x="${n(x + 8)}" y="${n(y + 18)}" font-size="12" fill="#888">Chart</text>`
  );
}

function fillToAttr(ctx: SvgContext, fill: RenderFill | undefined, box: PlacedBox, isBackground: boolean): string {
  if (!fill || fill.kind === "none") return 'fill="none"';
  if (fill.kind === "solid") return `fill="${esc((fill as { color: string }).color)}"`;
  if (fill.kind === "image") {
    const dataUrl = (fill as { dataUrl?: string }).dataUrl;
    return dataUrl ? `fill="none"` : 'fill="none"';
  }
  if (fill.kind === "gradient") {
    const gradient = fill as { stops?: Array<{ pos: number; color: string }>; angleDeg?: number; radial?: boolean };
    const stops = gradient.stops ?? [];
    if (stops.length >= 2) {
      const id = `grad-${++ctx.defSeq}`;
      const stopXml = stops
        .map((s) => `<stop offset="${n(Math.max(0, Math.min(1, s.pos)) * 100)}%" stop-color="${esc(s.color)}"/>`)
        .join("");
      if (gradient.radial) {
        ctx.defs.push(`<radialGradient id="${id}" cx="50%" cy="50%" r="65%">${stopXml}</radialGradient>`);
      } else {
        // Convert OOXML angle (clockwise from east) to SVG gradient vector.
        const rad = ((gradient.angleDeg ?? 0) * Math.PI) / 180;
        const dx = Math.cos(rad) / 2;
        const dy = -Math.sin(rad) / 2;
        ctx.defs.push(
          `<linearGradient id="${id}" x1="${n(0.5 - dx)}" y1="${n(0.5 - dy)}" x2="${n(0.5 + dx)}" y2="${n(0.5 + dy)}">${stopXml}</linearGradient>`
        );
      }
      return `fill="url(#${id})"`;
    }
  }
  if (fill.kind === "pattern") {
    return `fill="${esc((fill as { bgColor?: string }).bgColor ?? "#eeeeee")}"`;
  }
  return isBackground ? 'fill="#ffffff"' : 'fill="none"';
}

function strokeToAttr(stroke: RenderStroke | undefined): string {
  if (!stroke || stroke.widthPx === 0) return "";
  const dash = stroke.dashStyle && stroke.dashStyle !== "solid" ? ` stroke-dasharray="6 4"` : "";
  return `stroke="${esc(stroke.color)}" stroke-width="${n(stroke.widthPx)}"${dash}`;
}

/** Type assertion helpers used by the preview integration. */
export function isRenderSlideLike(value: unknown): value is RenderSlide {
  return (
    typeof value === "object" &&
    value !== null &&
    "nodes" in value &&
    "widthPx" in value &&
    Array.isArray((value as RenderSlide).nodes)
  );
}

export type { RenderSlide, RenderNode, PptxDeck };
