/**
 * SVG → PNG rasterization (headless, §91 Visual Layer).
 *
 * Static visual output no longer needs any host: engine draw lists → SVG
 * (render-svg) → PNG via sharp (prebuilt native, zero install scripts).
 * Serving sizes respect byte budgets — decoded pixels live only in the
 * raster call, only the encoded PNG bytes are cached.
 */

import sharp from "sharp";
import { ByteBudgetCache } from "../artifact/cache/byte-budget-cache.js";

export interface RasterOptions {
  /** Target render width in px (height follows the SVG aspect). */
  width?: number;
}

export interface RasterResult {
  png: Buffer;
  width: number;
  height: number;
}

/** Encoded PNG cache keyed by (svg hash, width). §101: byte budget, not count. */
export class RasterCache {
  private readonly cache: ByteBudgetCache<RasterResult>;

  constructor(maxItemShare = 0.2) {
    this.cache = new ByteBudgetCache<RasterResult>(
      "VisualPreviewCache",
      64 * 1024 * 1024,
      (value) => value.png.length,
      { maxItemShareOfBudget: maxItemShare }
    );
  }

  async rasterize(svg: string, options: RasterOptions = {}): Promise<RasterResult> {
    const width = options.width ?? 960;
    const key = `${Buffer.from(svg).length}:${hashSvg(svg)}:${width}`;
    const cached = this.cache.get(key);
    if (cached) return cached;

    const { data, info } = await sharp(Buffer.from(svg), { density: 96 })
      .resize({ width, withoutEnlargement: false })
      .png({ compressionLevel: 6 })
      .toBuffer({ resolveWithObject: true });
    const result: RasterResult = { png: data, width: info.width, height: info.height };
    if (this.cache.admit(result.png.length)) {
      this.cache.set(key, result);
    }
    return result;
  }

  stats(): { bytes: number; items: number } {
    return { bytes: this.cache.currentBytes, items: this.cache.length };
  }

  trimAll(): number {
    return this.cache.trimAll();
  }
}

function hashSvg(svg: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < svg.length; i++) {
    hash ^= svg.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(36);
}
