/**
 * Small XML helpers: entity decoding, streaming text/row extraction for
 * previews, and a canonical form used by PackageDiff NORMALIZATION_ONLY
 * classification (§80). Deliberately minimal — this is not a DOM.
 */

export function decodeXmlEntities(input: string): string {
  if (!input.includes("&")) return input;
  return input
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
    .replace(/&amp;/g, "&");
}

export function encodeXmlText(input: string): string {
  return input
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

/**
 * Extract inner texts of a repeating tag (e.g. `<a:t>…</a:t>`) from a chunk.
 * `maxPerChunk` bounds extraction work per call. Works incrementally: callers
 * may concatenate text across chunk boundaries via `carry`.
 */
export function extractTagTexts(
  chunk: string,
  tag: string,
  carry: { pending: string }
): string[] {
  const results: string[] = [];
  let data = carry.pending + chunk;
  carry.pending = "";
  const open = `<${tag}`;
  const close = `</${tag}>`;
  let searchFrom = 0;
  for (;;) {
    const openIdx = data.indexOf(open, searchFrom);
    if (openIdx < 0) break;
    const gtIdx = data.indexOf(">", openIdx);
    if (gtIdx < 0) {
      carry.pending = data.slice(Math.max(0, openIdx - 1));
      break;
    }
    if (data[gtIdx - 1] === "/") {
      searchFrom = gtIdx + 1;
      continue;
    }
    const closeIdx = data.indexOf(close, gtIdx);
    if (closeIdx < 0) {
      carry.pending = data.slice(openIdx);
      break;
    }
    results.push(decodeXmlEntities(data.slice(gtIdx + 1, closeIdx)));
    searchFrom = closeIdx + close.length;
  }
  if (!carry.pending && searchFrom > 0) {
    // Keep a tail window in case an open tag straddles the next chunk boundary.
    carry.pending = data.slice(Math.max(0, data.length - close.length - open.length - 2));
  }
  return results;
}

/**
 * Canonical form for normalization comparison: attribute order, quote style,
 * and whitespace-only text differences vanish; everything else remains.
 */
export function canonicalizeXml(xml: string): string {
  const tokens: string[] = [];
  const len = xml.length;
  let i = 0;
  while (i < len) {
    if (xml[i] === "<") {
      if (xml.startsWith("<!--", i)) {
        const end = xml.indexOf("-->", i);
        i = end < 0 ? len : end + 3;
        continue;
      }
      const end = xml.indexOf(">", i);
      if (end < 0) break;
      let tag = xml.slice(i + 1, end);
      const isClose = tag.startsWith("/");
      const isSelfClose = tag.endsWith("/");
      if (isSelfClose) tag = tag.slice(0, -1);
      if (isClose) {
        tokens.push(`</${tagNameOf(tag)}>`);
      } else {
        const { name, attrs } = parseTag(tag);
        const sorted = [...attrs].sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0));
        tokens.push(
          `<${name}${sorted.map(([k, v]) => ` ${k}=${JSON.stringify(v)}`).join("")}${isSelfClose ? "/" : ""}>`
        );
      }
      i = end + 1;
    } else {
      const next = xml.indexOf("<", i);
      const text = xml.slice(i, next < 0 ? len : next);
      if (text.trim().length > 0) tokens.push(decodeXmlEntities(text.trim()));
      i = next < 0 ? len : next;
    }
  }
  return tokens.join("\u0001");
}

function tagNameOf(tag: string): string {
  return tag.replace(/^[/\s]+/, "").split(/[\s/>]/)[0] ?? "";
}

function parseTag(tag: string): { name: string; attrs: Array<[string, string]> } {
  const spaceIdx = tag.search(/[\s]/);
  const name = spaceIdx < 0 ? tag : tag.slice(0, spaceIdx);
  const attrs: Array<[string, string]> = [];
  const attrRe = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let rest = spaceIdx < 0 ? "" : tag.slice(spaceIdx);
  let m: RegExpExecArray | null;
  while ((m = attrRe.exec(rest)) !== null) {
    attrs.push([m[1]!, m[3] ?? m[4] ?? ""]);
  }
  // Drop namespace prefix noise like mc:Ignorable ordering — keep values intact.
  void rest;
  return { name, attrs };
}

/**
 * Split a stream into complete `<tag ...>…</tag>` records (namespace-tolerant:
 * `<x:row>`/`<x:si>` match too), carrying partial tags across chunk
 * boundaries. Chunks may end mid open-tag — the tail is carried from its
 * last `<` only when no `>` follows it (provably partial), so records are
 * never re-emitted.
 */
export function splitTagged(chunk: string, tag: string, carry: { pending: string }): string[] {
  const records: string[] = [];
  let data = carry.pending + chunk;
  carry.pending = "";
  const open = new RegExp(`<(?:[\\w.-]+:)?${tag}\\b`, "g");
  const close = new RegExp(`</(?:[\\w.-]+:)?${tag}>`, "g");
  let lastScan = 0;
  for (;;) {
    open.lastIndex = lastScan;
    const opened = open.exec(data);
    if (!opened) break;
    close.lastIndex = opened.index + opened[0].length;
    const closed = close.exec(data);
    if (!closed) {
      carry.pending = data.slice(opened.index);
      break;
    }
    records.push(data.slice(opened.index, closed.index + closed[0].length));
    lastScan = closed.index + closed[0].length;
  }
  if (!carry.pending) {
    const unconsumed = data.slice(lastScan);
    const lastLt = unconsumed.lastIndexOf("<");
    if (lastLt >= 0 && !unconsumed.slice(lastLt).includes(">")) {
      carry.pending = unconsumed.slice(lastLt);
    }
  }
  return records;
}

/** Split a large string across row boundaries `<row ...>…</row>`. */
export function splitRows(chunk: string, carry: { pending: string }): string[] {
  return splitTagged(chunk, "row", carry);
}

/**
 * Concatenate ALL inner `<tag>` texts of one COMPLETE record, namespace-
 * tolerant and entity-decoded. Shared strings use this over `<si>`: a Rich
 * Text `<si>` holds many `<r><t>` runs that must join into ONE entry —
 * indexing per `<t>` silently shifted every later index (round 9).
 */
export function concatenatedTagTexts(record: string, tag: string): string {
  const inner = new RegExp(
    `<(?:[\\w.-]+:)?${tag}(?:\\s[^>]*)?>([^<]*)</(?:[\\w.-]+:)?${tag}>`,
    "g"
  );
  let text = "";
  for (const m of record.matchAll(inner)) text += decodeXmlEntities(m[1] ?? "");
  return text;
}
