import { describe, expect, it } from "vitest";
import {
  canonicalizeXml,
  extractTagTexts,
  splitRows
} from "../../src/support/xml-lite.js";

describe("xml-lite", () => {
  it("extractTagTexts decodes entities and spans chunk boundaries", () => {
    const carry = { pending: "" };
    expect(extractTagTexts("<a:t>Hello &amp; bye</a:t>", "a:t", carry)).toEqual(["Hello & bye"]);
    expect(extractTagTexts("<a:t>Part1</a:t><a:t>Part", "a:t", carry)).toEqual(["Part1"]);
    expect(extractTagTexts("2</a:t>", "a:t", carry)).toEqual(["Part2"]);
  });

  it("canonicalizeXml is insensitive to attribute order, quotes and whitespace", () => {
    const a = `<a x="1" y='2'><b>text</b>  <c/></a>`;
    const b = `<a y="2" x="1">\n  <b>text</b><c/></a>`;
    expect(canonicalizeXml(a)).toBe(canonicalizeXml(b));
    const c = `<a x="1" y="2"><b>different</b><c/></a>`;
    expect(canonicalizeXml(a)).not.toBe(canonicalizeXml(c));
  });

  it("splitRows returns complete rows and carries partial tails", () => {
    const carry = { pending: "" };
    expect(splitRows('<row r="1"><c/></row><row r=', carry)).toEqual(['<row r="1"><c/></row>']);
    expect(splitRows('"2"><c/></row>', carry)).toEqual(['<row r="2"><c/></row>']);
  });

  it("splitRows keeps a namespaced partial open-tag after an extracted row (round 8)", () => {
    // Chunk 1 already yielded a complete row; the tail is half of `<x:row`.
    // Pre-fix, lastScan>0 dropped the tail and row 2 vanished forever.
    const carry = { pending: "" };
    expect(splitRows('<x:row r="1"><x:c/></x:row><x:r', carry)).toEqual([
      '<x:row r="1"><x:c/></x:row>'
    ]);
    expect(splitRows('ow r="2"><x:c/></x:row>', carry)).toEqual(['<x:row r="2"><x:c/></x:row>']);
  });

  it("splitRows is deterministic across every chunk boundary", () => {
    const xml =
      '<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>' +
      '<x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row>' +
      '<x:row r="3"><x:c r="A3"><x:v>3</x:v></x:c></x:row>';
    const expected = [
      '<x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row>',
      '<x:row r="2"><x:c r="A2"><x:v>2</x:v></x:c></x:row>',
      '<x:row r="3"><x:c r="A3"><x:v>3</x:v></x:c></x:row>'
    ];
    // Split at EVERY character boundary; every split must reassemble the
    // same three rows (chunk boundaries in ZIP streaming are arbitrary).
    for (let cut = 1; cut < xml.length - 1; cut++) {
      const carry = { pending: "" };
      const head = splitRows(xml.slice(0, cut), carry);
      const tail = splitRows(xml.slice(cut), carry);
      expect([...head, ...tail]).toEqual(expected);
    }
    // Close-tag straddling the boundary: the open row is carried whole and
    // completes on the next chunk.
    const carry2 = { pending: "" };
    expect(splitRows('<row r="1"><c/></ro', carry2)).toEqual([]);
    expect(splitRows('w>', carry2)).toEqual(['<row r="1"><c/></row>']);
  });
});
