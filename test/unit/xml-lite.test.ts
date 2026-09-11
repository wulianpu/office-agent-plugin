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
});
