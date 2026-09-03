import { describe, expect, it } from "vitest";
import { CsvParseError, parseCsv } from "../lib/csv-parser";

describe("parseCsv", () => {
  it("parses a simple comma-delimited file", () => {
    expect(parseCsv("Name,Email\nJordan,jordan@example.test")).toEqual([
      ["Name", "Email"],
      ["Jordan", "jordan@example.test"],
    ]);
  });

  it("handles a quoted field containing a comma", () => {
    expect(parseCsv('Name,Role\n"Smith, Jordan",Coordinator')).toEqual([
      ["Name", "Role"],
      ["Smith, Jordan", "Coordinator"],
    ]);
  });

  it("handles an escaped quote (doubled-quote) inside a quoted field", () => {
    expect(parseCsv('Name\n"Jordan ""JJ"" Example"')).toEqual([["Name"], ['Jordan "JJ" Example']]);
  });

  it("handles CRLF line endings", () => {
    expect(parseCsv("Name,Email\r\nJordan,jordan@example.test\r\nTaylor,taylor@example.test")).toEqual([
      ["Name", "Email"],
      ["Jordan", "jordan@example.test"],
      ["Taylor", "taylor@example.test"],
    ]);
  });

  it("handles bare LF line endings", () => {
    expect(parseCsv("Name,Email\nJordan,jordan@example.test\nTaylor,taylor@example.test")).toEqual([
      ["Name", "Email"],
      ["Jordan", "jordan@example.test"],
      ["Taylor", "taylor@example.test"],
    ]);
  });

  it("strips a UTF-8 BOM at the start of the file", () => {
    const withBom = "﻿Name,Email\nJordan,jordan@example.test";
    expect(parseCsv(withBom)).toEqual([
      ["Name", "Email"],
      ["Jordan", "jordan@example.test"],
    ]);
  });

  it("skips blank lines entirely, rather than producing a one-empty-cell row", () => {
    expect(parseCsv("Name,Email\nJordan,jordan@example.test\n\nTaylor,taylor@example.test\n")).toEqual([
      ["Name", "Email"],
      ["Jordan", "jordan@example.test"],
      ["Taylor", "taylor@example.test"],
    ]);
  });

  it("preserves a genuinely empty optional cell within a populated row", () => {
    expect(parseCsv("Name,Email,Role\nMorgan,,Partner Lead")).toEqual([
      ["Name", "Email", "Role"],
      ["Morgan", "", "Partner Lead"],
    ]);
  });

  it("handles a quoted field containing an embedded newline", () => {
    expect(parseCsv('Name,Notes\nJordan,"Line one\nLine two"')).toEqual([
      ["Name", "Notes"],
      ["Jordan", "Line one\nLine two"],
    ]);
  });

  it("handles a file with no trailing newline", () => {
    expect(parseCsv("Name,Email\nJordan,jordan@example.test")).toEqual([
      ["Name", "Email"],
      ["Jordan", "jordan@example.test"],
    ]);
  });

  it("throws CsvParseError on an unterminated quoted field", () => {
    expect(() => parseCsv('Name\n"Jordan')).toThrow(CsvParseError);
    expect(() => parseCsv('Name\n"Jordan')).toThrow(/never closed/);
  });

  it("throws CsvParseError on a stray quote in the middle of an unquoted field", () => {
    expect(() => parseCsv('Name\nJor"dan')).toThrow(CsvParseError);
    expect(() => parseCsv('Name\nJor"dan')).toThrow(/middle of a field/);
  });

  it("returns an empty array for an empty file", () => {
    expect(parseCsv("")).toEqual([]);
  });
});
