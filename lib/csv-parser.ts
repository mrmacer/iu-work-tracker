// Patch 8F — a small, dependency-free, pure CSV parser used only by lib/contact-import.ts.
// Handles the RFC-4180-ish shape real spreadsheet exports produce: quoted fields, commas and
// newlines inside quotes, doubled-quote escaping, CRLF/LF line endings, a UTF-8 BOM, and blank
// lines — while surfacing a clear error for genuinely malformed quoting rather than silently
// misreading a row. No external dependency; independently testable.

export class CsvParseError extends Error {
  constructor(
    message: string,
    public readonly line: number,
  ) {
    super(message);
    this.name = "CsvParseError";
  }
}

const BOM = "﻿";

/**
 * Parses CSV text into an array of rows, each an array of cell strings. A line consisting of
 * nothing at all (an empty line between two newlines) is skipped rather than becoming a
 * single-empty-cell row — genuinely empty cells within an otherwise-populated row are preserved
 * as `""`. Throws CsvParseError on an unterminated quoted field or a stray `"` appearing
 * somewhere other than the start of a field, rather than guessing what was meant.
 */
export function parseCsv(text: string): string[][] {
  const input = text.startsWith(BOM) ? text.slice(BOM.length) : text;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let inQuotes = false;
  let line = 1;
  let i = 0;

  const pushField = () => {
    row.push(field);
    field = "";
  };
  const pushRow = () => {
    pushField();
    // A line that contributed nothing at all (just "\n" or "\r\n") is a blank line, not a
    // one-empty-cell row — skip it entirely rather than polluting the row list.
    if (!(row.length === 1 && row[0] === "")) rows.push(row);
    row = [];
  };

  while (i < input.length) {
    const char = input[i];

    if (inQuotes) {
      if (char === '"') {
        if (input[i + 1] === '"') {
          field += '"';
          i += 2;
          continue;
        }
        inQuotes = false;
        i++;
        continue;
      }
      if (char === "\n") line++;
      field += char;
      i++;
      continue;
    }

    if (char === '"') {
      if (field.length !== 0) {
        throw new CsvParseError(
          `Malformed CSV at line ${line}: a quote character appeared in the middle of a field that didn't start with one.`,
          line,
        );
      }
      inQuotes = true;
      i++;
      continue;
    }
    if (char === ",") {
      pushField();
      i++;
      continue;
    }
    if (char === "\r") {
      i++; // drop bare CR; the following \n (CRLF) or end-of-field is handled normally
      continue;
    }
    if (char === "\n") {
      pushRow();
      line++;
      i++;
      continue;
    }
    field += char;
    i++;
  }

  if (inQuotes) {
    throw new CsvParseError(`Malformed CSV: a quoted field starting before line ${line} was never closed.`, line);
  }
  // Flush a final row when the file doesn't end with a trailing newline.
  if (field !== "" || row.length > 0) pushRow();

  return rows;
}
