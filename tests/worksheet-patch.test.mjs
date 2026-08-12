import assert from "node:assert/strict";
import { createRequire } from "node:module";
import test from "node:test";

import { createUpdatedWorkbook, patchWorksheetCell } from "../xlsx-engine.js";

const require = createRequire(import.meta.url);
globalThis.fflate = require("../vendor/fflate.js");

test("patches only the requested worksheet cell", () => {
  const original = '<?xml version="1.0"?><worksheet xmlns="main"><sheetData><row r="2"><c r="E2" t="s"><v>7</v></c><c r="F2" s="4" t="s"><v>19</v></c><c r="G2" s="5" t="s"><v>20</v></c></row></sheetData><extLst><ext uri="keep-me"/></extLst></worksheet>';
  const updated = patchWorksheetCell(original, 2, "G", "新內容 & <安全>");

  assert.ok(updated.includes('<c r="F2" s="4" t="s"><v>19</v></c>'));
  assert.ok(updated.includes('<extLst><ext uri="keep-me"/></extLst>'));
  assert.ok(updated.includes('<c r="G2" s="5" t="inlineStr"><is><t xml:space="preserve">新內容 &amp; &lt;安全&gt;</t></is></c>'));
  assert.equal(updated.replace(updated.match(/<c r="G2"[\s\S]*?<\/c>/)[0], "CELL"), original.replace(/<c r="G2"[\s\S]*?<\/c>/, "CELL"));
});

test("inserts a missing cell in column order without rewriting the row", () => {
  const original = '<worksheet><sheetData><row r="9"><c r="E9"><v>1</v></c><c r="G9"><v>2</v></c></row></sheetData></worksheet>';
  const updated = patchWorksheetCell(original, 9, "F", "續訂");

  assert.match(updated, /<c r="E9">[\s\S]*<c r="F9" t="inlineStr">[\s\S]*<c r="G9">/);
});

test("creates an updated workbook on the first output attempt", async () => {
  const path = "xl/worksheets/sheet1.xml";
  const original = '<worksheet><sheetData><row r="2"><c r="G2"><v>原內容</v></c></row></sheetData></worksheet>';
  const workbook = { files: { [path]: new TextEncoder().encode(original) } };
  const preview = {
    changes: [{ rowNumber: 2, column: "G", newReply: "新內容" }],
    targetSheet: { path, xmlText: original },
  };

  const blob = createUpdatedWorkbook(workbook, preview);
  const files = globalThis.fflate.unzipSync(new Uint8Array(await blob.arrayBuffer()));
  const updated = new TextDecoder().decode(files[path]);

  assert.match(updated, /<c r="G2" t="inlineStr">[\s\S]*新內容/);
  assert.doesNotMatch(updated, />原內容</);
});
