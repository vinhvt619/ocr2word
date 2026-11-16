import test from "node:test";
import assert from "node:assert/strict";
import { normalizeOptions } from "../lib/api";
import { outputDetails, XLSX_CONTENT_TYPE, DOCX_CONTENT_TYPE } from "../lib/config";

test("normalizeOptions - legacy jobs default to docx", () => {
  const options = normalizeOptions({});
  assert.equal(options.output_format, "docx");
  assert.equal(options.math_ocr, false);
});

test("normalizeOptions - accepts xlsx and preserves options", () => {
  const options = normalizeOptions({
    output_format: "xlsx",
    language: "vie+eng",
    dpi: 400,
  });
  assert.equal(options.output_format, "xlsx");
  assert.equal(options.language, "vie+eng");
  assert.equal(options.dpi, 400);
});

test("normalizeOptions - rejects xlsx with math_ocr", () => {
  assert.throws(() => {
    normalizeOptions({ output_format: "xlsx", math_ocr: true });
  }, /không hỗ trợ tính năng nhận diện công thức toán/);
});

test("normalizeOptions - rejects invalid output_format", () => {
  assert.throws(() => {
    normalizeOptions({ output_format: "pptx" });
  }, /chỉ hỗ trợ docx hoặc xlsx/);
});

test("outputDetails - single file for xlsx", () => {
  const output = outputDetails([{ name: "bao-cao-tai-chinh.pdf" }], "xlsx");
  assert.equal(output.name, "bao-cao-tai-chinh-editable.xlsx");
  assert.equal(output.extension, "xlsx");
  assert.equal(output.contentType, XLSX_CONTENT_TYPE);
});

test("outputDetails - multiple files for xlsx return zip", () => {
  const output = outputDetails([{ name: "file1.pdf" }, { name: "file2.png" }], "xlsx");
  assert.equal(output.name, "ocr2word-ket-qua.zip");
  assert.equal(output.extension, "zip");
  assert.equal(output.contentType, "application/zip");
});

test("outputDetails - single file for docx", () => {
  const output = outputDetails([{ name: "don-xin.pdf" }], "docx");
  assert.equal(output.name, "don-xin-editable.docx");
  assert.equal(output.contentType, DOCX_CONTENT_TYPE);
});
