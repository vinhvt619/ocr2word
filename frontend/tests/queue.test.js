import test from "node:test";
import assert from "node:assert/strict";
import { stageLockScope, runStage } from "../lib/queue-consumer.js";

/** @typedef {import("../lib/types.js").ConvertQueueMessage} ConvertQueueMessage */
/** @typedef {import("../lib/types.js").JobRecord} JobRecord */

test("stageLockScope - returns 'job' for xlsx output_format", () => {
  /** @type {ConvertQueueMessage} */
  const message = { job_id: "job12345678", action: "convert" };
  const xlsxJob = /** @type {JobRecord} */ (/** @type {unknown} */ ({ options: { output_format: "xlsx" } }));
  assert.equal(stageLockScope(message, xlsxJob), "job");

  /** @type {ConvertQueueMessage} */
  const chunkMsg = { job_id: "job12345678", action: "chunk", chunk_index: 2 };
  assert.equal(stageLockScope(chunkMsg, xlsxJob), "job");
});

test("stageLockScope - returns granular scope for docx legacy/chunk jobs", () => {
  const docxJob = /** @type {JobRecord} */ (/** @type {unknown} */ ({ options: { output_format: "docx" } }));
  /** @type {ConvertQueueMessage} */
  const chunkMsg = { job_id: "job12345678", action: "chunk", chunk_index: 2 };
  assert.equal(stageLockScope(chunkMsg, docxJob), "chunk-2");

  /** @type {ConvertQueueMessage} */
  const convertMsg = { job_id: "job12345678", action: "convert" };
  assert.equal(stageLockScope(convertMsg, docxJob), "convert");
});

test("runStage - rejects incompatible actions for xlsx job", async () => {
  const xlsxJob = /** @type {JobRecord} */ (/** @type {unknown} */ ({
    job_id: "job12345678",
    options: { output_format: "xlsx" },
  }));

  await assert.rejects(
    runStage({ job_id: "job12345678", action: "chunk", chunk_index: 0 }, xlsxJob),
    /Tác vụ 'chunk' không được hỗ trợ cho định dạng xuất Excel/,
  );

  await assert.rejects(
    runStage({ job_id: "job12345678", action: "merge" }, xlsxJob),
    /Tác vụ 'merge' không được hỗ trợ cho định dạng xuất Excel/,
  );
});
