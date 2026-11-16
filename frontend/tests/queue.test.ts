import test from "node:test";
import assert from "node:assert/strict";
import { stageLockScope, runStage } from "../lib/queue-consumer";
import type { JobRecord, ConvertQueueMessage } from "../lib/types";

test("stageLockScope - returns 'job' for xlsx output_format", () => {
  const message: ConvertQueueMessage = { job_id: "job12345678", action: "convert" };
  const xlsxJob = { options: { output_format: "xlsx" } } as JobRecord;
  assert.equal(stageLockScope(message, xlsxJob), "job");

  const chunkMsg: ConvertQueueMessage = { job_id: "job12345678", action: "chunk", chunk_index: 2 };
  assert.equal(stageLockScope(chunkMsg, xlsxJob), "job");
});

test("stageLockScope - returns granular scope for docx legacy/chunk jobs", () => {
  const docxJob = { options: { output_format: "docx" } } as JobRecord;
  const chunkMsg: ConvertQueueMessage = { job_id: "job12345678", action: "chunk", chunk_index: 2 };
  assert.equal(stageLockScope(chunkMsg, docxJob), "chunk-2");

  const convertMsg: ConvertQueueMessage = { job_id: "job12345678", action: "convert" };
  assert.equal(stageLockScope(convertMsg, docxJob), "convert");
});

test("runStage - rejects incompatible actions for xlsx job", async () => {
  const xlsxJob = {
    job_id: "job12345678",
    options: { output_format: "xlsx" },
  } as JobRecord;

  await assert.rejects(
    runStage({ job_id: "job12345678", action: "chunk", chunk_index: 0 }, xlsxJob),
    /Tác vụ 'chunk' không được hỗ trợ cho định dạng xuất Excel/,
  );

  await assert.rejects(
    runStage({ job_id: "job12345678", action: "merge" }, xlsxJob),
    /Tác vụ 'merge' không được hỗ trợ cho định dạng xuất Excel/,
  );
});
