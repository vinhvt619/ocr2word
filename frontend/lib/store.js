import { Redis } from "@upstash/redis";

import { CLEANUP_SET_KEY, JOB_TTL_SECONDS } from "./config.js";

/** @typedef {import("./types.js").CleanupRecord} CleanupRecord */
/** @typedef {import("./types.js").JobRecord} JobRecord */
/** @typedef {import("./types.js").PdfChunkRecord} PdfChunkRecord */

/**
 * @returns {Redis}
 */
function redisClient() {
  const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;
  if (!url || !token) {
    throw new Error("Redis is not configured. Connect Upstash Redis to this Vercel project.");
  }
  return new Redis({ url, token });
}

/**
 * @param {string} jobId
 * @returns {string}
 */
function jobKey(jobId) {
  return `ocr2word:job:${jobId}`;
}

/**
 * @param {string} jobId
 * @returns {string}
 */
function cleanupKey(jobId) {
  return `ocr2word:cleanup:${jobId}`;
}

/**
 * @param {string} jobId
 * @param {string} scope
 * @returns {string}
 */
function lockKey(jobId, scope) {
  return `ocr2word:lock:${jobId}:${scope.replace(/[^A-Za-z0-9_-]/g, "-")}`;
}

/**
 * @param {string} jobId
 * @param {number} index
 * @returns {string}
 */
function chunkKey(jobId, index) {
  return `ocr2word:chunk:${jobId}:${index}`;
}

/**
 * @param {string} jobId
 * @returns {Promise<JobRecord | null>}
 */
export async function getJob(jobId) {
  return /** @type {Promise<JobRecord | null>} */ (redisClient().get(jobKey(jobId)));
}

/**
 * @param {JobRecord} job
 * @returns {Promise<void>}
 */
export async function saveJob(job) {
  job.updated_at = new Date().toISOString();
  await redisClient().set(jobKey(job.job_id), job, { ex: JOB_TTL_SECONDS });
}

/**
 * @param {string} jobId
 * @param {Partial<JobRecord>} patch
 * @returns {Promise<JobRecord | null>}
 */
export async function updateJob(jobId, patch) {
  const updatedAt = new Date().toISOString();
  const applied = Number(await redisClient().eval(
    `
      local raw = redis.call('GET', KEYS[1])
      if not raw then return 0 end
      local job = cjson.decode(raw)
      local patch = cjson.decode(ARGV[1])
      for key, value in pairs(patch) do
        if key == 'percent' or key == 'completed_files' or key == 'chunk_completed' then
          job[key] = math.max(tonumber(job[key] or 0), tonumber(value or 0))
        else
          job[key] = value
        end
      end
      job['updated_at'] = ARGV[2]
      redis.call('SET', KEYS[1], cjson.encode(job), 'EX', tonumber(ARGV[3]))
      return 1
    `,
    [jobKey(jobId)],
    [JSON.stringify(patch), updatedAt, JOB_TTL_SECONDS],
  ));
  return applied === 1 ? getJob(jobId) : null;
}

/**
 * @param {string} jobId
 * @param {CleanupRecord} record
 * @returns {Promise<void>}
 */
export async function registerCleanup(jobId, record) {
  const redis = redisClient();
  await Promise.all([
    redis.set(cleanupKey(jobId), record, { ex: JOB_TTL_SECONDS + 24 * 60 * 60 }),
    redis.sadd(CLEANUP_SET_KEY, jobId),
  ]);
}

/**
 * @param {string} jobId
 * @param {string[]} pathnames
 * @returns {Promise<void>}
 */
export async function addCleanupPathnames(jobId, pathnames) {
  if (!pathnames.length) return;
  const record = await getCleanup(jobId);
  if (!record) return;
  record.pathnames = [...new Set([...record.pathnames, ...pathnames])];
  await redisClient().set(cleanupKey(jobId), record, { ex: JOB_TTL_SECONDS + 24 * 60 * 60 });
}

/**
 * @param {PdfChunkRecord} chunk
 * @returns {Promise<void>}
 */
export async function saveChunk(chunk) {
  await redisClient().set(chunkKey(chunk.job_id, chunk.index), chunk, { ex: JOB_TTL_SECONDS });
}

/**
 * @param {string} jobId
 * @param {number} index
 * @returns {Promise<PdfChunkRecord | null>}
 */
export async function getChunk(jobId, index) {
  return /** @type {Promise<PdfChunkRecord | null>} */ (redisClient().get(chunkKey(jobId, index)));
}

/**
 * @param {PdfChunkRecord | null} chunk
 * @returns {chunk is PdfChunkRecord}
 */
function isPresentChunk(chunk) {
  return chunk !== null;
}

/**
 * @param {string} jobId
 * @param {number} total
 * @returns {Promise<PdfChunkRecord[]>}
 */
export async function getChunks(jobId, total) {
  const chunks = await Promise.all(
    Array.from({ length: total }, (_, index) => getChunk(jobId, index)),
  );
  return chunks.filter(isPresentChunk);
}

/**
 * @returns {Promise<string[]>}
 */
export async function listCleanupIds() {
  return redisClient().smembers(CLEANUP_SET_KEY);
}

/**
 * @param {string} jobId
 * @returns {Promise<CleanupRecord | null>}
 */
export async function getCleanup(jobId) {
  return /** @type {Promise<CleanupRecord | null>} */ (redisClient().get(cleanupKey(jobId)));
}

/**
 * @param {string} jobId
 * @returns {Promise<void>}
 */
export async function removeCleanup(jobId) {
  const redis = redisClient();
  const job = await getJob(jobId);
  const chunkKeys = Array.from(
    { length: Math.max(0, job?.chunk_total ?? 0) },
    (_, index) => chunkKey(jobId, index),
  );
  await Promise.all([
    redis.del(cleanupKey(jobId)),
    redis.del(jobKey(jobId)),
    redis.srem(CLEANUP_SET_KEY, jobId),
    chunkKeys.length ? redis.del(...chunkKeys) : Promise.resolve(0),
  ]);
}

/**
 * @param {string} jobId
 * @param {string} token
 * @param {string} [scope="job"]
 * @returns {Promise<boolean>}
 */
export async function acquireJobLock(jobId, token, scope = "job") {
  const result = await redisClient().set(lockKey(jobId, scope), token, { nx: true, ex: 31 * 60 });
  return result === "OK";
}

/**
 * @param {string} jobId
 * @param {string} token
 * @param {string} [scope="job"]
 * @returns {Promise<void>}
 */
export async function releaseJobLock(jobId, token, scope = "job") {
  await redisClient().eval(
    "if redis.call('get', KEYS[1]) == ARGV[1] then return redis.call('del', KEYS[1]) else return 0 end",
    [lockKey(jobId, scope)],
    [token],
  );
}
