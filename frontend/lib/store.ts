import { Redis } from "@upstash/redis";

import { CLEANUP_SET_KEY, JOB_TTL_SECONDS } from "./config";
import type { CleanupRecord, JobRecord, PdfChunkRecord } from "./types";

function redisClient(): Redis {
  const url = process.env.UPSTASH_REDIS_REST_URL || process.env.KV_REST_API_URL;
  const token = process.env.UPSTASH_REDIS_REST_TOKEN || process.env.KV_REST_API_TOKEN;
  if (!url || !token) {
    throw new Error("Redis is not configured. Connect Upstash Redis to this Vercel project.");
  }
  return new Redis({ url, token });
}

function jobKey(jobId: string): string {
  return `ocr2word:job:${jobId}`;
}

function cleanupKey(jobId: string): string {
  return `ocr2word:cleanup:${jobId}`;
}

function lockKey(jobId: string, scope: string): string {
  return `ocr2word:lock:${jobId}:${scope.replace(/[^A-Za-z0-9_-]/g, "-")}`;
}

function chunkKey(jobId: string, index: number): string {
  return `ocr2word:chunk:${jobId}:${index}`;
}

export async function getJob(jobId: string): Promise<JobRecord | null> {
  return redisClient().get<JobRecord>(jobKey(jobId));
}

export async function saveJob(job: JobRecord): Promise<void> {
  job.updated_at = new Date().toISOString();
  await redisClient().set(jobKey(job.job_id), job, { ex: JOB_TTL_SECONDS });
}

export async function updateJob(
  jobId: string,
  patch: Partial<JobRecord>,
): Promise<JobRecord | null> {
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

export async function registerCleanup(jobId: string, record: CleanupRecord): Promise<void> {
  const redis = redisClient();
  await Promise.all([
    redis.set(cleanupKey(jobId), record, { ex: JOB_TTL_SECONDS + 24 * 60 * 60 }),
    redis.sadd(CLEANUP_SET_KEY, jobId),
  ]);
}

export async function addCleanupPathnames(jobId: string, pathnames: string[]): Promise<void> {
  if (!pathnames.length) return;
  const record = await getCleanup(jobId);
  if (!record) return;
  record.pathnames = [...new Set([...record.pathnames, ...pathnames])];
  await redisClient().set(cleanupKey(jobId), record, { ex: JOB_TTL_SECONDS + 24 * 60 * 60 });
}

export async function saveChunk(chunk: PdfChunkRecord): Promise<void> {
  await redisClient().set(chunkKey(chunk.job_id, chunk.index), chunk, { ex: JOB_TTL_SECONDS });
}

export async function getChunk(jobId: string, index: number): Promise<PdfChunkRecord | null> {
  return redisClient().get<PdfChunkRecord>(chunkKey(jobId, index));
}

export async function getChunks(jobId: string, total: number): Promise<PdfChunkRecord[]> {
  const chunks = await Promise.all(
    Array.from({ length: total }, (_, index) => getChunk(jobId, index)),
  );
  return chunks.filter((chunk): chunk is PdfChunkRecord => chunk !== null);
}

export async function listCleanupIds(): Promise<string[]> {
  return redisClient().smembers(CLEANUP_SET_KEY);
}

export async function getCleanup(jobId: string): Promise<CleanupRecord | null> {
  return redisClient().get<CleanupRecord>(cleanupKey(jobId));
}

export async function removeCleanup(jobId: string): Promise<void> {
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

export async function acquireJobLock(jobId: string, token: string, scope = "job"): Promise<boolean> {
  const result = await redisClient().set(lockKey(jobId, scope), token, { nx: true, ex: 31 * 60 });
  return result === "OK";
}

export async function releaseJobLock(jobId: string, token: string, scope = "job"): Promise<void> {
  await redisClient().eval(
    "if redis.call('get', KEYS[1]) == ARGV[1] then return redis.call('del', KEYS[1]) else return 0 end",
    [lockKey(jobId, scope)],
    [token],
  );
}
