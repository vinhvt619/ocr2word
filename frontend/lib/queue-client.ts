import { QueueClient, type VercelRegion } from "@vercel/queue";

const VERCEL_REGIONS = new Set<VercelRegion>([
  "arn1",
  "bom1",
  "cdg1",
  "cle1",
  "cpt1",
  "dub1",
  "dxb1",
  "fra1",
  "gru1",
  "hkg1",
  "hnd1",
  "iad1",
  "icn1",
  "kix1",
  "lhr1",
  "pdx1",
  "sfo1",
  "sin1",
  "syd1",
  "yul1",
]);

export function configuredQueueRegion(): VercelRegion {
  const requested = (process.env.OCR2WORD_QUEUE_REGION || "sin1").trim().toLowerCase() as VercelRegion;
  return VERCEL_REGIONS.has(requested) ? requested : "sin1";
}

export const conversionQueue = new QueueClient({ region: configuredQueueRegion() });
