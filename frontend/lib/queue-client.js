import { QueueClient } from "@vercel/queue";

/** @typedef {import("@vercel/queue").VercelRegion} VercelRegion */

/** @type {Set<VercelRegion>} */
const VERCEL_REGIONS = new Set(/** @type {VercelRegion[]} */ ([
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
]));

/**
 * @returns {VercelRegion}
 */
export function configuredQueueRegion() {
  const requested = /** @type {VercelRegion} */ (
    (process.env.OCR2WORD_QUEUE_REGION || "sin1").trim().toLowerCase()
  );
  return VERCEL_REGIONS.has(requested) ? requested : "sin1";
}

export const conversionQueue = new QueueClient({ region: configuredQueueRegion() });
