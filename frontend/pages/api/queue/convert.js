import { conversionQueue } from "../../../lib/queue-client.js";
import {
  conversionCallbackOptions,
  processConversionJob,
} from "../../../lib/queue-consumer.js";

/** @typedef {import("../../../lib/types.js").ConvertQueueMessage} ConvertQueueMessage */
/** @typedef {import("@vercel/queue").MessageMetadata} MessageMetadata */

export default conversionQueue.handleNodeCallback(
  /**
   * @param {ConvertQueueMessage} message
   * @param {MessageMetadata} metadata
   */
  async (message, metadata) => processConversionJob(message, metadata.deliveryCount),
  conversionCallbackOptions,
);

export const config = {
  api: { bodyParser: false },
};
