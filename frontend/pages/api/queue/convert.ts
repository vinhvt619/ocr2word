import { conversionQueue } from "../../../lib/queue-client";
import {
  conversionCallbackOptions,
  processConversionJob,
} from "../../../lib/queue-consumer";
import type { ConvertQueueMessage } from "../../../lib/types";

export default conversionQueue.handleNodeCallback<ConvertQueueMessage>(
  async (message, metadata) => processConversionJob(message, metadata.deliveryCount),
  conversionCallbackOptions,
);

export const config = {
  api: { bodyParser: false },
};
