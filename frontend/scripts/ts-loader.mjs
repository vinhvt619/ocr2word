import fs from "node:fs";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

export async function resolve(specifier, context, nextResolve) {
  try {
    return await nextResolve(specifier, context);
  } catch (err) {
    if (specifier.startsWith(".")) {
      const parentDir = path.dirname(fileURLToPath(context.parentURL));
      const target = path.resolve(parentDir, specifier);
      if (fs.existsSync(target + ".ts")) {
        return { url: pathToFileURL(target + ".ts").href, shortCircuit: true };
      }
    }
    throw err;
  }
}
