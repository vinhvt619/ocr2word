import { spawn, spawnSync } from "node:child_process";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const port = process.env.PORT || "3000";
if (!/^\d+$/.test(port)) {
  console.error(`PORT must be numeric, received: ${port}`);
  process.exit(1);
}

const sync = spawnSync(process.execPath, ["scripts/sync-static.mjs"], {
  cwd: process.cwd(),
  env: process.env,
  stdio: "inherit",
});
if (sync.status !== 0) process.exit(sync.status ?? 1);

const nextCli = require.resolve("next/dist/bin/next");
const child = spawn(process.execPath, [nextCli, "dev", "--port", port], {
  cwd: process.cwd(),
  env: process.env,
  stdio: "inherit",
});

for (const signal of ["SIGINT", "SIGTERM"]) {
  process.on(signal, () => child.kill(signal));
}

child.on("exit", (code, signal) => {
  if (signal) process.kill(process.pid, signal);
  else process.exit(code ?? 1);
});
