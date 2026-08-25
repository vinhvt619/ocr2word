import { copyFile, mkdir, rm } from "node:fs/promises";
import { fileURLToPath } from "node:url";

const source = fileURLToPath(new URL("../../app/static/", import.meta.url));
const destination = fileURLToPath(new URL("../public/", import.meta.url));

await rm(destination, { recursive: true, force: true });
await mkdir(`${destination}/static`, { recursive: true });
await copyFile(`${source}/index.html`, `${destination}/index.html`);
await copyFile(`${source}/app.js`, `${destination}/static/app.js`);
await copyFile(`${source}/styles.css`, `${destination}/static/styles.css`);
