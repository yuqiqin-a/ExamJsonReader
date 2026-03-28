import { mkdir, rm, copyFile, writeFile, readFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const root = path.resolve(__dirname, "..");
const dist = path.join(root, "dist");
const assets = path.join(dist, "assets");

await rm(dist, { recursive: true, force: true });
await mkdir(assets, { recursive: true });

const sourceIndex = await readFile(path.join(root, "index.html"), "utf8");
const builtIndex = sourceIndex
  .replace("./src/styles.css", "./assets/styles.css")
  .replace("./src/main.js", "./assets/main.js");

await writeFile(path.join(dist, "index.html"), builtIndex, "utf8");
await copyFile(path.join(root, "src", "styles.css"), path.join(assets, "styles.css"));
await copyFile(path.join(root, "src", "main.js"), path.join(assets, "main.js"));

console.log("Build complete:", dist);
