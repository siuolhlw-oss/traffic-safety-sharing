import fs from "node:fs";
import path from "node:path";

const root = path.resolve(process.argv[2] || ".");
const html = fs.readFileSync(path.join(root, "index.html"), "utf8");
const slideCount = (html.match(/<article class="slide/g) || []).length;
if (slideCount !== 51) throw new Error(`Expected 51 slides; found ${slideCount}`);

const localRefs = [...html.matchAll(/(?:src|href)="([^"]+)"/g)]
  .map((match) => match[1])
  .filter((ref) => !/^(?:https?:|#)/.test(ref));
const missing = localRefs.filter((ref) => !fs.existsSync(path.join(root, ref)));
if (missing.length) throw new Error(`Missing local references: ${missing.join(", ")}`);

console.log(`HTML OK: ${slideCount} slides; ${localRefs.length} local references`);
