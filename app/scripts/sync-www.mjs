// Copies the canonical, hand-edited web apps from the repo root into www/ so the
// Capacitor iOS bundle stays a mirror of them (single source of truth = repo root).
// Run via `npm run sync-www` (also runs automatically before cap add/sync).
import { copyFileSync, mkdirSync, existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = join(here, "..", ".."); // app/scripts -> app -> repo root
const www = join(here, "..", "www");

// [sourceFileAtRepoRoot, destNameInWww]
const files = [
  ["centerbeam_tally_recommender.html", "tally.html"],
  ["centerbeam_layout_planner.html", "planner.html"],
];

mkdirSync(www, { recursive: true });

let copied = 0;
for (const [src, dst] of files) {
  const from = join(repoRoot, src);
  if (!existsSync(from)) {
    console.error(`  ! missing source: ${src} (looked in ${repoRoot})`);
    process.exitCode = 1;
    continue;
  }
  copyFileSync(from, join(www, dst));
  console.log(`  copied ${src} -> www/${dst}`);
  copied++;
}
console.log(`sync-www: ${copied}/${files.length} file(s) copied. (www/index.html is the app shell, not synced.)`);
