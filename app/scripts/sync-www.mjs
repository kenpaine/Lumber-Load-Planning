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
// The combined app (lumber_loader.html — Tally + Loader modes with the built-in
// tally→loader handoff) is the single source of truth and the iOS entry document.
const files = [
  ["lumber_loader.html", "index.html"],
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
console.log(`sync-www: ${copied}/${files.length} file(s) copied (lumber_loader.html -> www/index.html).`);
