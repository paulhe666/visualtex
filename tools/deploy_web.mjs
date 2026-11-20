// Publishes the web edition: the `web` branch, which Cloudflare builds and serves (npm run build -> dist/,
// wrangler.jsonc), becomes the tree of origin/main as a new commit on top of it (parents: the previous `web` and
// the main commit), so the push is a fast-forward and nothing is rewritten.
//
//   node tools/deploy_web.mjs [--remote <name>] [--dry-run]
//
// It publishes what is on GitHub's main, not local work: push main first.
import { execFileSync } from "node:child_process";
import process from "node:process";

const args = process.argv.slice(2);
const remote = args.includes("--remote") ? args[args.indexOf("--remote") + 1] : "origin";
const dryRun = args.includes("--dry-run");
const git = (...command) => execFileSync("git", command, { encoding: "utf8" }).trim();

// both branches by name: a clone made with one branch does not track web
git("fetch", "--quiet", remote, `+refs/heads/main:refs/remotes/${remote}/main`, `+refs/heads/web:refs/remotes/${remote}/web`);
const main = git("rev-parse", `${remote}/main`);
const web = git("rev-parse", `${remote}/web`);
if (git("rev-parse", `${web}^{tree}`) === git("rev-parse", `${main}^{tree}`)) {
  console.log(`web already serves main ${main.slice(0, 8)}`);
  process.exit(0);
}
const subject = git("log", "-1", "--format=%s", main);
const commit = execFileSync(
  "git",
  ["commit-tree", `${main}^{tree}`, "-p", web, "-p", main, "-m", `Deploy main ${main.slice(0, 8)} to the web\n\n${subject}`],
  { encoding: "utf8" },
).trim();
if (dryRun) {
  console.log(`would push ${commit.slice(0, 8)} (tree of main ${main.slice(0, 8)}) to ${remote}/web`);
  process.exit(0);
}
git("push", remote, `${commit}:refs/heads/web`);
console.log(`web now serves main ${main.slice(0, 8)} (${commit.slice(0, 8)}); Cloudflare deploys it in about a minute`);
