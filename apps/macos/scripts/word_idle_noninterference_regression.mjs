import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { basename, join, resolve } from "node:path";

// Run only against an isolated diagnostic template. Never restart Word or
// close/save the user's documents to run this acceptance test.
const index = process.argv.indexOf("--template");
assert(index >= 0 && process.argv[index + 1], "Provide --template /path/to/diagnostic.dotm");
const template = resolve(process.argv[index + 1]);
assert(existsSync(template), "Diagnostic template does not exist");
const name = basename(template);
assert(name.startsWith("VisualTeXWordBuild"), "Use an isolated VisualTeXWordBuild diagnostic template");
execFileSync("python3", [new URL("verify_word_vba_source.py", import.meta.url).pathname, template, "--diagnostics"]);
const fixtureRoot = join(homedir(), "Library/Application Scripts/com.microsoft.Word/VisualTeXRuntime/Tests");
const reportPath = join(fixtureRoot, "word-idle-noninterference-result.txt");
mkdirSync(fixtureRoot, { recursive: true });
// A deterministic 200 x 100 transparent PNG with two black bars. Word cannot
// report ScaleHeight reliably for the production 1 x 1 placeholder.
writeFileSync(join(fixtureRoot, "word-idle-noninterference-fixture.png"), Buffer.from(
  "iVBORw0KGgoAAAANSUhEUgAAAMgAAABkCAYAAADDhn8LAAAAhElEQVR4nO3ZwQnAMAADsey/dLtEwfQiwS3gr88BAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAADgIo/0cSnrMdUrZT2megEAAAAAAAD82/p1Va+U9ZjqlbIeU70AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAACDiBTw5erCr5LHlAAAAAElFTkSuQmCC",
  "base64",
));
function word(lines) {
  return execFileSync("/usr/bin/osascript", ["-e", 'tell application "Microsoft Word"', ...lines.flatMap(line => ["-e", line]), "-e", "end tell"], {
    encoding: "utf8", timeout: 60_000,
  }).trim();
}
function snapshot() {
  return word([
    'set output to ""', "set documentCount to count of documents",
    "repeat with i from 1 to documentCount", "set d to document i",
    'set output to output & (name of d as text) & "|" & (saved of d as text) & "|" & (full name of d as text) & linefeed',
    "end repeat", "return output",
  ]).split(/\r?\n/).filter(Boolean).sort();
}
assert.equal(word([`return exists document ${JSON.stringify(name)}`]), "false", "The diagnostic template is already open");
const before = snapshot();
const activeName = word(['if count of documents = 0 then return ""', "return name of active document"]);
const installed = word(['if not (exists add in "VisualTeX.dotm") then return "MISSING"', 'return installed of add in "VisualTeX.dotm" as text']);
let opened = false;
rmSync(reportPath, { force: true });
try {
  if (installed === "true") word(['set installed of add in "VisualTeX.dotm" to false']);
  word([`open file name ${JSON.stringify(template)} read only false`]);
  opened = true;
  word([`activate object document ${JSON.stringify(name)}`, 'run VB macro macro name "VisualTeX_RunWordIdleNoninterferenceRegression"']);
  const report = readFileSync(reportPath, "utf8");
  assert(report.startsWith("PASS\n"), report);
  assert.match(report, /error=0:/, report);
  process.stdout.write(report + "\n");
} finally {
  if (opened) word([`if exists document ${JSON.stringify(name)} then close document ${JSON.stringify(name)} saving no`]);
  if (activeName) word([`activate object document ${JSON.stringify(activeName)}`]);
  if (installed === "true") word(['set installed of add in "VisualTeX.dotm" to true']);
  assert.deepEqual(snapshot(), before, "The user's Word document snapshot changed");
}
